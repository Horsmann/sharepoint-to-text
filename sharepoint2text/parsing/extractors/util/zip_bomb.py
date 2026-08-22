from __future__ import annotations

import io
import zipfile
from contextlib import contextmanager
from contextvars import ContextVar, Token
from dataclasses import dataclass
from typing import Iterator

from sharepoint2text.parsing.exceptions import ExtractionZipBombError


@dataclass(frozen=True)
class ZipBombLimits:
    """
    Heuristics for rejecting probable ZIP bombs.

    These defaults are intentionally set very high to avoid false positives in
    legitimate, large SharePoint exports while still catching extreme bombs.
    """

    max_entries: int = 50_000
    max_total_uncompressed_bytes: int = 4 * 1024 * 1024 * 1024  # 4 GiB
    max_single_uncompressed_bytes: int = 1 * 1024 * 1024 * 1024  # 1 GiB
    max_total_compression_ratio: float = 200.0
    max_entry_compression_ratio: float = 500.0


DEFAULT_ZIP_BOMB_LIMITS = ZipBombLimits()


class _ZipBombChecksDisabled(ZipBombLimits):
    """Mark a scoped extraction call whose ZIP-bomb checks are disabled."""


_ZIP_BOMB_CHECKS_DISABLED = _ZipBombChecksDisabled()

# Guidance appended to every "limit exceeded" error so callers know how to
# raise the threshold for trusted files.
_LIMIT_HINT = (
    "If this file is trusted, use --zip-bomb-limit-multiplier 2..10 in the "
    "CLI (or 'none' to disable ZIP-bomb checks), "
    "or raise the relevant threshold via "
    "the zip_bomb_limits argument on sharepoint2text.read_file(), "
    "sharepoint2text.read_bytes(), or sharepoint2text.read_many() "
    "(see ZipBombLimits for the available fields)."
)

# Extraction entry points temporarily set this value while advancing their
# underlying extractor. Context-local state keeps concurrent calls isolated.
_scoped_limits: ContextVar[ZipBombLimits] = ContextVar(
    "sharepoint2text_zip_bomb_limits",
    default=DEFAULT_ZIP_BOMB_LIMITS,
)


def _validate_zip_bomb_limits(limits: ZipBombLimits | None) -> None:
    """Validate a public per-call ZIP-bomb limit value.

    Args:
        limits: Per-call limits, or ``None`` to enforce library defaults.

    Raises:
        TypeError: If ``limits`` is neither ``None`` nor ``ZipBombLimits``.
    """
    if limits is not None and not isinstance(limits, ZipBombLimits):
        raise TypeError(
            "zip_bomb_limits must be a ZipBombLimits instance or None, "
            f"got {type(limits).__name__}"
        )


@contextmanager
def _zip_bomb_limits_scope(
    limits: ZipBombLimits | None,
) -> Iterator[None]:
    """Apply ZIP-bomb limits until the current extraction step completes.

    Args:
        limits: Per-call limits, or ``None`` to enforce library defaults.

    Yields:
        Control while the selected limits are active in the current context.

    Raises:
        TypeError: If ``limits`` is neither ``None`` nor ``ZipBombLimits``.
    """
    _validate_zip_bomb_limits(limits)
    selected_limits = limits if limits is not None else DEFAULT_ZIP_BOMB_LIMITS
    token: Token[ZipBombLimits] = _scoped_limits.set(selected_limits)
    try:
        yield
    finally:
        _scoped_limits.reset(token)


def _resolve_limits(limits: ZipBombLimits | None) -> ZipBombLimits:
    """Return explicit limits or the limits scoped to the current call.

    Args:
        limits: Explicit helper limits, or ``None`` to use the call scope.

    Returns:
        The :class:`ZipBombLimits` instance to enforce for this call.
    """
    return limits if limits is not None else _scoped_limits.get()


def _limit_error(message: str, *, source: str | None) -> ExtractionZipBombError:
    """Build a limit-exceeded error annotated with source and remediation hint.

    Args:
        message: The core reason the container was rejected.
        source: Optional identifier of the calling context (for example the
            ZIP context class name), included in brackets when present.

    Returns:
        An :class:`ExtractionZipBombError` whose message points the reader to
        the public extraction call's ``zip_bomb_limits`` argument.
    """
    suffix = f" [{source}]" if source else ""
    return ExtractionZipBombError(f"{message}{suffix}. {_LIMIT_HINT}")


def _is_directory(info: zipfile.ZipInfo) -> bool:
    # ZipInfo.is_dir exists on modern Python; fall back to filename heuristic.
    is_dir = getattr(info, "is_dir", None)
    if callable(is_dir):
        return bool(is_dir())
    return info.filename.endswith("/")


def validate_zipfile(
    zf: zipfile.ZipFile,
    *,
    limits: ZipBombLimits | None = None,
    source: str | None = None,
) -> None:
    """
    Validate a ZIP container against high-confidence ZIP-bomb indicators.

    This is a best-effort DoS mitigation, not a complete sandbox.

    Args:
        zf: The open ZIP container to inspect.
        limits: Explicit limits to enforce. When ``None`` (the default), the
            current extraction call's limits are used.
        source: Optional identifier of the calling context, surfaced in error
            messages to aid debugging.

    Raises:
        ExtractionZipBombError: If any configured limit is exceeded or the
            container cannot be inspected.
    """
    limits = _resolve_limits(limits)
    if isinstance(limits, _ZipBombChecksDisabled):
        return

    try:
        infos = zf.infolist()
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
        raise ExtractionZipBombError(
            "Failed to inspect ZIP container", cause=exc
        ) from exc

    if len(infos) > limits.max_entries:
        raise _limit_error(
            f"ZIP container has too many entries ({len(infos)} > {limits.max_entries})",
            source=source,
        )

    total_uncompressed = 0
    total_compressed = 0

    for info in infos:
        if _is_directory(info):
            continue

        file_size = int(getattr(info, "file_size", 0) or 0)
        compressed_size = int(getattr(info, "compress_size", 0) or 0)

        if file_size > limits.max_single_uncompressed_bytes:
            raise _limit_error(
                f"ZIP entry too large ({file_size} bytes > {limits.max_single_uncompressed_bytes})",
                source=source,
            )

        if file_size > 0:
            if compressed_size <= 0:
                raise ExtractionZipBombError(
                    "ZIP entry has zero compressed size but non-zero uncompressed size"
                    + (f" [{source}]" if source else "")
                )
            ratio = file_size / compressed_size
            if ratio > limits.max_entry_compression_ratio:
                raise _limit_error(
                    f"ZIP entry compression ratio too high ({ratio:.1f} > {limits.max_entry_compression_ratio})",
                    source=source,
                )

        total_uncompressed += file_size
        total_compressed += compressed_size

        if total_uncompressed > limits.max_total_uncompressed_bytes:
            raise _limit_error(
                f"ZIP total uncompressed size too large ({total_uncompressed} bytes > {limits.max_total_uncompressed_bytes})",
                source=source,
            )

    if total_uncompressed > 0:
        if total_compressed <= 0:
            raise ExtractionZipBombError(
                "ZIP container has non-zero uncompressed content but zero total compressed size"
                + (f" [{source}]" if source else "")
            )
        total_ratio = total_uncompressed / total_compressed
        if total_ratio > limits.max_total_compression_ratio:
            raise _limit_error(
                f"ZIP total compression ratio too high ({total_ratio:.1f} > {limits.max_total_compression_ratio})",
                source=source,
            )


def open_zipfile(
    file_like: io.BytesIO,
    *,
    limits: ZipBombLimits | None = None,
    source: str | None = None,
) -> zipfile.ZipFile:
    """
    Open a ZIP file and validate it for ZIP-bomb indicators.

    Caller owns the returned ZipFile and must close it.

    Args:
        file_like: The ZIP container stream to open and validate.
        limits: Explicit limits to enforce. When ``None`` (the default), the
            current extraction call's limits are used.
        source: Optional identifier of the calling context for error messages.

    Returns:
        The validated, open :class:`zipfile.ZipFile`.

    Raises:
        ExtractionZipBombError: If any configured limit is exceeded.
    """
    file_like.seek(0)
    zf = zipfile.ZipFile(file_like, "r")
    try:
        validate_zipfile(zf, limits=limits, source=source)
    except (ExtractionZipBombError, zipfile.BadZipFile, OSError):
        zf.close()
        raise
    return zf


def validate_zip_bytesio(
    file_like: io.BytesIO,
    *,
    limits: ZipBombLimits | None = None,
    source: str | None = None,
) -> None:
    """
    Validate a BytesIO ZIP container without keeping it open.

    Restores the original stream position.

    Args:
        file_like: The ZIP container stream to validate.
        limits: Explicit limits to enforce. When ``None`` (the default), the
            current extraction call's limits are used.
        source: Optional identifier of the calling context for error messages.

    Raises:
        ExtractionZipBombError: If any configured limit is exceeded.
    """
    original_pos = file_like.tell()
    try:
        file_like.seek(0)
        with zipfile.ZipFile(file_like, "r") as zf:
            validate_zipfile(zf, limits=limits, source=source)
    finally:
        file_like.seek(original_pos)
