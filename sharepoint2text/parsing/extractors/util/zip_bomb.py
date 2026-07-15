from __future__ import annotations

import io
import threading
import zipfile
from dataclasses import dataclass

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

# Guidance appended to every "limit exceeded" error so callers know how to
# raise the threshold for trusted files.
_LIMIT_HINT = (
    "If this file is trusted, raise the relevant threshold via "
    "sharepoint2text.set_zip_bomb_limits(sharepoint2text.ZipBombLimits(...)) "
    "before extraction (see ZipBombLimits for the available fields)."
)

# Process-wide active limits. Resolved at call time (not import time) so that
# sharepoint2text.set_zip_bomb_limits(...) takes effect for every subsequent
# extraction without any monkeypatching.
_active_limits: ZipBombLimits = DEFAULT_ZIP_BOMB_LIMITS
_active_limits_lock = threading.Lock()


def get_zip_bomb_limits() -> ZipBombLimits:
    """Return the ZIP-bomb limits currently in effect for extraction.

    Returns:
        The active :class:`ZipBombLimits` instance. Defaults to
        :data:`DEFAULT_ZIP_BOMB_LIMITS` until overridden via
        :func:`set_zip_bomb_limits`.
    """
    return _active_limits


def set_zip_bomb_limits(limits: ZipBombLimits) -> None:
    """Override the process-wide ZIP-bomb limits used during extraction.

    The new limits apply to every subsequent ZIP-based extraction (OOXML,
    ODF, archives, ...) that does not pass an explicit ``limits`` argument.
    Call this once during application startup.

    Args:
        limits: The replacement limits. Construct a :class:`ZipBombLimits`
            with only the fields you want to change; unspecified fields fall
            back to the library defaults.

    Raises:
        TypeError: If ``limits`` is not a :class:`ZipBombLimits` instance.

    Example:
        >>> import sharepoint2text
        >>> sharepoint2text.set_zip_bomb_limits(
        ...     sharepoint2text.ZipBombLimits(max_entry_compression_ratio=1500.0)
        ... )
    """
    if not isinstance(limits, ZipBombLimits):
        raise TypeError(
            f"limits must be a ZipBombLimits instance, got {type(limits).__name__}"
        )
    global _active_limits
    with _active_limits_lock:
        _active_limits = limits


def reset_zip_bomb_limits() -> None:
    """Restore the process-wide ZIP-bomb limits to the library defaults."""
    global _active_limits
    with _active_limits_lock:
        _active_limits = DEFAULT_ZIP_BOMB_LIMITS


def _resolve_limits(limits: ZipBombLimits | None) -> ZipBombLimits:
    """Return ``limits`` if provided, else the active process-wide limits.

    Args:
        limits: Explicit per-call limits, or ``None`` to use the active limits.

    Returns:
        The :class:`ZipBombLimits` instance to enforce for this call.
    """
    return limits if limits is not None else _active_limits


def _limit_error(message: str, *, source: str | None) -> ExtractionZipBombError:
    """Build a limit-exceeded error annotated with source and remediation hint.

    Args:
        message: The core reason the container was rejected.
        source: Optional identifier of the calling context (for example the
            ZIP context class name), included in brackets when present.

    Returns:
        An :class:`ExtractionZipBombError` whose message points the reader to
        :func:`set_zip_bomb_limits`.
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
            process-wide active limits are used (see :func:`set_zip_bomb_limits`).
        source: Optional identifier of the calling context, surfaced in error
            messages to aid debugging.

    Raises:
        ExtractionZipBombError: If any configured limit is exceeded or the
            container cannot be inspected.
    """
    limits = _resolve_limits(limits)
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
            process-wide active limits are used (see :func:`set_zip_bomb_limits`).
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
            process-wide active limits are used (see :func:`set_zip_bomb_limits`).
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
