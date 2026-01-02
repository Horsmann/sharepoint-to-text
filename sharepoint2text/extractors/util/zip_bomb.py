from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

from sharepoint2text.exceptions import ExtractionZipBombError


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


def _is_directory(info: zipfile.ZipInfo) -> bool:
    # ZipInfo.is_dir exists on modern Python; fall back to filename heuristic.
    is_dir = getattr(info, "is_dir", None)
    if callable(is_dir):
        return bool(is_dir())
    return info.filename.endswith("/")


def validate_zipfile(
    zf: zipfile.ZipFile,
    *,
    limits: ZipBombLimits = DEFAULT_ZIP_BOMB_LIMITS,
    source: str | None = None,
) -> None:
    """
    Validate a ZIP container against high-confidence ZIP-bomb indicators.

    This is a best-effort DoS mitigation, not a complete sandbox.
    """
    try:
        infos = zf.infolist()
    except Exception as exc:
        raise ExtractionZipBombError(
            "Failed to inspect ZIP container", cause=exc
        ) from exc

    if len(infos) > limits.max_entries:
        raise ExtractionZipBombError(
            f"ZIP container has too many entries ({len(infos)} > {limits.max_entries})"
            + (f" [{source}]" if source else "")
        )

    total_uncompressed = 0
    total_compressed = 0

    for info in infos:
        if _is_directory(info):
            continue

        file_size = int(getattr(info, "file_size", 0) or 0)
        compressed_size = int(getattr(info, "compress_size", 0) or 0)

        if file_size > limits.max_single_uncompressed_bytes:
            raise ExtractionZipBombError(
                f"ZIP entry too large ({file_size} bytes > {limits.max_single_uncompressed_bytes})"
                + (f" [{source}]" if source else "")
            )

        if file_size > 0:
            if compressed_size <= 0:
                raise ExtractionZipBombError(
                    "ZIP entry has zero compressed size but non-zero uncompressed size"
                    + (f" [{source}]" if source else "")
                )
            ratio = file_size / compressed_size
            if ratio > limits.max_entry_compression_ratio:
                raise ExtractionZipBombError(
                    f"ZIP entry compression ratio too high ({ratio:.1f} > {limits.max_entry_compression_ratio})"
                    + (f" [{source}]" if source else "")
                )

        total_uncompressed += file_size
        total_compressed += compressed_size

        if total_uncompressed > limits.max_total_uncompressed_bytes:
            raise ExtractionZipBombError(
                f"ZIP total uncompressed size too large ({total_uncompressed} bytes > {limits.max_total_uncompressed_bytes})"
                + (f" [{source}]" if source else "")
            )

    if total_uncompressed > 0:
        if total_compressed <= 0:
            raise ExtractionZipBombError(
                "ZIP container has non-zero uncompressed content but zero total compressed size"
                + (f" [{source}]" if source else "")
            )
        total_ratio = total_uncompressed / total_compressed
        if total_ratio > limits.max_total_compression_ratio:
            raise ExtractionZipBombError(
                f"ZIP total compression ratio too high ({total_ratio:.1f} > {limits.max_total_compression_ratio})"
                + (f" [{source}]" if source else "")
            )


def open_zipfile(
    file_like: io.BytesIO,
    *,
    limits: ZipBombLimits = DEFAULT_ZIP_BOMB_LIMITS,
    source: str | None = None,
) -> zipfile.ZipFile:
    """
    Open a ZIP file and validate it for ZIP-bomb indicators.

    Caller owns the returned ZipFile and must close it.
    """
    file_like.seek(0)
    zf = zipfile.ZipFile(file_like, "r")
    try:
        validate_zipfile(zf, limits=limits, source=source)
    except Exception:
        zf.close()
        raise
    return zf


def validate_zip_bytesio(
    file_like: io.BytesIO,
    *,
    limits: ZipBombLimits = DEFAULT_ZIP_BOMB_LIMITS,
    source: str | None = None,
) -> None:
    """
    Validate a BytesIO ZIP container without keeping it open.

    Restores the original stream position.
    """
    original_pos = file_like.tell()
    try:
        file_like.seek(0)
        with zipfile.ZipFile(file_like, "r") as zf:
            validate_zipfile(zf, limits=limits, source=source)
    finally:
        file_like.seek(original_pos)
