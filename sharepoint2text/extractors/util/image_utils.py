"""
Shared Image Utilities
======================

Common image detection, parsing, and conversion utilities used by
legacy Office format extractors (PPT, XLS, DOC).

These utilities handle the OfficeArt BLIP (Binary Large Image Picture)
format used to embed images in OLE2 compound documents.
"""

import struct
from typing import Optional

# =============================================================================
# OfficeArt BLIP Record Type Constants (from MS-ODRAW specification)
# =============================================================================
# These are the recType values that identify image formats in Office streams
BLIP_TYPE_EMF = 0xF01A  # OfficeArtBlipEMF - Enhanced Metafile
BLIP_TYPE_WMF = 0xF01B  # OfficeArtBlipWMF - Windows Metafile
BLIP_TYPE_PICT = 0xF01C  # OfficeArtBlipPICT - Macintosh PICT
BLIP_TYPE_JPEG = 0xF01D  # OfficeArtBlipJPEG - JPEG image
BLIP_TYPE_PNG = 0xF01E  # OfficeArtBlipPNG - PNG image
BLIP_TYPE_DIB = 0xF01F  # OfficeArtBlipDIB - Device Independent Bitmap
BLIP_TYPE_TIFF = 0xF029  # OfficeArtBlipTIFF - TIFF image

# All valid BLIP types as a tuple for easy membership testing
BLIP_TYPES = (
    BLIP_TYPE_EMF,
    BLIP_TYPE_WMF,
    BLIP_TYPE_PICT,
    BLIP_TYPE_JPEG,
    BLIP_TYPE_PNG,
    BLIP_TYPE_DIB,
    BLIP_TYPE_TIFF,
)

# =============================================================================
# recInstance Values for BLIP Type Detection
# =============================================================================
# Used for secondary UID detection in BLIP headers
BLIP_INSTANCE_PNG = 0x6E0
BLIP_INSTANCE_PNG_2 = 0x6E1  # Has secondary UID
BLIP_INSTANCE_JPEG = 0x46A
BLIP_INSTANCE_JPEG_2 = 0x46B  # Has secondary UID or CMYK

# =============================================================================
# Image Signatures for Format Detection
# =============================================================================
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
JPEG_SIGNATURE = b"\xff\xd8\xff"
GIF_SIGNATURE = b"GIF8"
BMP_SIGNATURE = b"BM"
TIFF_LE_SIGNATURE = b"II\x2a\x00"
TIFF_BE_SIGNATURE = b"MM\x00\x2a"


def detect_image_type(data: bytes) -> tuple[str, str] | None:
    """
    Detect image type from binary data by checking file signatures.

    Args:
        data: Raw image bytes (at least first 8 bytes needed).

    Returns:
        Tuple of (extension, content_type) or None if not recognized.

    Supported formats:
        - PNG: image/png
        - JPEG: image/jpeg
        - GIF: image/gif
        - BMP: image/bmp
        - TIFF: image/tiff (both little-endian and big-endian)
    """
    if len(data) < 8:
        return None

    if data[:8] == PNG_SIGNATURE:
        return ("png", "image/png")
    if data[:3] == JPEG_SIGNATURE:
        return ("jpeg", "image/jpeg")
    if data[:4] == GIF_SIGNATURE:
        return ("gif", "image/gif")
    if data[:2] == BMP_SIGNATURE:
        return ("bmp", "image/bmp")
    if data[:4] == TIFF_LE_SIGNATURE or data[:4] == TIFF_BE_SIGNATURE:
        return ("tiff", "image/tiff")

    return None


def wrap_dib_as_bmp(dib_data: bytes) -> bytes | None:
    """
    Wrap DIB (Device Independent Bitmap) data in a BMP file header.

    DIB data is the bitmap pixel data without the BMP file header.
    This function prepends a valid BMP file header to create a
    complete BMP file.

    Args:
        dib_data: Raw DIB data starting with BITMAPINFOHEADER (40 bytes).

    Returns:
        Complete BMP file data, or None if the DIB is invalid.

    Notes:
        - Only handles BITMAPINFOHEADER format (40-byte header)
        - Supports 1, 4, 8, 16, 24, and 32 bits per pixel
        - Calculates color table size for indexed formats
    """
    if len(dib_data) < 40:
        return None

    # Check for BITMAPINFOHEADER (40 bytes)
    header_size = struct.unpack_from("<I", dib_data, 0)[0]
    if header_size != 40:
        return None

    try:
        bits_per_pixel = struct.unpack_from("<H", dib_data, 14)[0]

        if bits_per_pixel not in (1, 4, 8, 16, 24, 32):
            return None

        # Calculate color table size for indexed formats
        color_table_size = 0
        if bits_per_pixel <= 8:
            color_table_size = (1 << bits_per_pixel) * 4

        # BMP file header (14 bytes)
        file_size = 14 + len(dib_data)
        pixel_offset = 14 + header_size + color_table_size

        bmp_header = b"BM" + struct.pack("<IHHI", file_size, 0, 0, pixel_offset)
        return bmp_header + dib_data

    except struct.error:
        return None


def get_image_dimensions(
    data: bytes, image_type: str
) -> tuple[Optional[int], Optional[int]]:
    """
    Extract image dimensions from image data.

    Args:
        data: Image binary data.
        image_type: Image type extension ("png", "jpeg", "jpg", "bmp", "gif").

    Returns:
        Tuple of (width, height) or (None, None) if not extractable.

    Supported formats:
        - PNG: Reads from IHDR chunk
        - JPEG: Parses SOF markers
        - BMP: Reads from info header
        - GIF: Reads from logical screen descriptor
    """
    try:
        if image_type == "png" and len(data) >= 24:
            # PNG: IHDR chunk starts at byte 8, width at 16, height at 20
            if data[12:16] == b"IHDR":
                width = struct.unpack(">I", data[16:20])[0]
                height = struct.unpack(">I", data[20:24])[0]
                return (width, height)

        elif image_type in ("jpeg", "jpg") and len(data) >= 4:
            return get_jpeg_dimensions(data)

        elif image_type == "bmp" and len(data) >= 26:
            # BMP: Width at offset 18, height at offset 22
            if data[:2] == b"BM":
                width = struct.unpack_from("<i", data, 18)[0]
                height = abs(struct.unpack_from("<i", data, 22)[0])
                return (width, height)

        elif image_type == "gif" and len(data) >= 10:
            # GIF: Width at offset 6, height at offset 8
            width = struct.unpack_from("<H", data, 6)[0]
            height = struct.unpack_from("<H", data, 8)[0]
            return (width, height)

    except (struct.error, IndexError):
        pass

    return (None, None)


def get_jpeg_dimensions(data: bytes) -> tuple[Optional[int], Optional[int]]:
    """
    Extract dimensions from JPEG data by parsing markers.

    JPEG stores dimensions in Start of Frame (SOF) markers. This function
    scans for any SOF marker and extracts the dimensions.

    Args:
        data: JPEG binary data.

    Returns:
        Tuple of (width, height) or (None, None) if not found.

    Notes:
        - Scans for SOF0-SOF15 markers (excluding DHT, DAC, RST, SOI, EOI)
        - Handles padded marker sequences
        - Skips unknown marker segments safely
    """
    offset = 2  # Skip SOI marker

    while offset < len(data) - 9:
        if data[offset] != 0xFF:
            offset += 1
            continue

        marker = data[offset + 1]

        # Skip padding bytes
        if marker == 0xFF:
            offset += 1
            continue

        # Start of Frame markers (SOF0-SOF15, excluding DHT, DAC, RST, SOI, EOI)
        if marker in (
            0xC0,
            0xC1,
            0xC2,
            0xC3,
            0xC5,
            0xC6,
            0xC7,
            0xC9,
            0xCA,
            0xCB,
            0xCD,
            0xCE,
            0xCF,
        ):
            if offset + 9 <= len(data):
                height = struct.unpack(">H", data[offset + 5 : offset + 7])[0]
                width = struct.unpack(">H", data[offset + 7 : offset + 9])[0]
                return (width, height)

        # Skip this marker segment
        if offset + 4 <= len(data):
            segment_len = struct.unpack(">H", data[offset + 2 : offset + 4])[0]
            offset += 2 + segment_len
        else:
            break

    return (None, None)
