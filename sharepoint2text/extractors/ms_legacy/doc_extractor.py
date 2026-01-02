"""
DOC Document Extractor
======================

Extracts text content and metadata from legacy Microsoft Word .doc files
(Word 97-2003 binary format, also known as the OLE2/CFBF format).

File Format Background
----------------------
The .doc format is a complex binary format based on the Compound File Binary
Format (CFBF). The document content is stored in the "WordDocument" stream
as a series of binary structures defined by the MS-DOC specification.

Key structures in the WordDocument stream:
    - FIB (File Information Block): Header at offset 0, contains document
      metadata and pointers to other structures
    - Text: Character data stored after the FIB, may be UTF-16LE or CP1252
    - CLX (Character Property and Paragraph Property structures): Formatting
    - Data streams for embedded objects, images, etc.

Document Text Regions:
    The FIB contains character counts (ccp values) for different text regions:
    - ccpText: Main document body
    - ccpFtn: Footnotes
    - ccpHdd: Headers and footers
    - ccpAtn: Annotations/comments

    These regions are stored contiguously in the stream after the text start.

Dependencies
------------
olefile: https://github.com/decalage2/olefile
    pip install olefile

    Provides:
    - OLE compound document parsing
    - Stream enumeration and reading
    - Metadata extraction from SummaryInformation stream

Known Limitations
-----------------
- Encrypted/password-protected files raise LegacyMicrosoftParsingError
- Complex tables may not preserve structure (extracted as plain text)
- OLE objects are not extracted
- Text boxes and shapes may be missing from extraction
- Files created by very old Word versions (<97) may fail
- Complex nested structures may not parse correctly

Encoding Detection
------------------
The extractor uses heuristic detection to determine text encoding:
    1. Scans for UTF-16LE patterns (ASCII with null bytes between chars)
    2. Scans for CP1252 patterns (high-bit characters for Western European)
    3. Uses scoring system to determine most likely encoding
    4. Falls back to CP1252 if detection fails

The encoding affects the byte multiplier for text extraction:
    - UTF-16LE: 2 bytes per character
    - CP1252: 1 byte per character

Text Cleaning
-------------
Extracted text undergoes cleaning to remove/replace control characters:
    - \\x07 (cell marker) -> tab
    - \\x0b (vertical tab) -> newline
    - \\x0c (page break) -> double newline
    - \\x0d (carriage return) -> newline
    - \\x13, \\x14, \\x15 (field markers) -> removed or space
    - Various other control chars -> removed

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_legacy.doc_extractor import read_doc
    >>>
    >>> with open("document.doc", "rb") as f:
    ...     for doc in read_doc(io.BytesIO(f.read()), path="document.doc"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(f"Main text: {doc.main_text[:200]}...")
    ...         print(f"Footnotes: {doc.footnotes}")

Binary Structure Reference
--------------------------
FIB (File Information Block) key offsets:
    - 0x00: wIdent (magic number, must be 0xA5EC)
    - 0x0A: flags (bit 8 = encrypted)
    - 0x4C: ccpText (main text character count)
    - 0x50: ccpFtn (footnotes character count)
    - 0x54: ccpHdd (headers/footers character count)
    - 0x5C: ccpAtn (annotations character count)

Magic Number:
    Valid .doc files must have 0xA5EC at offset 0. Other values indicate
    either corruption or a different file type (e.g., .docx or template).

See Also
--------
- MS-DOC specification: https://docs.microsoft.com/en-us/openspecs/office_file_formats/
- ppt_extractor: For PowerPoint files
- xls_extractor: For Excel files

Maintenance Notes
-----------------
- This module was initially AI-generated and has been blackbox tested
- The text start detection algorithm may need adjustment for edge cases
- Consider adding support for embedded objects in future versions
- The _find_text_start_and_enc heuristic works for most files but may
  fail on documents with unusual character distributions
"""

import datetime
import hashlib
import io
import logging
import re
import struct
from typing import Any, Generator, List, Optional

import olefile

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFileEncryptedError,
    LegacyMicrosoftParsingError,
)
from sharepoint2text.extractors.data_types import (
    DocContent,
    DocImage,
    DocMetadata,
)

logger = logging.getLogger(__name__)

# =============================================================================
# FIB (File Information Block) Constants
# =============================================================================
# Magic numbers identifying valid .doc files
FIB_MAGIC_WORD97 = 0xA5EC  # Word 97-2003 document
FIB_MAGIC_WORD95 = 0xA5DC  # Word 95 document (also accepted)

# FIB offset for encryption flag
FIB_FLAGS_OFFSET = 0x0A
FIB_ENCRYPTED_FLAG = 0x0100  # Bit 8 set = encrypted

# FIB offsets for character counts (ccp values)
FIB_CCP_TEXT_OFFSET = 0x4C  # Main document body character count
FIB_CCP_FTN_OFFSET = 0x50  # Footnotes character count
FIB_CCP_HDD_OFFSET = 0x54  # Headers/footers character count
FIB_CCP_ATN_OFFSET = 0x5C  # Annotations character count

# Minimum valid document size and text scan range
MIN_DOC_SIZE = 0x200
TEXT_SCAN_START = 0x200
TEXT_SCAN_END = 0x2000
TEXT_SCAN_STEP = 0x40


def read_doc(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocContent, Any, None]:
    """
    Extract all relevant content from a legacy Word .doc file.

    Primary entry point for DOC file extraction. Opens the OLE container,
    parses the WordDocument stream, and extracts text content from all
    document regions (main body, footnotes, headers/footers, annotations).

    This function uses a generator pattern for API consistency with other
    extractors, even though DOC files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete DOC file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned DocContent.metadata. Useful for batch processing.

    Yields:
        DocContent: Single DocContent object containing:
            - main_text: Primary document body text
            - footnotes: Footnote text (if any)
            - headers_footers: Header and footer text (if any)
            - annotations: Comment/annotation text (if any)
            - metadata: DocMetadata with title, author, dates, counts

    Raises:
        LegacyMicrosoftParsingError: For various parsing failures:
            - Not a valid OLE file
            - No WordDocument stream found
            - File too small (<0x200 bytes)
            - Invalid magic number (not 0xA5EC)
            - Encrypted file (flag bit 8 set)

    Example:
        >>> import io
        >>> with open("report.doc", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_doc(data, path="report.doc"):
        ...         print(f"Author: {doc.metadata.author}")
        ...         print(f"Words: {doc.metadata.num_words}")
        ...         print(doc.main_text)

    Performance Notes:
        - Entire file is loaded into memory
        - OLE container is opened and closed within this function
        - Large documents may use significant memory during parsing
    """
    try:
        file_like.seek(0)
        with _DocReader(file_like) as doc:
            document = doc.read()
            document.metadata = doc.get_metadata()
            document.metadata.populate_from_path(path)

            text_len = len(document.main_text)
            logger.info(
                "Extracted DOC: %d characters, %d words",
                text_len,
                document.metadata.num_words or len(document.main_text.split()),
            )

            yield document
    except ExtractionError:
        raise
    except Exception as exc:
        raise LegacyMicrosoftParsingError(
            "Failed to extract DOC file", cause=exc
        ) from exc


class _DocReader:
    """
    Internal reader class for parsing Word .doc binary format.

    This class handles the low-level binary parsing of the WordDocument
    stream within an OLE container. It implements the context manager
    protocol for proper resource cleanup.

    The parsing process:
        1. Open OLE container via olefile
        2. Read the WordDocument stream
        3. Validate the FIB (File Information Block) header
        4. Detect text encoding (UTF-16LE or CP1252)
        5. Extract text regions based on character counts from FIB
        6. Clean extracted text of control characters

    Attributes:
        file_like: Input BytesIO containing the DOC file
        ole: OleFileIO instance for container access
        _content: Cached DocContent after first parse
        _is_unicode: Detected encoding (True=UTF-16LE, False=CP1252)
        _text_start: Detected offset where text begins in stream

    Implementation Notes:
        - Results are cached after first parse for efficiency
        - The FIB header is always at offset 0 in WordDocument stream
        - Text follows the FIB, but exact start offset varies by file
    """

    def __init__(self, file_like: io.BytesIO):
        """
        Initialize the DOC reader with file data.

        Args:
            file_like: BytesIO containing the complete DOC file.
        """
        self.file_like = file_like
        self.ole = None
        self._content: Optional[DocContent] = None
        self._is_unicode: Optional[bool] = None
        self._text_start: Optional[int] = None

    def __enter__(self):
        """Open the OLE container for reading."""
        self.ole = olefile.OleFileIO(self.file_like)
        return self

    def __exit__(self, *args):
        """Close the OLE container and release resources."""
        if self.ole:
            self.ole.close()

    def _get_stream(self, name: str) -> bytes:
        """
        Read a named stream from the OLE container.

        Args:
            name: Stream name (e.g., "WordDocument", "1Table").

        Returns:
            Raw bytes of the stream, or empty bytes if stream doesn't exist.
        """
        if self.ole and self.ole.exists(name):
            return self.ole.openstream(name).read()
        return b""

    @staticmethod
    def _build_bmp_from_dib(
        dib_data: bytes, header_size: int, color_table_size: int
    ) -> bytes:
        """Wrap DIB data in a BMP header to produce a valid BMP file.

        Assumptions/limits:
        - Input is a BITMAPINFOHEADER-style DIB (header size 40 bytes).
        - Only constructs a minimal BMP header; no validation of DIB content.
        """
        file_size = 14 + len(dib_data)
        pixel_offset = 14 + header_size + color_table_size
        bmp_header = b"BM" + struct.pack("<IHHI", file_size, 0, 0, pixel_offset)
        return bmp_header + dib_data

    @staticmethod
    def _extract_images_from_word_document(word_doc: bytes) -> List[DocImage]:
        """
        Extract embedded images from the WordDocument stream.

        This implementation targets DIB (bitmap) payloads commonly embedded in
        legacy .doc files and wraps them as BMP for consistent image handling.

        Assumptions/limits:
        - Only DIB payloads with a BITMAPINFOHEADER (40 bytes) are detected.
        - Other embedded formats (WMF/EMF/JPEG/PNG) are not parsed here.
        - Width/height are taken directly from the DIB header and may not
          reflect scaling applied in the document.
        - Duplicates are removed by content hash and low-entropy filtering.
        """
        images: List[DocImage] = []
        seen_hashes: set[str] = set()
        image_counter = 0

        i = 0
        data_len = len(word_doc)
        while i + 40 <= data_len:
            if word_doc[i : i + 4] != b"\x28\x00\x00\x00":
                i += 1
                continue

            header_size = struct.unpack_from("<I", word_doc, i)[0]
            if header_size != 40:
                i += 1
                continue

            width = struct.unpack_from("<i", word_doc, i + 4)[0]
            height = struct.unpack_from("<i", word_doc, i + 8)[0]
            planes = struct.unpack_from("<H", word_doc, i + 12)[0]
            bits_per_pixel = struct.unpack_from("<H", word_doc, i + 14)[0]
            compression = struct.unpack_from("<I", word_doc, i + 16)[0]
            size_image = struct.unpack_from("<I", word_doc, i + 20)[0]

            if (
                planes != 1
                or bits_per_pixel not in (1, 4, 8, 16, 24, 32)
                or compression not in (0, 1, 2, 3, 4, 5)
            ):
                i += 1
                continue

            abs_width = abs(width)
            abs_height = abs(height)
            if (
                abs_width == 0
                or abs_height == 0
                or abs_width > 10000
                or abs_height > 10000
            ):
                i += 1
                continue

            if size_image == 0:
                row_size = ((bits_per_pixel * abs_width + 31) // 32) * 4
                size_image = row_size * abs_height

            color_table_size = 0
            if bits_per_pixel <= 8:
                color_table_size = (1 << bits_per_pixel) * 4

            dib_len = header_size + color_table_size + size_image
            if dib_len <= 0 or i + dib_len > data_len:
                i += 1
                continue

            dib_data = word_doc[i : i + dib_len]
            bmp_data = _DocReader._build_bmp_from_dib(
                dib_data, header_size, color_table_size
            )
            digest = hashlib.sha1(bmp_data).hexdigest()
            if digest in seen_hashes:
                i += dib_len
                continue

            seen_hashes.add(digest)
            image_counter += 1
            images.append(
                DocImage(
                    image_number=image_counter,
                    content_type="image/bmp",
                    data=bmp_data,
                    size_bytes=len(bmp_data),
                    width=abs_width,
                    height=abs_height,
                )
            )
            i += dib_len

        images = _DocReader._filter_low_entropy_images(images)
        for idx, image in enumerate(images, start=1):
            image.image_number = idx
        return images

    @staticmethod
    def _extract_png_images_from_bytes(data: bytes) -> List[DocImage]:
        signature = b"\x89PNG\r\n\x1a\n"
        images: List[DocImage] = []
        seen_hashes: set[str] = set()
        image_counter = 0

        offset = 0
        while True:
            start = data.find(signature, offset)
            if start < 0:
                break

            pos = start + len(signature)
            width: Optional[int] = None
            height: Optional[int] = None

            try:
                while pos + 12 <= len(data):
                    length = struct.unpack(">I", data[pos : pos + 4])[0]
                    chunk_type = data[pos + 4 : pos + 8]
                    chunk_start = pos + 8
                    chunk_end = chunk_start + length
                    crc_end = chunk_end + 4
                    if crc_end > len(data):
                        break

                    if chunk_type == b"IHDR" and length == 13:
                        width = struct.unpack(
                            ">I", data[chunk_start : chunk_start + 4]
                        )[0]
                        height = struct.unpack(
                            ">I", data[chunk_start + 4 : chunk_start + 8]
                        )[0]

                    pos = crc_end
                    if chunk_type == b"IEND":
                        png_bytes = data[start:pos]
                        digest = hashlib.sha1(png_bytes).hexdigest()
                        if digest not in seen_hashes:
                            seen_hashes.add(digest)
                            image_counter += 1
                            images.append(
                                DocImage(
                                    image_number=image_counter,
                                    content_type="image/png",
                                    data=png_bytes,
                                    size_bytes=len(png_bytes),
                                    width=width,
                                    height=height,
                                )
                            )
                        break
            except Exception:
                pass

            offset = start + 1

        return images

    @staticmethod
    def _extract_simple_tables_from_text(text: str) -> List[List[List[str]]]:
        """Best-effort table extraction from flattened legacy DOC text."""
        tables: List[List[List[str]]] = []
        for line in (text or "").splitlines():
            tokens = [t for t in line.split() if t]
            if len(tokens) < 4 or len(tokens) % 2 != 0:
                continue
            half = len(tokens) // 2
            headers = tokens[:half]
            values = tokens[half:]
            if not headers or not values:
                continue
            if not all(re.fullmatch(r"[A-Za-z]+", h) for h in headers):
                continue
            if not all(re.fullmatch(r"\d+(?:[.,]\d+)?", v) for v in values):
                continue
            tables.append([headers, values])
        return tables

    @staticmethod
    def _filter_low_entropy_images(images: List[DocImage]) -> List[DocImage]:
        """Filter likely mask/placeholder bitmaps when duplicates exist.

        Assumptions/limits:
        - Heuristic based on byte diversity in the pixel payload.
        - Only applied within groups of identical size/format.
        - May drop valid low-color images or keep some masks.
        """
        if len(images) < 2:
            return images

        grouped: dict[tuple[Optional[int], Optional[int], int, str], List[DocImage]] = (
            {}
        )
        for image in images:
            key = (image.width, image.height, image.size_bytes, image.content_type)
            grouped.setdefault(key, []).append(image)

        filtered: List[DocImage] = []
        for group in grouped.values():
            if len(group) == 1:
                filtered.extend(group)
                continue

            diversities = []
            for image in group:
                data = image.data
                payload = (
                    data[54:] if data.startswith(b"BM") and len(data) > 54 else data
                )
                sample = payload[:100000]
                diversities.append(len(set(sample)))

            max_diversity = max(diversities) if diversities else 0
            if max_diversity <= 8:
                filtered.extend(group)
                continue

            threshold = max_diversity * 0.6
            for image, diversity in zip(group, diversities):
                if diversity >= threshold:
                    filtered.append(image)

        return filtered

    @staticmethod
    def _extract_image_captions(text: str) -> List[str]:
        """Extract caption-like lines from document text for image mapping.

        Assumptions/limits:
        - Relies on SEQ field markers to identify captions across languages.
        - Captions are assigned to images in document order without anchors.
        - If no SEQ markers are present, returns an empty list.
        """
        if not text:
            return []

        seq_field_pattern = re.compile(
            r"\bSEQ\b\s+\"?[^\"\\]+\"?\s+(?:\\\*\w+\s+)?",
            re.IGNORECASE,
        )

        captions: List[str] = []
        seen: set[str] = set()
        for line in text.splitlines():
            cleaned = line.strip()
            if not cleaned:
                continue
            cleaned, count = seq_field_pattern.subn("", cleaned)
            if count == 0:
                continue
            cleaned = cleaned.strip()
            cleaned = re.sub(r"\s+", " ", cleaned)
            if cleaned and cleaned not in seen:
                seen.add(cleaned)
                captions.append(cleaned)
        return captions

    def _parse_content(self) -> DocContent:
        """
        Parse the WordDocument stream and extract all text content.

        This is the core parsing method that reads the binary structure,
        validates the file format, and extracts text from all regions.

        Returns:
            DocContent: Populated dataclass with main_text, footnotes,
                headers_footers, and annotations.

        Raises:
            LegacyMicrosoftParsingError: For parsing failures including:
                - File not opened (ole is None)
                - Missing WordDocument stream
                - File too small (<0x200 bytes)
                - Invalid magic number
                - Encrypted file

        Implementation Details:
            1. Reads character counts from FIB offsets 0x4C-0x5C
            2. Detects encoding via _find_text_start_and_enc()
            3. Calculates byte multiplier (2 for Unicode, 1 for CP1252)
            4. Extracts each text region sequentially
            5. Decodes and cleans text for each region
        """
        if self._content is not None:
            return self._content

        if not self.ole:
            raise LegacyMicrosoftParsingError("File not opened")

        word_doc = self._get_stream("WordDocument")
        if not word_doc:
            raise LegacyMicrosoftParsingError("No WordDocument Stream")

        if len(word_doc) < MIN_DOC_SIZE:
            raise LegacyMicrosoftParsingError("File too small")

        # Validate magic number
        magic = struct.unpack_from("<H", word_doc, 0)[0]
        if magic not in (FIB_MAGIC_WORD97, FIB_MAGIC_WORD95):
            raise LegacyMicrosoftParsingError(
                f"Not a valid .doc file (Magic: {hex(magic)})"
            )

        # Check encryption flag
        flags = struct.unpack_from("<H", word_doc, FIB_FLAGS_OFFSET)[0]
        if flags & FIB_ENCRYPTED_FLAG:
            raise ExtractionFileEncryptedError("DOC is encrypted or password-protected")

        # Extract character counts from FIB
        ccp_text = struct.unpack_from("<I", word_doc, FIB_CCP_TEXT_OFFSET)[0]
        ccp_ftn = struct.unpack_from("<I", word_doc, FIB_CCP_FTN_OFFSET)[0]
        ccp_hdd = struct.unpack_from("<I", word_doc, FIB_CCP_HDD_OFFSET)[0]
        ccp_atn = struct.unpack_from("<I", word_doc, FIB_CCP_ATN_OFFSET)[0]

        self._text_start, self._is_unicode = self._find_text_start_and_enc(word_doc)

        # Byte multiplier: 2 for UTF-16LE, 1 for CP1252
        mult = 2 if self._is_unicode else 1
        encoding = "utf-16-le" if self._is_unicode else "cp1252"

        pos = self._text_start

        # Main text
        main_data = word_doc[pos : pos + ccp_text * mult]
        pos += ccp_text * mult

        # Footnotes
        ftn_data = word_doc[pos : pos + ccp_ftn * mult] if ccp_ftn > 0 else b""
        pos += ccp_ftn * mult

        # Headers/Footers
        hdd_data = word_doc[pos : pos + ccp_hdd * mult] if ccp_hdd > 0 else b""
        pos += ccp_hdd * mult

        # Annotations
        atn_data = word_doc[pos : pos + ccp_atn * mult] if ccp_atn > 0 else b""

        images = self._extract_images_from_word_document(word_doc)
        table_stream = self._get_stream("1Table") or self._get_stream("0Table")
        images.extend(self._extract_png_images_from_bytes(word_doc))
        if table_stream:
            images.extend(self._extract_png_images_from_bytes(table_stream))

        # De-duplicate across sources and normalize indices.
        dedup: List[DocImage] = []
        seen_hashes: set[str] = set()
        for image in images:
            digest = hashlib.sha1(image.data).hexdigest() if image.data else ""
            if digest and digest in seen_hashes:
                continue
            if digest:
                seen_hashes.add(digest)
            dedup.append(image)
        images = dedup
        for idx, image in enumerate(images, start=1):
            image.image_number = idx

        main_text = self._clean_text(main_data.decode(encoding, errors="replace"))
        footnotes_text = (
            self._clean_text(ftn_data.decode(encoding, errors="replace"))
            if ftn_data
            else ""
        )
        headers_text = (
            self._clean_text(hdd_data.decode(encoding, errors="replace"))
            if hdd_data
            else ""
        )
        annotations_text = (
            self._clean_text(atn_data.decode(encoding, errors="replace"))
            if atn_data
            else ""
        )
        caption_source = "\n".join(
            part
            for part in [main_text, footnotes_text, headers_text, annotations_text]
            if part
        )
        captions = self._extract_image_captions(caption_source)
        for idx, image in enumerate(images):
            if idx < len(captions):
                image.caption = captions[idx]

        tables = self._extract_simple_tables_from_text(main_text)

        self._content = DocContent(
            main_text=main_text,
            footnotes=footnotes_text,
            headers_footers=headers_text,
            annotations=annotations_text,
            images=images,
            tables=tables,
        )

        return self._content

    def read(self) -> DocContent:
        """
        Extract and return all text content from the document.

        This is the primary extraction method that triggers parsing
        (if not already done) and returns the complete content.

        Returns:
            DocContent: Dataclass with all extracted text regions.
        """
        return self._parse_content()

    def get_main_text(self) -> str:
        """Get only the main document body text."""
        return self._parse_content().main_text

    def get_headers_footers(self) -> str:
        """Get only the header and footer text."""
        return self._parse_content().headers_footers

    def get_footnotes(self) -> str:
        """Get only the footnote text."""
        return self._parse_content().footnotes

    def get_annotations(self) -> str:
        """Get only the annotation/comment text."""
        return self._parse_content().annotations

    def get_all_parts(self) -> DocContent:
        """Get all document parts as a DocContent dataclass."""
        return self._parse_content()

    @staticmethod
    def _find_text_start_and_enc(word_doc: bytes) -> tuple:
        """
        Detect the text start offset and character encoding in WordDocument stream.

        The text in a .doc file starts somewhere after the FIB header (which
        ends around offset 0x200). This method scans the stream to find where
        readable text begins and whether it's encoded as UTF-16LE or CP1252.

        Detection Algorithm:
            1. Scan in 64-byte chunks starting at offset 0x200
            2. For each chunk, score both UTF-16LE and CP1252 patterns
            3. UTF-16LE: Look for ASCII chars followed by null bytes
            4. CP1252: Look for printable ASCII and high-bit Western chars
            5. Return when a clear pattern emerges (score threshold met)
            6. Fall back to offset 0x800, CP1252 if no pattern found

        Args:
            word_doc: Raw bytes of the WordDocument stream.

        Returns:
            Tuple of (offset, is_unicode):
                - offset: Byte offset where text content begins
                - is_unicode: True if UTF-16LE, False if CP1252

        Implementation Notes:
            - UTF-16LE score threshold: >20 matching pairs in 64 bytes
            - CP1252 score threshold: >45 printable chars in 64 bytes
            - German umlauts (ä, ö, ü, etc.) are explicitly checked for UTF-16LE
            - Scan stops at min(file_length - 64, 0x2000) for safety
        """
        scan_end = min(len(word_doc) - 64, TEXT_SCAN_END)
        for offset in range(TEXT_SCAN_START, scan_end, TEXT_SCAN_STEP):
            sample = word_doc[offset : offset + 64]

            utf16_score = 0
            for i in range(0, min(len(sample) - 1, 60), 2):
                b1, b2 = sample[i], sample[i + 1]
                if (0x20 <= b1 <= 0x7E or b1 in (0x0D, 0x0A)) and b2 == 0x00:
                    utf16_score += 1
                elif b1 in (0xE4, 0xF6, 0xFC, 0xC4, 0xD6, 0xDC, 0xDF) and b2 == 0x00:
                    utf16_score += 1

            cp1252_score = sum(
                1
                for b in sample
                if (0x20 <= b <= 0x7E) or b in (0x0D, 0x0A, 0x09) or (0xC0 <= b <= 0xFF)
            )

            if utf16_score > 20:
                return offset, True
            elif cp1252_score > 45:
                return offset, False

        # Fallback: assume text starts at 0x800, CP1252 encoding
        return 0x800, False

    @staticmethod
    def _clean_text(text: str) -> str:
        """
        Clean extracted text by replacing/removing control characters.

        Word documents contain various control characters that mark
        formatting, field boundaries, and structure. This method
        normalizes these to produce readable plain text.

        Character Mappings:
            - \\x07 (cell marker) -> tab
            - \\x0b (vertical tab) -> newline
            - \\x0c (page break) -> double newline
            - \\x0d (carriage return) -> newline
            - \\x13 (field begin) -> removed
            - \\x14 (field separator) -> space
            - \\x15 (field end) -> removed
            - \\x01, \\x08, \\x19, \\x1e, \\x1f -> removed
            - \\xa0 (non-breaking space) -> regular space

        Whitespace Normalization:
            - Multiple spaces/tabs collapsed to single space
            - Three or more newlines collapsed to double newline

        Args:
            text: Raw text extracted from document.

        Returns:
            Cleaned, normalized text string.
        """
        if not text:
            return ""

        replacements = {
            "\x07": "\t",
            "\x0b": "\n",
            "\x0c": "\n\n",
            "\x0d": "\n",
            "\x13": "",
            "\x14": " ",
            "\x15": "",
            "\x01": "",
            "\x08": "",
            "\x19": "",
            "\x1e": "",
            "\x1f": "",
            "\xa0": " ",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)

        text = re.sub(r"[\x00-\x08\x0e-\x1f\x7f]", "", text)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def get_metadata(self) -> DocMetadata:
        """
        Extract document metadata from OLE SummaryInformation stream.

        Uses olefile's built-in metadata extraction to read the standard
        document properties stored in the OLE container.

        Returns:
            DocMetadata: Populated metadata including:
                - title, author, subject, keywords
                - last_saved_by, create_time, last_saved_time
                - num_pages, num_words, num_chars

        Notes:
            - Returns empty DocMetadata if OLE is not open or extraction fails
            - Dates are converted to ISO format strings
            - Text fields are decoded from bytes as UTF-8
            - Failures are logged at debug level but don't raise exceptions
        """
        if not self.ole:
            return DocMetadata()
        try:
            m = self.ole.get_metadata()
            return DocMetadata(
                title=m.title.decode("utf-8"),
                author=m.author.decode("utf-8"),
                subject=m.subject.decode("utf-8"),
                keywords=m.keywords.decode("utf-8"),
                last_saved_by=m.last_saved_by.decode("utf-8"),
                create_time=(
                    m.create_time.isoformat()
                    if isinstance(m.create_time, datetime.datetime)
                    else ""
                ),
                last_saved_time=(
                    m.last_saved_time.isoformat()
                    if isinstance(m.last_saved_time, datetime.datetime)
                    else ""
                ),
                num_pages=m.num_pages,
                num_words=m.num_words,
                num_chars=m.num_chars,
            )
        except Exception as e:
            logger.debug(f"Metadata extraction failed: [{e}]")
            return DocMetadata()
