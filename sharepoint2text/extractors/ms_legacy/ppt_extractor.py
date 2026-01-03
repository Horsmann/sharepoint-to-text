"""
PPT Presentation Extractor

Extracts text content and metadata from legacy Microsoft PowerPoint .ppt files
(PowerPoint 97-2003 binary format, OLE2/CFBF format).

Parses the "PowerPoint Document" stream within the OLE container and extracts
text from various record types according to the MS-PPT specification.
"""

import hashlib
import logging
import struct
from datetime import datetime
from typing import Any, BinaryIO, Generator, Iterator, NamedTuple

import olefile

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFileEncryptedError,
    LegacyMicrosoftParsingError,
)
from sharepoint2text.extractors.data_types import (
    PPT_TEXT_TYPE_BODY,
    PPT_TEXT_TYPE_CENTER_BODY,
    PPT_TEXT_TYPE_CENTER_TITLE,
    PPT_TEXT_TYPE_HALF_BODY,
    PPT_TEXT_TYPE_NOTES,
    PPT_TEXT_TYPE_QUARTER_BODY,
    PPT_TEXT_TYPE_TITLE,
    PptContent,
    PptImage,
    PptMetadata,
    PptSlideContent,
    PptTextBlock,
)
from sharepoint2text.extractors.util.encryption import is_ppt_encrypted
from sharepoint2text.extractors.util.image_utils import (
    BLIP_INSTANCE_JPEG_2,
    BLIP_INSTANCE_PNG_2,
    BLIP_TYPE_DIB,
    BLIP_TYPE_EMF,
    BLIP_TYPE_WMF,
    BLIP_TYPES,
    detect_image_type,
    get_image_dimensions,
    wrap_dib_as_bmp,
)

logger = logging.getLogger(__name__)

# =============================================================================
# PPT Record Type Constants (from MS-PPT specification)
# =============================================================================

# Text atom records
RT_TEXT_CHARS_ATOM = 0x0FA0  # Unicode text (UTF-16LE)
RT_TEXT_BYTES_ATOM = 0x0FA8  # ASCII/ANSI text (Latin-1)
RT_CSTRING = 0x0FBA  # Unicode string for titles

# Text context record
RT_TEXT_HEADER_ATOM = 0x0F9F  # Text type indicator

# Container record types
RT_SLIDE_CONTAINER = 0x03EE  # Individual slide
RT_NOTES_CONTAINER = 0x03F0  # Speaker notes
RT_MAIN_MASTER_CONTAINER = 0x03F8  # Master slide template

# Slide list containers
RT_SLIDE_LIST_WITH_TEXT = 0x0FF0  # Contains all slide text in order
RT_SLIDE_PERSIST_ATOM = 0x03F3  # Marks slide boundaries

# =============================================================================
# Pre-compiled struct for record header parsing (performance optimization)
# =============================================================================

_RECORD_HEADER = struct.Struct("<HHI")  # ver_instance, type, length
_RECORD_HEADER_SIZE = 8
_UINT32 = struct.Struct("<I")

# =============================================================================
# Type sets for O(1) membership tests
# =============================================================================

_TITLE_TYPES = frozenset({PPT_TEXT_TYPE_TITLE, PPT_TEXT_TYPE_CENTER_TITLE})
_BODY_TYPES = frozenset(
    {
        PPT_TEXT_TYPE_BODY,
        PPT_TEXT_TYPE_CENTER_BODY,
        PPT_TEXT_TYPE_HALF_BODY,
        PPT_TEXT_TYPE_QUARTER_BODY,
    }
)
_TEXT_RECORD_TYPES = frozenset({RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM, RT_CSTRING})

# =============================================================================
# Translation table for fast text cleaning
# =============================================================================

_CONTROL_CHARS = "".join(chr(i) for i in range(32) if i not in (9, 10))  # Keep \t, \n
_CLEAN_TRANS = str.maketrans(
    {
        "\x00": None,
        "\r": "\n",
        "\x0b": "\n",
        "\x0c": "\n",
        **{c: None for c in _CONTROL_CHARS},
    }
)

_PLACEHOLDER_PREFIXES = ("___PPT", "Click to edit")


# =============================================================================
# Record parsing structures
# =============================================================================


class Record(NamedTuple):
    """Parsed PPT record header with data."""

    rec_type: int
    rec_instance: int
    is_container: bool
    data: bytes
    offset: int
    end_offset: int


def _iter_records(data: bytes, start: int = 0) -> Iterator[Record]:
    """
    Iterate over PPT records in the data stream.

    Yields Record namedtuples with parsed header information and data slice.
    Skips malformed records with invalid lengths.
    """
    offset = start
    data_len = len(data)
    min_size = _RECORD_HEADER_SIZE

    while offset <= data_len - min_size:
        try:
            ver_instance, rec_type, rec_len = _RECORD_HEADER.unpack_from(data, offset)
        except struct.error:
            break

        # Validate record length
        if rec_len > data_len - offset - min_size:
            offset += 1
            continue

        rec_ver = ver_instance & 0x0F
        rec_instance = (ver_instance >> 4) & 0x0FFF
        is_container = rec_ver == 0x0F

        data_start = offset + min_size
        data_end = data_start + rec_len

        yield Record(
            rec_type=rec_type,
            rec_instance=rec_instance,
            is_container=is_container,
            data=data[data_start:data_end],
            offset=offset,
            end_offset=data_end,
        )

        # Containers: step into, non-containers: skip over
        offset = data_start if is_container else data_end


def _decode_text(rec_type: int, data: bytes) -> str | None:
    """Decode text from a text record based on its type."""
    try:
        if rec_type == RT_TEXT_BYTES_ATOM:
            return data.decode("latin-1")
        return data.decode("utf-16-le")  # TextCharsAtom and CString
    except (UnicodeDecodeError, ValueError):
        return None


def _make_text_block(text: str, text_type: int | None) -> PptTextBlock:
    """Create a PptTextBlock with computed type flags."""
    return PptTextBlock(
        text=text,
        text_type=text_type,
        is_title=text_type in _TITLE_TYPES,
        is_body=text_type in _BODY_TYPES,
        is_notes=text_type == PPT_TEXT_TYPE_NOTES,
    )


def _clean_text(text: str) -> str:
    """
    Clean extracted text by removing control characters and normalizing whitespace.
    Uses translation table for fast character replacement.
    """
    if not text:
        return ""

    # Fast character translation
    text = text.translate(_CLEAN_TRANS)

    # Normalize whitespace per line and filter placeholder lines
    lines = []
    for line in text.split("\n"):
        line = " ".join(line.split())
        if (
            line
            and not line.startswith(_PLACEHOLDER_PREFIXES)
            and line != "*"
            and not line.endswith("Outline Level")
        ):
            lines.append(line)

    return "\n".join(lines)


# =============================================================================
# Main Entry Point
# =============================================================================


def read_ppt(
    file_like: BinaryIO, path: str | None = None
) -> Generator[PptContent, Any, None]:
    """
    Extract text content and metadata from a legacy PowerPoint .ppt file.

    Uses a generator pattern for API consistency. PPT files yield exactly one
    PptContent object containing slides, metadata, and extracted text.
    """
    try:
        file_like.seek(0)
        if is_ppt_encrypted(file_like):
            raise ExtractionFileEncryptedError("PPT is encrypted or password-protected")

        content = _extract_ppt_content_structured(file_like)
        content.metadata.populate_from_path(path)
        yield content
    except ExtractionError:
        raise
    except Exception as exc:
        raise LegacyMicrosoftParsingError(
            "Failed to extract PPT file", cause=exc
        ) from exc


def _extract_ppt_content_structured(file_like: BinaryIO) -> PptContent:
    """Extract content from PPT file into structured PptContent object."""
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise LegacyMicrosoftParsingError(
            message="Not a valid OLE file (legacy PowerPoint format)"
        )

    file_like.seek(0)
    content = PptContent()

    with olefile.OleFileIO(file_like) as ole:
        content.streams = ole.listdir()
        content.metadata = _extract_metadata(ole)

        if not ole.exists("PowerPoint Document"):
            raise LegacyMicrosoftParsingError(
                message="No 'PowerPoint Document' stream found"
            )

        stream_data = ole.openstream("PowerPoint Document").read()
        _parse_ppt_document(stream_data, content)

        images = _extract_images_from_pictures_stream(ole)
        if images:
            _distribute_images_to_slides(content, images)

    return content


# =============================================================================
# Metadata Extraction
# =============================================================================


def _extract_metadata(ole: olefile.OleFileIO) -> PptMetadata:
    """Extract document metadata from OLE SummaryInformation stream."""
    result = PptMetadata()

    try:
        meta = ole.get_metadata()

        def decode_if_bytes(value) -> str:
            if isinstance(value, bytes):
                return value.decode("utf-8", errors="replace")
            return str(value) if value else ""

        # String fields
        for field in (
            "title",
            "subject",
            "author",
            "keywords",
            "comments",
            "last_saved_by",
            "revision_number",
            "category",
            "company",
            "manager",
            "creating_application",
        ):
            setattr(result, field, decode_if_bytes(getattr(meta, field, None)))

        # Date fields
        for src, dst in (("create_time", "created"), ("last_saved_time", "modified")):
            val = getattr(meta, src, None)
            if isinstance(val, datetime):
                setattr(result, dst, val.isoformat())

        # Numeric fields
        for src, dst in (
            ("slides", "num_slides"),
            ("notes", "num_notes"),
            ("hidden_slides", "num_hidden_slides"),
        ):
            val = getattr(meta, src, None)
            if val is not None:
                setattr(result, dst, int(val))

    except Exception as e:
        logger.debug(e)

    return result


# =============================================================================
# Document Parsing
# =============================================================================


def _parse_ppt_document(data: bytes, content: PptContent) -> None:
    """
    Parse the PowerPoint Document stream and populate PptContent.

    Uses a multi-pass approach:
    1. Extract from SlideListWithText containers (most reliable)
    2. Parse container hierarchy for additional context
    3. Raw extraction as fallback
    """
    slide_list_texts = _extract_slide_list_texts(data)
    container_texts = _parse_containers(data)

    # Build slides from SlideListWithText (primary source)
    if slide_list_texts:
        _build_slides_from_text_blocks(content, slide_list_texts)
    elif container_texts["slides"]:
        _build_slides_from_text_blocks(content, container_texts["slides"])

    # Add notes from container parsing
    if container_texts["notes"] and content.slides:
        for i, notes_texts in enumerate(container_texts["notes"]):
            if i < len(content.slides):
                slide = content.slides[i]
                existing_notes = set(slide.notes)
                for block in notes_texts:
                    if block.text not in existing_notes:
                        slide.notes.append(block.text)

    # Master slide text
    for block in container_texts.get("master", []):
        content.master_text.append(block.text)

    # Fallback: raw extraction if no text found
    if not content.all_text:
        raw_texts = _extract_all_text_raw(data)
        if raw_texts:
            content.all_text = raw_texts
            slide = PptSlideContent(slide_number=1)
            for text in raw_texts:
                slide.other_text.append(text)
                slide.all_text.append(PptTextBlock(text=text))
            content.slides.append(slide)


def _build_slides_from_text_blocks(
    content: PptContent, slides_texts: list[list[PptTextBlock]]
) -> None:
    """Build slide content from extracted text blocks."""
    for slide_num, texts in enumerate(slides_texts, 1):
        slide = PptSlideContent(slide_number=slide_num)

        for block in texts:
            slide.all_text.append(block)
            content.all_text.append(block.text)

            if block.text_type in _TITLE_TYPES or block.is_title:
                if not slide.title:
                    slide.title = block.text
                else:
                    slide.other_text.append(block.text)
            elif block.text_type in _BODY_TYPES or block.is_body:
                slide.body_text.append(block.text)
            elif block.text_type == PPT_TEXT_TYPE_NOTES or block.is_notes:
                slide.notes.append(block.text)
            else:
                slide.other_text.append(block.text)

        content.slides.append(slide)


# =============================================================================
# SlideListWithText Extraction
# =============================================================================


def _extract_slide_list_texts(data: bytes) -> list[list[PptTextBlock]]:
    """
    Extract text from SlideListWithText containers (instance 0 = slides).
    Most reliable source for slide text in presentation order.
    """
    slides_text: list[list[PptTextBlock]] = []

    for record in _iter_records(data):
        if record.rec_type == RT_SLIDE_LIST_WITH_TEXT and record.rec_instance == 0:
            slides_text.extend(_parse_slide_list_container(record.data))

    return slides_text


def _parse_slide_list_container(data: bytes) -> list[list[PptTextBlock]]:
    """Parse a SlideListWithText container to extract text organized by slide."""
    slides: list[list[PptTextBlock]] = []
    current_slide_text: list[PptTextBlock] = []
    current_text_type: int | None = None
    started = False
    any_text_blocks = False  # Track if ANY text found in entire container

    for record in _iter_records(data):
        if record.rec_type == RT_SLIDE_PERSIST_ATOM:
            # Slide boundary - save previous slide
            if started:
                # If we found text anywhere, only keep slides with text
                # Otherwise keep all slides (even empty) based on SlidePersistAtom
                if any_text_blocks:
                    if current_slide_text:
                        slides.append(current_slide_text)
                else:
                    slides.append(current_slide_text)
            started = True
            current_slide_text = []

        elif record.rec_type == RT_TEXT_HEADER_ATOM and len(record.data) >= 4:
            current_text_type = _UINT32.unpack_from(record.data)[0]

        elif record.rec_type in (RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM):
            text = _decode_text(record.rec_type, record.data)
            if text:
                text = _clean_text(text)
                if text:
                    any_text_blocks = True
                    current_slide_text.append(_make_text_block(text, current_text_type))

    # Last slide
    if started:
        if any_text_blocks:
            if current_slide_text:
                slides.append(current_slide_text)
        else:
            slides.append(current_slide_text)

    return slides


# =============================================================================
# Container Hierarchy Parsing
# =============================================================================


def _parse_containers(data: bytes) -> dict[str, list]:
    """
    Parse container hierarchy to extract text by context (slides, notes, master).
    Provides secondary extraction for content SlideListWithText might miss.
    """
    result: dict[str, list] = {"slides": [], "notes": [], "master": []}

    # Container tracking: list of (type, end_offset)
    container_stack: list[tuple[int, int]] = []
    current_slide_texts: list[PptTextBlock] = []
    current_notes_texts: list[PptTextBlock] = []
    current_master_texts: list[PptTextBlock] = []
    current_text_type: int | None = None

    in_slide = in_notes = in_master = False

    for record in _iter_records(data):
        # Pop ended containers
        while container_stack and record.offset >= container_stack[-1][1]:
            ended_type, _ = container_stack.pop()
            if ended_type == RT_SLIDE_CONTAINER:
                if current_slide_texts:
                    result["slides"].append(current_slide_texts)
                    current_slide_texts = []
                in_slide = False
            elif ended_type == RT_NOTES_CONTAINER:
                if current_notes_texts:
                    result["notes"].append(current_notes_texts)
                    current_notes_texts = []
                in_notes = False
            elif ended_type == RT_MAIN_MASTER_CONTAINER:
                in_master = False

        # Track containers
        if record.is_container:
            container_stack.append((record.rec_type, record.end_offset))
            if record.rec_type == RT_SLIDE_CONTAINER:
                in_slide, in_notes = True, False
            elif record.rec_type == RT_NOTES_CONTAINER:
                in_notes, in_slide = True, False
            elif record.rec_type == RT_MAIN_MASTER_CONTAINER:
                in_master = True

        # Extract text
        if record.rec_type == RT_TEXT_HEADER_ATOM and len(record.data) >= 4:
            current_text_type = _UINT32.unpack_from(record.data)[0]

        elif record.rec_type in _TEXT_RECORD_TYPES:
            text = _decode_text(record.rec_type, record.data)
            if text:
                text = _clean_text(text)
                if text:
                    block = _make_text_block(text, current_text_type)
                    if in_notes:
                        current_notes_texts.append(block)
                    elif in_slide:
                        current_slide_texts.append(block)
                    elif in_master:
                        current_master_texts.append(block)

    # Collect remaining text
    if current_slide_texts:
        result["slides"].append(current_slide_texts)
    if current_notes_texts:
        result["notes"].append(current_notes_texts)
    if current_master_texts:
        result["master"] = current_master_texts

    return result


# =============================================================================
# Raw Text Extraction (Fallback)
# =============================================================================


def _extract_all_text_raw(data: bytes) -> list[str]:
    """
    Extract all text from stream without structure parsing.
    Fallback when structured parsing yields no results.
    """
    texts = []

    for record in _iter_records(data):
        if record.rec_type in _TEXT_RECORD_TYPES:
            text = _decode_text(record.rec_type, record.data)
            if text:
                text = _clean_text(text)
                if text:
                    texts.append(text)

    return texts


# =============================================================================
# Image Extraction
# =============================================================================


def _distribute_images_to_slides(content: PptContent, images: list[PptImage]) -> None:
    """Distribute extracted images to slides round-robin."""
    if not images:
        return

    if not content.slides:
        content.slides.append(PptSlideContent(slide_number=1))

    num_slides = len(content.slides)
    for i, image in enumerate(images):
        slide = content.slides[i % num_slides]
        image.slide_number = slide.slide_number
        slide.images.append(image)


def _extract_images_from_pictures_stream(ole: olefile.OleFileIO) -> list[PptImage]:
    """Extract images from the Pictures stream (BLIP format)."""
    if not ole.exists("Pictures"):
        return []

    try:
        data = ole.openstream("Pictures").read()
    except Exception as e:
        logger.debug(f"Failed to read Pictures stream: {e}")
        return []

    if len(data) < 25:
        return []

    images: list[PptImage] = []
    seen_hashes: set[str] = set()
    image_index = 0

    for record in _iter_records(data):
        if record.rec_type not in BLIP_TYPES or len(record.data) <= 17:
            continue

        # BLIP header size: 17 bytes normally, 33 with secondary UID
        header_size = 17
        if record.rec_instance in (BLIP_INSTANCE_PNG_2, BLIP_INSTANCE_JPEG_2):
            header_size = 33

        if header_size >= len(record.data):
            continue

        image_data = record.data[header_size:]
        detected = detect_image_type(image_data)

        # Handle metafiles and DIB
        if detected is None:
            if record.rec_type == BLIP_TYPE_EMF:
                detected = ("emf", "image/x-emf")
            elif record.rec_type == BLIP_TYPE_WMF:
                detected = ("wmf", "image/x-wmf")
            elif record.rec_type == BLIP_TYPE_DIB:
                image_data = wrap_dib_as_bmp(image_data)
                if image_data:
                    detected = ("bmp", "image/bmp")

        if not detected or not image_data:
            continue

        # Deduplicate
        digest = hashlib.sha1(image_data).hexdigest()
        if digest in seen_hashes:
            continue
        seen_hashes.add(digest)

        image_index += 1
        width, height = get_image_dimensions(image_data, detected[0])

        images.append(
            PptImage(
                image_index=image_index,
                content_type=detected[1],
                data=image_data,
                size_bytes=len(image_data),
                width=width,
                height=height,
            )
        )

    return images


# =============================================================================
# Utility Functions
# =============================================================================


def _extract_ppt_metadata(file_like: BinaryIO) -> PptMetadata:
    """Extract only metadata from a PPT file without parsing content."""
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise LegacyMicrosoftParsingError("Not a valid OLE file")

    file_like.seek(0)

    with olefile.OleFileIO(file_like) as ole:
        return _extract_metadata(ole)
