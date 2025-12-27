"""
PPT Text Extractor - Extract text and metadata from legacy PowerPoint (.ppt) files.

This module uses olefile to parse the OLE2 compound document structure
and extracts text content from the PowerPoint Document stream.

Based on MS-PPT specification:
- TextCharsAtom (0x0FA0): Unicode text
- TextBytesAtom (0x0FA8): ASCII/ANSI text
"""

import logging
import struct
from datetime import datetime
from typing import Any, BinaryIO, Generator

import olefile

from sharepoint2text.extractors.data_types import (
    PPT_TEXT_TYPE_BODY,
    PPT_TEXT_TYPE_CENTER_BODY,
    PPT_TEXT_TYPE_CENTER_TITLE,
    PPT_TEXT_TYPE_HALF_BODY,
    PPT_TEXT_TYPE_NOTES,
    PPT_TEXT_TYPE_QUARTER_BODY,
    PPT_TEXT_TYPE_TITLE,
    PptContent,
    PptMetadata,
    PptSlideContent,
    PptTextBlock,
)

logger = logging.getLogger(__name__)

# PPT Record Types for text extraction
RT_TEXT_CHARS_ATOM = 0x0FA0  # Unicode text (UTF-16LE)
RT_TEXT_BYTES_ATOM = 0x0FA8  # ASCII/ANSI text
RT_CSTRING = 0x0FBA  # Unicode string (titles, etc.)

# Text header types (indicates what kind of text follows)
RT_TEXT_HEADER_ATOM = 0x0F9F

# Container record types
RT_DOCUMENT_CONTAINER = 0x03E8
RT_SLIDE_CONTAINER = 0x03EE
RT_NOTES_CONTAINER = 0x03F0
RT_MAIN_MASTER_CONTAINER = 0x03F8
RT_HANDOUT_CONTAINER = 0x0FC9

# Slide list containers
RT_SLIDE_LIST_WITH_TEXT = 0x0FF0
RT_SLIDE_PERSIST_ATOM = 0x03F3

# Drawing containers (may contain text)
RT_PP_DRAWING = 0x040C
RT_PP_DRAWING_GROUP = 0x040B

# Text-related containers
RT_TEXT_SPEC_INFO_ATOM = 0x0FAA
RT_STYLE_TEXT_PROP_ATOM = 0x0FA1
RT_TEXT_RULES_ATOM = 0x0F98
RT_TEXT_INTERACTIVE_INFO_ATOM = 0x0FDF

# Outline text
RT_OUTLINE_TEXT_REF_ATOM = 0x0F9E


def read_ppt(
    file_like: BinaryIO, path: str | None = None
) -> Generator[PptContent, Any, None]:
    """
    Extract text content and metadata from a legacy PowerPoint (.ppt) file.

    Args:
        file_like: A file-like object (e.g., io.BytesIO) containing the PPT file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        PPTContent dataclass containing:
            - metadata: PPTMetadata dataclass with document metadata
            - slides: List of SlideContent objects
            - master_text: Text from master slides
            - all_text: All text in order
            - streams: List of stream paths

    Raises:
        ValueError: If the file is not a valid OLE file or PPT document.
        IOError: If there's an error reading the file.
    """
    file_like.seek(0)
    content = _extract_ppt_content_structured(file_like)
    content.metadata.populate_from_path(path)
    yield content


def _extract_ppt_content_structured(file_like: BinaryIO) -> PptContent:
    """
    Extract content as structured PPTContent object.

    Args:
        file_like: A file-like object containing the PPT file data.

    Returns:
        PPTContent object with all extracted data.
    """
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise ValueError("Not a valid OLE file (legacy PowerPoint format)")

    file_like.seek(0)

    content = PptContent()

    with olefile.OleFileIO(file_like) as ole:
        content.streams = ole.listdir()
        content.metadata = _extract_metadata(ole)

        if ole.exists("PowerPoint Document"):
            ppt_stream = ole.openstream("PowerPoint Document")
            stream_data = ppt_stream.read()

            # Parse the document structure
            _parse_ppt_document(stream_data, content)
        else:
            raise ValueError(
                "No 'PowerPoint Document' stream found - may not be a valid PPT file"
            )

    return content


def _extract_metadata(ole: olefile.OleFileIO) -> PptMetadata:
    """Extract metadata from the OLE file."""
    result = PptMetadata()

    try:
        meta = ole.get_metadata()

        def decode_if_bytes(value) -> str:
            if isinstance(value, bytes):
                return value.decode("utf-8", errors="replace")
            return str(value) if value else ""

        result.title = decode_if_bytes(getattr(meta, "title", None))
        result.subject = decode_if_bytes(getattr(meta, "subject", None))
        result.author = decode_if_bytes(getattr(meta, "author", None))
        result.keywords = decode_if_bytes(getattr(meta, "keywords", None))
        result.comments = decode_if_bytes(getattr(meta, "comments", None))
        result.last_saved_by = decode_if_bytes(getattr(meta, "last_saved_by", None))
        result.revision_number = decode_if_bytes(getattr(meta, "revision_number", None))
        result.category = decode_if_bytes(getattr(meta, "category", None))
        result.company = decode_if_bytes(getattr(meta, "company", None))
        result.manager = decode_if_bytes(getattr(meta, "manager", None))
        result.creating_application = decode_if_bytes(
            getattr(meta, "creating_application", None)
        )

        create_time = getattr(meta, "create_time", None)
        if isinstance(create_time, datetime):
            result.created = create_time.isoformat()

        last_saved_time = getattr(meta, "last_saved_time", None)
        if isinstance(last_saved_time, datetime):
            result.modified = last_saved_time.isoformat()

        slides = getattr(meta, "slides", None)
        if slides is not None:
            result.num_slides = int(slides)

        notes = getattr(meta, "notes", None)
        if notes is not None:
            result.num_notes = int(notes)

        hidden_slides = getattr(meta, "hidden_slides", None)
        if hidden_slides is not None:
            result.num_hidden_slides = int(hidden_slides)

    except Exception as e:
        logger.debug(e)

    return result


def _parse_current_user(data: bytes) -> dict[str, Any] | None:
    """Parse the Current User stream for additional info."""
    if len(data) < 20:
        return None

    try:
        if len(data) >= 24:
            username_len = struct.unpack("<I", data[20:24])[0]
            if len(data) >= 24 + username_len and 0 < username_len < 256:
                username_offset = 24
                if len(data) >= username_offset + username_len:
                    try:
                        ansi_name = data[
                            username_offset : username_offset + username_len
                        ]
                        if b"\x00" in ansi_name:
                            ansi_name = ansi_name[: ansi_name.index(b"\x00")]
                        current_user = ansi_name.decode("latin-1").strip()
                        if current_user:
                            return {"username": current_user}
                    except Exception as e:
                        logger.exception(e)
                        pass
    except Exception as e:
        logger.exception(e)
        pass

    return None


def _parse_ppt_document(data: bytes, content: PptContent) -> None:
    """
    Parse the PowerPoint Document stream and extract content.

    Uses a state machine to track container hierarchy and associate
    text with the correct slides.
    """
    # First pass: Find all SlideListWithText containers
    # These contain the actual slide text in presentation order
    slide_list_texts = _extract_slide_list_texts(data)

    # Second pass: Parse container structure for additional context
    container_texts = _parse_containers(data)

    # Build slides from SlideListWithText (primary source)
    if slide_list_texts:
        for slide_num, texts in enumerate(slide_list_texts, 1):
            slide = PptSlideContent(slide_number=slide_num)

            for text_block in texts:
                slide.all_text.append(text_block)
                content.all_text.append(text_block.text)

                if text_block.text_type in (
                    PPT_TEXT_TYPE_TITLE,
                    PPT_TEXT_TYPE_CENTER_TITLE,
                ):
                    if not slide.title:
                        slide.title = text_block.text
                    else:
                        slide.other_text.append(text_block.text)
                elif (
                    text_block.text_type == PPT_TEXT_TYPE_BODY
                    or text_block.text_type
                    in (
                        PPT_TEXT_TYPE_CENTER_BODY,
                        PPT_TEXT_TYPE_HALF_BODY,
                        PPT_TEXT_TYPE_QUARTER_BODY,
                    )
                ):
                    slide.body_text.append(text_block.text)
                elif text_block.text_type == PPT_TEXT_TYPE_NOTES:
                    slide.notes.append(text_block.text)
                else:
                    slide.other_text.append(text_block.text)

            content.slides.append(slide)

    # If no SlideListWithText found, fall back to container parsing
    elif container_texts["slides"]:
        for slide_num, texts in enumerate(container_texts["slides"], 1):
            slide = PptSlideContent(slide_number=slide_num)

            for text_block in texts:
                slide.all_text.append(text_block)
                content.all_text.append(text_block.text)

                if text_block.is_title or text_block.text_type == PPT_TEXT_TYPE_TITLE:
                    if not slide.title:
                        slide.title = text_block.text
                    else:
                        slide.other_text.append(text_block.text)
                elif text_block.is_body:
                    slide.body_text.append(text_block.text)
                elif text_block.is_notes:
                    slide.notes.append(text_block.text)
                else:
                    slide.other_text.append(text_block.text)

            content.slides.append(slide)

    # Add notes from container parsing if we have slides
    if container_texts["notes"] and content.slides:
        for i, notes_texts in enumerate(container_texts["notes"]):
            if i < len(content.slides):
                for text_block in notes_texts:
                    if text_block.text not in content.slides[i].notes:
                        content.slides[i].notes.append(text_block.text)

    # Master slide text
    for text_block in container_texts.get("master", []):
        content.master_text.append(text_block.text)

    # If still no text found, do a raw extraction
    if not content.all_text:
        raw_texts = _extract_all_text_raw(data)
        content.all_text = raw_texts

        # Create a single slide with all text if we found any
        if raw_texts:
            slide = PptSlideContent(slide_number=1)
            for text in raw_texts:
                slide.other_text.append(text)
                slide.all_text.append(PptTextBlock(text=text))
            content.slides.append(slide)


def _extract_slide_list_texts(data: bytes) -> list[list[PptTextBlock]]:
    """
    Extract text from SlideListWithText containers.

    SlideListWithText (0x0FF0) is the most reliable source for slide text
    as it contains text in presentation order with proper slide boundaries.
    """
    slides_text: list[list[PptTextBlock]] = []
    offset = 0

    while offset < len(data) - 8:
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            break

        if rec_len > len(data) - offset - 8:
            offset += 1
            continue

        rec_ver = rec_ver_instance & 0x0F
        rec_instance = (rec_ver_instance >> 4) & 0x0FFF
        is_container = rec_ver == 0x0F

        if rec_type == RT_SLIDE_LIST_WITH_TEXT:
            # Found SlideListWithText container
            # rec_instance indicates type: 0=slides, 1=master, 2=notes
            container_data = data[offset + 8 : offset + 8 + rec_len]

            if rec_instance == 0:  # Slides
                slide_texts = _parse_slide_list_container(container_data)
                slides_text.extend(slide_texts)

        if is_container:
            offset += 8
        else:
            offset += 8 + rec_len

    return slides_text


def _parse_slide_list_container(data: bytes) -> list[list[PptTextBlock]]:
    """Parse a SlideListWithText container to extract per-slide text."""
    slides: list[list[PptTextBlock]] = []
    current_slide_text: list[PptTextBlock] = []
    current_text_type: int | None = None

    offset = 0

    while offset < len(data) - 8:
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            break

        if rec_len > len(data) - offset - 8:
            offset += 1
            continue

        rec_ver = rec_ver_instance & 0x0F
        is_container = rec_ver == 0x0F

        record_data = data[offset + 8 : offset + 8 + rec_len]

        if rec_type == RT_SLIDE_PERSIST_ATOM:
            # New slide boundary
            if current_slide_text:
                slides.append(current_slide_text)
                current_slide_text = []

        elif rec_type == RT_TEXT_HEADER_ATOM:
            # Text type indicator
            if rec_len >= 4:
                current_text_type = struct.unpack("<I", record_data[:4])[0]

        elif rec_type == RT_TEXT_CHARS_ATOM:
            try:
                text = record_data.decode("utf-16-le")
                text = _clean_text(text)
                if text:
                    block = PptTextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (PPT_TEXT_TYPE_TITLE, PPT_TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            PPT_TEXT_TYPE_BODY,
                            PPT_TEXT_TYPE_CENTER_BODY,
                            PPT_TEXT_TYPE_HALF_BODY,
                            PPT_TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == PPT_TEXT_TYPE_NOTES,
                    )
                    current_slide_text.append(block)
            except UnicodeDecodeError:
                pass

        elif rec_type == RT_TEXT_BYTES_ATOM:
            try:
                text = record_data.decode("latin-1")
                text = _clean_text(text)
                if text:
                    block = PptTextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (PPT_TEXT_TYPE_TITLE, PPT_TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            PPT_TEXT_TYPE_BODY,
                            PPT_TEXT_TYPE_CENTER_BODY,
                            PPT_TEXT_TYPE_HALF_BODY,
                            PPT_TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == PPT_TEXT_TYPE_NOTES,
                    )
                    current_slide_text.append(block)
            except UnicodeDecodeError:
                pass

        if is_container:
            offset += 8
        else:
            offset += 8 + rec_len

    # Don't forget the last slide
    if current_slide_text:
        slides.append(current_slide_text)

    return slides


def _parse_containers(data: bytes) -> dict[str, list]:
    """
    Parse container structure to extract text organized by container type.
    """
    result = {
        "slides": [],
        "notes": [],
        "master": [],
    }

    # Track container stack
    container_stack: list[tuple[int, int, int]] = []  # (type, start_offset, end_offset)
    current_slide_texts: list[PptTextBlock] = []
    current_notes_texts: list[PptTextBlock] = []
    current_master_texts: list[PptTextBlock] = []
    current_text_type: int | None = None

    in_slide = False
    in_notes = False
    in_master = False

    offset = 0

    while offset < len(data) - 8:
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            break

        if rec_len > len(data) - offset - 8:
            offset += 1
            continue

        rec_ver = rec_ver_instance & 0x0F
        is_container = rec_ver == 0x0F

        record_data = data[offset + 8 : offset + 8 + rec_len]

        # Update container stack
        while container_stack and offset >= container_stack[-1][2]:
            ended_container = container_stack.pop()
            if ended_container[0] == RT_SLIDE_CONTAINER:
                if current_slide_texts:
                    result["slides"].append(current_slide_texts)
                    current_slide_texts = []
                in_slide = False
            elif ended_container[0] == RT_NOTES_CONTAINER:
                if current_notes_texts:
                    result["notes"].append(current_notes_texts)
                    current_notes_texts = []
                in_notes = False
            elif ended_container[0] == RT_MAIN_MASTER_CONTAINER:
                in_master = False

        if is_container:
            container_end = offset + 8 + rec_len
            container_stack.append((rec_type, offset, container_end))

            if rec_type == RT_SLIDE_CONTAINER:
                in_slide = True
                in_notes = False
            elif rec_type == RT_NOTES_CONTAINER:
                in_notes = True
                in_slide = False
            elif rec_type == RT_MAIN_MASTER_CONTAINER:
                in_master = True

        # Extract text
        if rec_type == RT_TEXT_HEADER_ATOM:
            if rec_len >= 4:
                current_text_type = struct.unpack("<I", record_data[:4])[0]

        elif rec_type in (RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM, RT_CSTRING):
            text = None

            if rec_type == RT_TEXT_CHARS_ATOM:
                try:
                    text = record_data.decode("utf-16-le")
                except UnicodeDecodeError:
                    pass
            elif rec_type == RT_TEXT_BYTES_ATOM:
                try:
                    text = record_data.decode("latin-1")
                except UnicodeDecodeError:
                    pass
            elif rec_type == RT_CSTRING:
                try:
                    text = record_data.decode("utf-16-le")
                except UnicodeDecodeError:
                    pass

            if text:
                text = _clean_text(text)
                if text:
                    block = PptTextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (PPT_TEXT_TYPE_TITLE, PPT_TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            PPT_TEXT_TYPE_BODY,
                            PPT_TEXT_TYPE_CENTER_BODY,
                            PPT_TEXT_TYPE_HALF_BODY,
                            PPT_TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == PPT_TEXT_TYPE_NOTES,
                    )

                    if in_notes:
                        current_notes_texts.append(block)
                    elif in_slide:
                        current_slide_texts.append(block)
                    elif in_master:
                        current_master_texts.append(block)

        if is_container:
            offset += 8
        else:
            offset += 8 + rec_len

    # Collect any remaining text
    if current_slide_texts:
        result["slides"].append(current_slide_texts)
    if current_notes_texts:
        result["notes"].append(current_notes_texts)
    if current_master_texts:
        result["master"] = current_master_texts

    return result


def _extract_all_text_raw(data: bytes) -> list[str]:
    """
    Extract all text from the stream without structure parsing.
    Fallback for when structured parsing fails.
    """
    texts = []
    offset = 0

    while offset < len(data) - 8:
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            break

        if rec_len > len(data) - offset - 8:
            offset += 1
            continue

        rec_ver = rec_ver_instance & 0x0F
        is_container = rec_ver == 0x0F

        record_data = data[offset + 8 : offset + 8 + rec_len]

        if rec_type == RT_TEXT_CHARS_ATOM:
            try:
                text = record_data.decode("utf-16-le")
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except UnicodeDecodeError:
                pass

        elif rec_type == RT_TEXT_BYTES_ATOM:
            try:
                text = record_data.decode("latin-1")
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except UnicodeDecodeError:
                pass

        elif rec_type == RT_CSTRING:
            try:
                text = record_data.decode("utf-16-le")
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except UnicodeDecodeError:
                pass

        if is_container:
            offset += 8
        else:
            offset += 8 + rec_len

    return texts


def _clean_text(text: str) -> str:
    """Clean extracted text by removing control characters and extra whitespace."""
    if not text:
        return ""

    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = text.replace("\x0b", "\n")
    text = text.replace("\x0c", "\n")

    cleaned = []
    for char in text:
        if char == "\n" or char == "\t" or ord(char) >= 32:
            cleaned.append(char)
    text = "".join(cleaned)

    lines = text.split("\n")
    lines = [" ".join(line.split()) for line in lines]
    text = "\n".join(line for line in lines if line)

    return text.strip()


def _extract_ppt_metadata(file_like: BinaryIO) -> PptMetadata:
    """
    Extract only metadata from a PPT file.

    Args:
        file_like: A file-like object containing the PPT file data.

    Returns:
        PPTMetadata dataclass with metadata properties.
    """
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise ValueError("Not a valid OLE file")

    file_like.seek(0)

    with olefile.OleFileIO(file_like) as ole:
        return _extract_metadata(ole)
