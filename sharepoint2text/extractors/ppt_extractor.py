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
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, BinaryIO

import olefile

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

# Text placeholder types (from TextHeaderAtom)
TEXT_TYPE_TITLE = 0  # Title
TEXT_TYPE_BODY = 1  # Body
TEXT_TYPE_NOTES = 2  # Notes
TEXT_TYPE_OTHER = 4  # Other (not title/body/notes)
TEXT_TYPE_CENTER_BODY = 5  # Center body (subtitle)
TEXT_TYPE_CENTER_TITLE = 6  # Center title
TEXT_TYPE_HALF_BODY = 7  # Half body
TEXT_TYPE_QUARTER_BODY = 8  # Quarter body


@dataclass
class TextBlock:
    """Represents a block of text with its type and context."""

    text: str
    text_type: int | None = None  # From TextHeaderAtom
    is_title: bool = False
    is_body: bool = False
    is_notes: bool = False

    @property
    def type_name(self) -> str:
        """Human-readable text type name."""
        type_names = {
            TEXT_TYPE_TITLE: "title",
            TEXT_TYPE_BODY: "body",
            TEXT_TYPE_NOTES: "notes",
            TEXT_TYPE_OTHER: "other",
            TEXT_TYPE_CENTER_BODY: "subtitle",
            TEXT_TYPE_CENTER_TITLE: "center_title",
            TEXT_TYPE_HALF_BODY: "half_body",
            TEXT_TYPE_QUARTER_BODY: "quarter_body",
        }
        return type_names.get(self.text_type, "unknown")


@dataclass
class SlideContent:
    """Represents the content of a single slide."""

    slide_number: int
    title: str | None = None
    body_text: list[str] = field(default_factory=list)
    other_text: list[str] = field(default_factory=list)
    all_text: list[TextBlock] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def text_combined(self) -> str:
        """All text from this slide combined."""
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "slide_number": self.slide_number,
            "title": self.title,
            "body_text": self.body_text,
            "other_text": self.other_text,
            # 'all_text': [{'text': tb.text, 'type': tb.type_name} for tb in self.all_text],
            "notes": self.notes,
            # 'text_combined': self.text_combined,
        }


@dataclass
class PPTContent:
    """Complete extracted content from a PPT file."""

    metadata: dict[str, Any] = field(default_factory=dict)
    slides: list[SlideContent] = field(default_factory=list)
    master_text: list[str] = field(default_factory=list)  # Text from master slides
    all_text: list[str] = field(default_factory=list)  # All text in order
    streams: list[list[str]] = field(default_factory=list)

    # @property
    # def text_combined(self) -> str:
    #     """All text combined."""
    #     return '\n'.join(self.all_text)

    @property
    def slide_count(self) -> int:
        """Number of slides extracted."""
        return len(self.slides)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "metadata": self.metadata,
            "slides": [s.to_dict() for s in self.slides],
            "slide_count": self.slide_count,
            "master_text": self.master_text,
            # 'all_text': self.all_text,
            # 'text_combined': self.text_combined,
            # 'streams': self.streams,
        }


def read_ppt(file_like: BinaryIO) -> dict[str, Any]:
    """
    Extract text content and metadata from a legacy PowerPoint (.ppt) file.

    Args:
        file_like: A file-like object (e.g., io.BytesIO) containing the PPT file data.

    Returns:
        A dictionary containing:
            - 'metadata': Dictionary of document metadata
            - 'slides': List of slide dictionaries, each with:
                - 'slide_number': int
                - 'title': str or None
                - 'body_text': list of body text strings
                - 'other_text': list of other text strings
                - 'notes': list of notes strings
            - 'slide_count': Number of slides
            - 'master_text': Text from master slides

    Raises:
        ValueError: If the file is not a valid OLE file or PPT document.
        IOError: If there's an error reading the file.
    """
    file_like.seek(0)
    content = _extract_ppt_content_structured(file_like)
    return content.to_dict()


def _extract_ppt_content_structured(file_like: BinaryIO) -> PPTContent:
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

    content = PPTContent()

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


def _extract_metadata(ole: olefile.OleFileIO) -> dict[str, Any]:
    """Extract metadata from the OLE file."""
    metadata = {}

    try:
        meta = ole.get_metadata()

        # Standard metadata attributes
        standard_attrs = [
            "title",
            "subject",
            "author",
            "keywords",
            "comments",
            "template",
            "last_saved_by",
            "revision_number",
            "total_edit_time",
            "last_printed",
            "create_time",
            "last_saved_time",
            "num_pages",
            "num_words",
            "num_chars",
            "creating_application",
            "security",
            "codepage",
        ]

        for attr in standard_attrs:
            value = getattr(meta, attr, None)
            if value is not None:
                if isinstance(value, datetime):
                    metadata[attr] = value.isoformat()
                else:
                    metadata[attr] = value

        # Document summary attributes
        docsum_attrs = [
            "category",
            "presentation_target",
            "manager",
            "company",
            "slides",
            "notes",
            "hidden_slides",
            "codepage_doc",
        ]

        for attr in docsum_attrs:
            value = getattr(meta, attr, None)
            if value is not None:
                metadata[attr] = value

    except Exception as e:
        logger.debug(e)
        pass

    if ole.exists("Current User"):
        try:
            current_user_stream = ole.openstream("Current User")
            current_user_data = current_user_stream.read()
            user_info = _parse_current_user(current_user_data)
            if user_info:
                metadata["current_user"] = user_info
        except Exception as e:
            logger.debug(e)
            pass

    return metadata


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


def _parse_ppt_document(data: bytes, content: PPTContent) -> None:
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
            slide = SlideContent(slide_number=slide_num)

            for text_block in texts:
                slide.all_text.append(text_block)
                content.all_text.append(text_block.text)

                if text_block.text_type in (TEXT_TYPE_TITLE, TEXT_TYPE_CENTER_TITLE):
                    if not slide.title:
                        slide.title = text_block.text
                    else:
                        slide.other_text.append(text_block.text)
                elif text_block.text_type == TEXT_TYPE_BODY or text_block.text_type in (
                    TEXT_TYPE_CENTER_BODY,
                    TEXT_TYPE_HALF_BODY,
                    TEXT_TYPE_QUARTER_BODY,
                ):
                    slide.body_text.append(text_block.text)
                elif text_block.text_type == TEXT_TYPE_NOTES:
                    slide.notes.append(text_block.text)
                else:
                    slide.other_text.append(text_block.text)

            content.slides.append(slide)

    # If no SlideListWithText found, fall back to container parsing
    elif container_texts["slides"]:
        for slide_num, texts in enumerate(container_texts["slides"], 1):
            slide = SlideContent(slide_number=slide_num)

            for text_block in texts:
                slide.all_text.append(text_block)
                content.all_text.append(text_block.text)

                if text_block.is_title or text_block.text_type == TEXT_TYPE_TITLE:
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
            slide = SlideContent(slide_number=1)
            for text in raw_texts:
                slide.other_text.append(text)
                slide.all_text.append(TextBlock(text=text))
            content.slides.append(slide)


def _extract_slide_list_texts(data: bytes) -> list[list[TextBlock]]:
    """
    Extract text from SlideListWithText containers.

    SlideListWithText (0x0FF0) is the most reliable source for slide text
    as it contains text in presentation order with proper slide boundaries.
    """
    slides_text: list[list[TextBlock]] = []
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


def _parse_slide_list_container(data: bytes) -> list[list[TextBlock]]:
    """Parse a SlideListWithText container to extract per-slide text."""
    slides: list[list[TextBlock]] = []
    current_slide_text: list[TextBlock] = []
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
                    block = TextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (TEXT_TYPE_TITLE, TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            TEXT_TYPE_BODY,
                            TEXT_TYPE_CENTER_BODY,
                            TEXT_TYPE_HALF_BODY,
                            TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == TEXT_TYPE_NOTES,
                    )
                    current_slide_text.append(block)
            except UnicodeDecodeError:
                pass

        elif rec_type == RT_TEXT_BYTES_ATOM:
            try:
                text = record_data.decode("latin-1")
                text = _clean_text(text)
                if text:
                    block = TextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (TEXT_TYPE_TITLE, TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            TEXT_TYPE_BODY,
                            TEXT_TYPE_CENTER_BODY,
                            TEXT_TYPE_HALF_BODY,
                            TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == TEXT_TYPE_NOTES,
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
    current_slide_texts: list[TextBlock] = []
    current_notes_texts: list[TextBlock] = []
    current_master_texts: list[TextBlock] = []
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
                    block = TextBlock(
                        text=text,
                        text_type=current_text_type,
                        is_title=current_text_type
                        in (TEXT_TYPE_TITLE, TEXT_TYPE_CENTER_TITLE),
                        is_body=current_text_type
                        in (
                            TEXT_TYPE_BODY,
                            TEXT_TYPE_CENTER_BODY,
                            TEXT_TYPE_HALF_BODY,
                            TEXT_TYPE_QUARTER_BODY,
                        ),
                        is_notes=current_text_type == TEXT_TYPE_NOTES,
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


def _extract_ppt_metadata(file_like: BinaryIO) -> dict[str, Any]:
    """
    Extract only metadata from a PPT file.

    Args:
        file_like: A file-like object containing the PPT file data.

    Returns:
        Dictionary of metadata properties.
    """
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise ValueError("Not a valid OLE file")

    file_like.seek(0)

    with olefile.OleFileIO(file_like) as ole:
        return _extract_metadata(ole)
