"""
PPT Presentation Extractor
==========================

Extracts text content and metadata from legacy Microsoft PowerPoint .ppt files
(PowerPoint 97-2003 binary format, also known as the OLE2/CFBF format).

This module parses the "PowerPoint Document" stream within the OLE container
and extracts text from various record types according to the MS-PPT specification.

File Format Background
----------------------
The .ppt format stores presentation data in a binary stream called "PowerPoint
Document" within an OLE2 compound file. The stream contains a sequence of
records, each with an 8-byte header:
    - Bytes 0-1: Version and instance info
    - Bytes 2-3: Record type (identifies what the record contains)
    - Bytes 4-7: Record length (not including header)

Records form a hierarchy with container records holding child records.

Key Record Types
----------------
Text records (contain actual text content):
    - TextCharsAtom (0x0FA0): Unicode text (UTF-16LE)
    - TextBytesAtom (0x0FA8): ASCII/ANSI text (Latin-1)
    - CString (0x0FBA): Unicode string for titles

Container records (organize structure):
    - DocumentContainer (0x03E8): Root document
    - SlideContainer (0x03EE): Individual slide
    - NotesContainer (0x03F0): Speaker notes
    - MainMasterContainer (0x03F8): Master slide template
    - SlideListWithText (0x0FF0): Primary text source

Text type records (context for following text):
    - TextHeaderAtom (0x0F9F): Indicates text type (title, body, notes, etc.)

Text Types
----------
The TextHeaderAtom indicates what kind of text follows:
    - 0: Title
    - 1: Body
    - 2: Notes
    - 4: Other
    - 5: Center body (subtitle)
    - 6: Center title
    - 7: Half body
    - 8: Quarter body

Dependencies
------------
olefile: https://github.com/decalage2/olefile
    pip install olefile

    Provides:
    - OLE compound document parsing
    - Stream enumeration and reading
    - Metadata extraction

Known Limitations
-----------------
- Encrypted/password-protected files are not supported
- Embedded OLE objects (charts, Excel sheets) are not extracted
- Image and shape text may be incomplete
- Very old PowerPoint versions (<97) may not parse correctly
- Complex animations and transitions are ignored
- SmartArt and diagrams may not extract text properly

Extraction Strategy
-------------------
The extractor uses a multi-pass approach:
    1. Primary: Extract from SlideListWithText containers (most reliable)
    2. Fallback: Parse container hierarchy for additional context
    3. Last resort: Raw text extraction from all text atoms

This approach handles various PPT file structures created by different
versions of PowerPoint and third-party applications.

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_legacy.ppt_extractor import read_ppt
    >>>
    >>> with open("presentation.ppt", "rb") as f:
    ...     for ppt in read_ppt(io.BytesIO(f.read()), path="presentation.ppt"):
    ...         print(f"Title: {ppt.metadata.title}")
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")
    ...             print(f"  Body: {slide.text_combined}")

See Also
--------
- MS-PPT specification: https://docs.microsoft.com/en-us/openspecs/office_file_formats/
- doc_extractor: For Word documents
- xls_extractor: For Excel spreadsheets

Maintenance Notes
-----------------
- Record type constants should match MS-PPT specification exactly
- The SlideListWithText extraction is preferred as it preserves order
- Container parsing is complex due to nested structure
- Text cleaning handles control characters from various PPT versions
"""

import hashlib
import logging
import struct
from datetime import datetime
from typing import Any, BinaryIO, Generator

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

# Text atom records - these contain actual text content
RT_TEXT_CHARS_ATOM = 0x0FA0  # Unicode text (UTF-16LE encoded)
RT_TEXT_BYTES_ATOM = 0x0FA8  # ASCII/ANSI text (Latin-1 encoded)
RT_CSTRING = 0x0FBA  # Unicode string (used for titles and other strings)

# Text context record - indicates what type of text follows
RT_TEXT_HEADER_ATOM = 0x0F9F  # Contains text type (title, body, notes, etc.)

# Container record types - these are parent records that contain child records
RT_DOCUMENT_CONTAINER = 0x03E8  # Root container for entire document
RT_SLIDE_CONTAINER = 0x03EE  # Container for a single slide
RT_NOTES_CONTAINER = 0x03F0  # Container for speaker notes
RT_MAIN_MASTER_CONTAINER = 0x03F8  # Container for master slide template
RT_HANDOUT_CONTAINER = 0x0FC9  # Container for handout master

# Slide list containers - primary source for ordered slide text
RT_SLIDE_LIST_WITH_TEXT = 0x0FF0  # Contains all slide text in order
RT_SLIDE_PERSIST_ATOM = 0x03F3  # Marks slide boundaries within SlideListWithText

# Drawing containers - may contain text in shapes
RT_PP_DRAWING = 0x040C  # Drawing object container
RT_PP_DRAWING_GROUP = 0x040B  # Group of drawing objects

# Text formatting and styling records (not directly used but documented)
RT_TEXT_SPEC_INFO_ATOM = 0x0FAA  # Text special info (language, etc.)
RT_STYLE_TEXT_PROP_ATOM = 0x0FA1  # Text style properties
RT_TEXT_RULES_ATOM = 0x0F98  # Text ruler/formatting rules
RT_TEXT_INTERACTIVE_INFO_ATOM = 0x0FDF  # Hyperlink info

# Outline text reference
RT_OUTLINE_TEXT_REF_ATOM = 0x0F9E  # Reference to outline text


def read_ppt(
    file_like: BinaryIO, path: str | None = None
) -> Generator[PptContent, Any, None]:
    """
    Extract text content and metadata from a legacy PowerPoint .ppt file.

    Primary entry point for PPT file extraction. Opens the OLE container,
    locates the PowerPoint Document stream, parses the record structure,
    and extracts text organized by slides.

    This function uses a generator pattern for API consistency with other
    extractors, even though PPT files contain exactly one presentation.

    Args:
        file_like: File-like object (e.g., io.BytesIO) containing the
            complete PPT file data. The stream position is reset to the
            beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned PptContent.metadata.

    Yields:
        PptContent: Single PptContent object containing:
            - metadata: PptMetadata with title, author, dates, slide count
            - slides: List of PptSlideContent objects with per-slide text
            - master_text: Text extracted from master slide templates
            - all_text: Flat list of all text in presentation order
            - streams: List of OLE stream paths (for debugging)

    Raises:
        LegacyMicrosoftParsingError: For various parsing failures:
            - File is not a valid OLE file
            - Missing required "PowerPoint Document" stream
            - Corrupted or unsupported file structure

    Example:
        >>> import io
        >>> with open("slides.ppt", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for ppt in read_ppt(data, path="slides.ppt"):
        ...         print(f"Slides: {len(ppt.slides)}")
        ...         for slide in ppt.slides:
        ...             print(f"  {slide.slide_number}: {slide.title}")

    Implementation Notes:
        - Uses multi-pass extraction for maximum text coverage
        - SlideListWithText is the preferred text source (preserves order)
        - Falls back to container parsing for notes and additional context
        - Raw extraction is used as last resort if structured parsing fails
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
    """
    Extract content from PPT file into structured PptContent object.

    This is the main extraction function that orchestrates the parsing
    process. It opens the OLE container, reads the PowerPoint Document
    stream, and delegates to specialized parsing functions.

    Args:
        file_like: File-like object containing the PPT file data.

    Returns:
        PptContent: Populated object with slides, metadata, and all text.

    Raises:
        LegacyMicrosoftParsingError: If file is not valid OLE or lacks
            PowerPoint Document stream.

    Implementation Steps:
        1. Validate OLE file format
        2. Open OLE container and list streams
        3. Extract metadata from SummaryInformation
        4. Read PowerPoint Document stream
        5. Parse stream to extract slide text and structure
    """
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

        if ole.exists("PowerPoint Document"):
            ppt_stream = ole.openstream("PowerPoint Document")
            stream_data = ppt_stream.read()

            # Parse the document structure
            _parse_ppt_document(stream_data, content)

            # Extract images from Pictures stream
            images = _extract_images_from_pictures_stream(ole)
            if images:
                _distribute_images_to_slides(content, images)
        else:
            raise LegacyMicrosoftParsingError(
                message="No 'PowerPoint Document' stream found - may not be a valid PPT file"
            )

    return content


def _distribute_images_to_slides(content: PptContent, images: list[PptImage]) -> None:
    """
    Distribute extracted images to slides.

    In PPT format, images are stored in a separate Pictures stream without
    direct references to which slide they belong to. This function distributes
    images to slides based on the number of slides and images.

    For simple cases:
    - If there's only one slide, all images go to that slide
    - If there are equal numbers of slides and images, one image per slide
    - Otherwise, images are distributed round-robin across slides

    Args:
        content: PptContent object with slides populated.
        images: List of extracted PptImage objects.
    """
    if not images:
        return

    # Ensure there's at least one slide
    if not content.slides:
        # Create a placeholder slide if none exist
        content.slides.append(PptSlideContent(slide_number=1))

    num_slides = len(content.slides)

    # Distribute images to slides
    for i, image in enumerate(images):
        slide_index = i % num_slides
        image.slide_number = content.slides[slide_index].slide_number
        content.slides[slide_index].images.append(image)


def _extract_metadata(ole: olefile.OleFileIO) -> PptMetadata:
    """
    Extract document metadata from OLE SummaryInformation stream.

    Uses olefile's built-in metadata extraction to read standard
    document properties.

    Args:
        ole: Open OleFileIO instance.

    Returns:
        PptMetadata: Populated with available metadata fields including:
            - title, subject, author, keywords, comments
            - last_saved_by, created, modified dates
            - num_slides, num_notes, num_hidden_slides
            - revision_number, category, company, manager

    Notes:
        - All exceptions are caught and logged; returns partial metadata
        - Bytes values are decoded as UTF-8 with replacement
        - Dates are converted to ISO format strings
    """
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


def _parse_ppt_document(data: bytes, content: PptContent) -> None:
    """
    Parse the PowerPoint Document stream and populate PptContent.

    This is the main parsing function that extracts text from the binary
    stream. It uses a multi-pass approach for comprehensive extraction.

    Args:
        data: Raw bytes of the PowerPoint Document stream.
        content: PptContent object to populate with extracted data.

    Parsing Strategy:
        1. First pass: Extract from SlideListWithText containers
           - These contain text in presentation order
           - Most reliable source for slide text
           - Instance 0 = slides, 1 = master, 2 = notes

        2. Second pass: Parse container hierarchy
           - Provides additional context (notes, master text)
           - Uses container stack to track current context
           - Associates text with correct slide/notes

        3. Fallback: Raw text extraction
           - Used only if structured parsing yields no results
           - Extracts all text atoms without structure

    Implementation Notes:
        - Uses state machine to track container nesting
        - TextHeaderAtom indicates type of following text
        - Slide boundaries marked by SlidePersistAtom
        - Text cleaning applied to all extracted strings
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
    Extract text from SlideListWithText containers in the stream.

    SlideListWithText (RT 0x0FF0) is the most reliable source for slide
    text as it stores text in presentation order with clear slide boundaries
    marked by SlidePersistAtom records.

    Args:
        data: Raw bytes of the PowerPoint Document stream.

    Returns:
        List of lists, where each inner list contains PptTextBlock objects
        for one slide. Order matches presentation order.

    Record Structure:
        SlideListWithText container contains:
            - SlidePersistAtom: Marks start of new slide's text
            - TextHeaderAtom: Indicates text type (title/body/notes)
            - TextCharsAtom or TextBytesAtom: The actual text content

    Instance Values:
        The SlideListWithText instance value indicates content type:
            - 0: Regular slide text
            - 1: Master slide text
            - 2: Notes text

        This function only processes instance 0 (regular slides).

    Implementation Notes:
        - Scans stream for RT_SLIDE_LIST_WITH_TEXT records
        - Delegates to _parse_slide_list_container for text extraction
        - Handles malformed records by skipping (offset += 1)
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
    """
    Parse a SlideListWithText container to extract text organized by slide.

    Processes the records within a SlideListWithText container, tracking
    slide boundaries and text context to build structured slide content.

    Args:
        data: Raw bytes of the SlideListWithText container content
            (excluding the 8-byte container header).

    Returns:
        List of lists, where each inner list contains PptTextBlock objects
        for one slide with text type information.

    Record Processing:
        - SlidePersistAtom (0x03F3): Marks slide boundary, saves current
          slide's text and starts new collection
        - TextHeaderAtom (0x0F9F): Sets current_text_type for following text
        - TextCharsAtom (0x0FA0): Unicode text, decoded as UTF-16LE
        - TextBytesAtom (0x0FA8): ASCII text, decoded as Latin-1

    PptTextBlock Attributes:
        Each text block includes:
        - text: The cleaned text content
        - text_type: Integer type from TextHeaderAtom
        - is_title: True if type indicates title text
        - is_body: True if type indicates body text
        - is_notes: True if type indicates notes text
    """
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
    Parse the container hierarchy to extract text by container context.

    This provides a secondary extraction pass that captures text the
    SlideListWithText approach might miss, particularly notes and
    master slide text.

    Args:
        data: Raw bytes of the PowerPoint Document stream.

    Returns:
        Dictionary with keys:
            - "slides": List of lists of PptTextBlock (per slide)
            - "notes": List of lists of PptTextBlock (per slide notes)
            - "master": List of PptTextBlock (master slide text)

    Container Tracking:
        Uses a stack to track nested container hierarchy:
        - Push container on stack when opening (recVer == 0x0F)
        - Pop when offset exceeds container end
        - Text is associated with innermost active container

    Container Types Tracked:
        - SlideContainer (0x03EE): Regular slide content
        - NotesContainer (0x03F0): Speaker notes
        - MainMasterContainer (0x03F8): Master slide template

    Implementation Notes:
        - Container records have recVer == 0x0F in first byte
        - Non-container records are skipped (offset += 8 + length)
        - Text from nested containers inherits context from parent
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

    This is a fallback method used when structured parsing (SlideListWithText
    and container parsing) yields no results. It simply finds all text atoms
    in the stream and extracts their content.

    Args:
        data: Raw bytes of the PowerPoint Document stream.

    Returns:
        List of extracted text strings in stream order.
        No slide or type information is preserved.

    Text Record Types Processed:
        - TextCharsAtom (0x0FA0): UTF-16LE encoded Unicode text
        - TextBytesAtom (0x0FA8): Latin-1 encoded ASCII text
        - CString (0x0FBA): UTF-16LE encoded string (titles, etc.)

    Notes:
        - Order may not match presentation order
        - Duplicate text may be extracted from different records
        - All text is cleaned before adding to result
        - Decode errors are silently ignored
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
    """
    Clean extracted text by removing control characters and normalizing whitespace.

    PowerPoint text may contain various control characters and formatting
    artifacts that need to be removed or converted for clean output.

    Args:
        text: Raw text extracted from PPT records.

    Returns:
        Cleaned text with:
        - Null characters removed
        - Line endings normalized (\\r\\n, \\r -> \\n)
        - Vertical tabs and form feeds -> newlines
        - Control characters (< 0x20) removed except \\n, \\t
        - Multiple spaces collapsed to single space per line
        - Empty lines removed

    Character Handling:
        - \\x00 (null): Removed (artifact of UTF-16LE decoding)
        - \\x0b (vertical tab): Converted to newline
        - \\x0c (form feed/page break): Converted to newline
        - \\r\\n (Windows line ending): Converted to \\n
        - \\r (old Mac line ending): Converted to \\n
    """
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
    Extract only metadata from a PPT file without parsing content.

    Utility function for quick metadata extraction when full content
    parsing is not needed. Opens the OLE container just long enough
    to read the SummaryInformation stream.

    Args:
        file_like: File-like object containing the PPT file data.

    Returns:
        PptMetadata: Populated metadata dataclass.

    Raises:
        LegacyMicrosoftParsingError: If file is not valid OLE format.

    Use Cases:
        - Quick file cataloging
        - Filtering files before full extraction
        - Displaying file properties without loading content
    """
    file_like.seek(0)

    if not olefile.isOleFile(file_like):
        raise LegacyMicrosoftParsingError("Not a valid OLE file")

    file_like.seek(0)

    with olefile.OleFileIO(file_like) as ole:
        return _extract_metadata(ole)


# =============================================================================
# Image Extraction from Pictures Stream
# =============================================================================


def _extract_images_from_pictures_stream(ole: olefile.OleFileIO) -> list[PptImage]:
    """
    Extract images from the Pictures stream of a PPT file.

    The Pictures stream contains embedded images in BLIP (Binary Large Image
    Picture) format. Each BLIP has a header followed by the actual image data.

    Args:
        ole: Open OleFileIO instance.

    Returns:
        List of PptImage objects with extracted image data.

    Notes:
        - Supports PNG, JPEG, GIF, BMP, and TIFF formats
        - EMF/WMF metafiles are extracted but may not be widely supported
        - Image dimensions are extracted where available
    """
    images: list[PptImage] = []

    if not ole.exists("Pictures"):
        return images

    try:
        pictures_stream = ole.openstream("Pictures")
        data = pictures_stream.read()
    except Exception as e:
        logger.debug(f"Failed to read Pictures stream: {e}")
        return images

    if len(data) < 25:
        return images

    offset = 0
    image_index = 0
    seen_hashes: set[str] = set()

    while offset < len(data) - 25:
        # BLIP records have an 8-byte header:
        # - 2 bytes: recVer (4 bits) + recInstance (12 bits)
        # - 2 bytes: recType
        # - 4 bytes: recLen
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            offset += 1
            continue

        if rec_len <= 0 or rec_len > len(data) - offset - 8:
            offset += 1
            continue

        record_data = data[offset + 8 : offset + 8 + rec_len]

        # Check if this is a BLIP record (image container)
        # BLIP records have specific recType values (OfficeArtBlip* types)
        is_blip = rec_type in BLIP_TYPES

        if is_blip and len(record_data) > 17:
            # BLIP structure:
            # - 16 bytes: rgbUid1 (MD4 hash of image data)
            # - 1 byte: tag
            # - remaining: image data (may have additional header depending on type)

            # Skip the 17-byte BLIP header to find the actual image data
            # For some types there's an additional 16 bytes for a secondary UID
            blip_header_size = 17

            # Check for secondary UID (indicated by instance value)
            rec_instance = (rec_ver_instance >> 4) & 0x0FFF
            if rec_instance in (BLIP_INSTANCE_PNG_2, BLIP_INSTANCE_JPEG_2):
                # Has secondary UID (rgbUid2)
                blip_header_size = 33

            if blip_header_size < len(record_data):
                image_data = record_data[blip_header_size:]

                # Try to detect the actual image type from signatures
                detected = detect_image_type(image_data)

                if detected is None:
                    # For metafiles (EMF/WMF), we may not detect by signature
                    # but still want to extract them
                    if rec_type == BLIP_TYPE_EMF:
                        detected = ("emf", "image/x-emf")
                    elif rec_type == BLIP_TYPE_WMF:
                        detected = ("wmf", "image/x-wmf")
                    elif rec_type == BLIP_TYPE_DIB:
                        # DIB needs BMP header wrapping
                        image_data = wrap_dib_as_bmp(image_data)
                        if image_data:
                            detected = ("bmp", "image/bmp")

                if detected and len(image_data) > 0:
                    # Compute hash to avoid duplicates
                    digest = hashlib.sha1(image_data).hexdigest()
                    if digest not in seen_hashes:
                        seen_hashes.add(digest)
                        image_index += 1

                        # Try to extract dimensions for PNG/JPEG
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

        offset += 8 + rec_len

    return images
