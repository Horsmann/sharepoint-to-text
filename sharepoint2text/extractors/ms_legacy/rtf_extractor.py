"""
RTF Document Extractor
======================

Extracts text content, metadata, and structure from RTF (Rich Text Format)
files. RTF is a cross-platform document format developed by Microsoft that
uses plain text with control words to represent formatting.

File Format Background
----------------------
RTF files are plain text documents using ASCII control sequences to encode
formatting. The format was designed for cross-platform document interchange.

Basic Structure:
    - Files start with "{\\rtf" header
    - Content is organized in nested groups (braces { })
    - Control words start with backslash (e.g., \\par, \\bold)
    - Control words may have numeric parameters (e.g., \\fs24 for font size)
    - Special characters are escaped or encoded

Key Control Words:
    - \\rtf1: RTF version 1 header
    - \\fonttbl: Font table definition
    - \\colortbl: Color table definition
    - \\stylesheet: Style definitions
    - \\info: Document metadata
    - \\par: Paragraph break
    - \\page: Page break
    - \\header, \\footer: Header/footer content
    - \\footnote: Footnote content
    - \\field: Field (hyperlinks, page numbers, etc.)

Text Encoding:
    - Plain ASCII: Characters 32-127
    - Hex escapes: \\'xx for bytes (e.g., \\'e4 for ä)
    - Unicode: \\u<decimal>? for Unicode code points
    - Code pages specified via \\ansicpgNNNN

Dependencies
------------
No external dependencies - uses Python standard library only.

Known Limitations
-----------------
- Very complex RTF structures may not parse completely
- Embedded OLE objects are not extracted
- Some advanced formatting may be lost in text extraction
- Equation/math content may not extract properly
- Right-to-left text handling is basic
- Nested tables may not preserve structure

Encoding Handling
-----------------
The parser attempts multiple encodings in order:
    1. UTF-8
    2. CP1252 (Windows-1252)
    3. Latin-1

Unicode escapes (\\uNNNN) are decoded to proper characters.
Hex escapes (\\'xx) are decoded using the document's code page.

Destination Groups
------------------
Some RTF groups are "destinations" that should be skipped during
text extraction (they contain metadata or non-text content):
    - fonttbl, colortbl, stylesheet
    - info (metadata is extracted separately)
    - header, footer (extracted separately)
    - pict (images)
    - fldinst (field instructions)

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_legacy.rtf_extractor import read_rtf
    >>>
    >>> with open("document.rtf", "rb") as f:
    ...     for doc in read_rtf(io.BytesIO(f.read()), path="document.rtf"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(f"Pages: {len(doc.pages)}")
    ...         print(doc.full_text[:500])

See Also
--------
- RTF 1.9.1 Specification: https://www.microsoft.com/en-us/download/details.aspx?id=10725
- doc_extractor: For binary Word .doc format
- docx_extractor: For modern Word .docx format

Maintenance Notes
-----------------
- The parser uses regex extensively for pattern matching
- Some RTF files may have non-standard extensions
- Error handling is generous - partial content returned on failures
- The _strip_rtf_full method does character-by-character parsing
- Page detection relies on \\page control word
"""

import io
import logging
import re
from typing import Any, Generator, List, Optional

from sharepoint2text.extractors.data_types import (
    RtfAnnotation,
    RtfBookmark,
    RtfColor,
    RtfContent,
    RtfField,
    RtfFont,
    RtfFootnote,
    RtfHeaderFooter,
    RtfHyperlink,
    RtfImage,
    RtfMetadata,
    RtfParagraph,
    RtfStyle,
)

logger = logging.getLogger(__name__)


class _RtfParser:
    """
    Low-level RTF parser that tokenizes and processes RTF content.

    This class handles the parsing of RTF documents, extracting text,
    formatting information, and metadata. It uses a combination of
    regex patterns and character-by-character parsing.

    Parsing Process:
        1. Decode raw bytes to string (UTF-8 -> CP1252 -> Latin-1)
        2. Validate RTF header ({\\rtf)
        3. Extract font table (\\fonttbl)
        4. Extract color table (\\colortbl)
        5. Extract stylesheet (\\stylesheet)
        6. Extract metadata (\\info)
        7. Extract headers/footers
        8. Extract body text with page break detection
        9. Extract hyperlinks, fields, images, footnotes

    Attributes:
        data: Raw RTF file bytes
        pos: Current parsing position
        length: Total data length
        fonts: Extracted font definitions
        colors: Extracted color definitions
        styles: Extracted style definitions
        metadata: Extracted document metadata
        paragraphs: Extracted paragraph content
        headers: Header content by type
        footers: Footer content by type
        hyperlinks: Extracted hyperlinks
        bookmarks: Extracted bookmarks
        fields: Extracted fields (page numbers, dates, etc.)
        images: Extracted image references
        footnotes: Extracted footnotes
        annotations: Extracted annotations/comments
        pages: Text content split by page breaks
        raw_text_blocks: All text blocks in order

    Error Handling:
        If parsing fails at any stage, falls back to simple text
        extraction, returning whatever content can be retrieved.
    """

    # Destination groups to skip during text extraction (contain metadata/non-text)
    SKIP_DESTINATIONS = frozenset(
        [
            "fonttbl",
            "colortbl",
            "stylesheet",
            "info",
            "pict",
            "object",
            "datafield",
            "fldinst",
            "ftnsep",
            "ftnsepc",
            "aftnsep",
            "aftnsepc",
            "header",
            "footer",
            "headerl",
            "headerr",
            "headerf",
            "footerl",
            "footerr",
            "footerf",
            "pnseclvl",
            "xmlnstbl",
            "rsidtbl",
            "mmathPr",
            "generator",
            "listtable",
            "listoverridetable",
            "revtbl",
        ]
    )

    # RTF special characters mapping - control words to actual characters
    SPECIAL_CHARS = {
        "par": "\n",
        "line": "\n",
        "tab": "\t",
        "lquote": "'",
        "rquote": "'",
        "ldblquote": '"',
        "rdblquote": '"',
        "bullet": "•",
        "endash": "–",
        "emdash": "—",
        "~": "\u00a0",  # Non-breaking space
        "_": "\u00ad",  # Optional hyphen
        "-": "\u00ad",  # Optional hyphen
        "enspace": "\u2002",
        "emspace": "\u2003",
        "qmspace": "\u2005",
    }

    # Unicode escapes for common characters
    UNICODE_MAP = {
        0x2018: "'",  # Left single quote
        0x2019: "'",  # Right single quote
        0x201C: '"',  # Left double quote
        0x201D: '"',  # Right double quote
        0x2013: "–",  # En dash
        0x2014: "—",  # Em dash
        0x2026: "...",  # Ellipsis
    }

    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.length = len(data)

        # Extracted data
        self.fonts: List[RtfFont] = []
        self.colors: List[RtfColor] = []
        self.styles: List[RtfStyle] = []
        self.metadata = RtfMetadata()
        self.paragraphs: List[RtfParagraph] = []
        self.headers: List[RtfHeaderFooter] = []
        self.footers: List[RtfHeaderFooter] = []
        self.hyperlinks: List[RtfHyperlink] = []
        self.bookmarks: List[RtfBookmark] = []
        self.fields: List[RtfField] = []
        self.images: List[RtfImage] = []
        self.footnotes: List[RtfFootnote] = []
        self.annotations: List[RtfAnnotation] = []
        self.pages: List[str] = []
        self.raw_text_blocks: List[str] = []

        # Parsing state
        self._current_text: List[str] = []
        self._current_para = RtfParagraph()
        self._in_header_footer: Optional[str] = None
        self._in_footnote = False
        self._in_field = False
        self._field_instruction = ""
        self._field_result = ""
        self._skip_destination = False
        self._group_depth = 0
        self._skip_until_depth = -1
        self._current_hyperlink_url = ""

    def parse(self) -> RtfContent:
        """Parse the RTF document and return extracted content."""
        try:
            # Decode the raw bytes to string
            text = self._decode_rtf()

            # Check RTF signature
            if not text.startswith("{\\rtf"):
                logger.warning("Not a valid RTF file - missing RTF header")
                return RtfContent(full_text=text)

            # Extract various components
            self._extract_fonts(text)
            self._extract_colors(text)
            self._extract_styles(text)
            self._extract_metadata(text)
            self._extract_headers_footers(text)
            self._extract_body_text(text)
            self._extract_hyperlinks(text)
            self._extract_fields(text)
            self._extract_images(text)
            self._extract_footnotes(text)

            # Build full text
            full_text = "\n".join(p.text for p in self.paragraphs if p.text.strip())

            return RtfContent(
                metadata=self.metadata,
                fonts=self.fonts,
                colors=self.colors,
                styles=self.styles,
                paragraphs=self.paragraphs,
                headers=self.headers,
                footers=self.footers,
                hyperlinks=self.hyperlinks,
                bookmarks=self.bookmarks,
                fields=self.fields,
                images=self.images,
                footnotes=self.footnotes,
                annotations=self.annotations,
                pages=self.pages,
                full_text=full_text,
                raw_text_blocks=self.raw_text_blocks,
            )
        except Exception as e:
            logger.error(f"RTF parsing failed: {e}")
            # Return basic content with raw text if parsing fails
            try:
                text = self._decode_rtf()
                plain = self._strip_rtf_simple(text)
                return RtfContent(full_text=plain)
            except Exception:
                return RtfContent()

    def _decode_rtf(self) -> str:
        """Decode RTF bytes to string, handling various encodings."""
        # Try UTF-8 first, then fall back to cp1252 (Windows-1252)
        for encoding in ["utf-8", "cp1252", "latin-1"]:
            try:
                return self.data.decode(encoding)
            except UnicodeDecodeError:
                continue
        # Last resort: decode with error handling
        return self.data.decode("cp1252", errors="replace")

    def _extract_fonts(self, text: str) -> None:
        """Extract font table from RTF."""
        # Find the font table
        fonttbl_match = re.search(r"\{\\fonttbl(.*?)\}", text, re.DOTALL)
        if not fonttbl_match:
            return

        fonttbl = fonttbl_match.group(1)

        # Parse individual font definitions
        # Pattern: {\f<id>\f<family>\fcharset<charset> <name>;}
        font_pattern = re.compile(
            r"\{\\f(\d+)"  # Font ID
            r"(?:\\fbidi\s*)?"  # Optional bidi
            r"\\f(\w+)"  # Font family
            r"(?:\\fcharset(\d+))?"  # Optional charset
            r"(?:\\fprq(\d+))?"  # Optional pitch
            r"(?:\{[^}]*\})?"  # Optional panose
            r"\s*([^;}]*)"  # Font name
            r";?\}",
            re.IGNORECASE,
        )

        for match in font_pattern.finditer(fonttbl):
            font_id = int(match.group(1))
            font_family = match.group(2) or ""
            charset = int(match.group(3)) if match.group(3) else 0
            pitch = int(match.group(4)) if match.group(4) else 0
            font_name = match.group(5).strip()

            self.fonts.append(
                RtfFont(
                    font_id=font_id,
                    font_family=font_family,
                    font_name=font_name,
                    charset=charset,
                    pitch=pitch,
                )
            )

    def _extract_colors(self, text: str) -> None:
        """Extract color table from RTF."""
        colortbl_match = re.search(r"\{\\colortbl\s*;?(.*?)\}", text, re.DOTALL)
        if not colortbl_match:
            return

        colortbl = colortbl_match.group(1)

        # Parse colors: \red<n>\green<n>\blue<n>;
        color_pattern = re.compile(r"\\red(\d+)\\green(\d+)\\blue(\d+)", re.IGNORECASE)

        idx = 0
        for match in color_pattern.finditer(colortbl):
            self.colors.append(
                RtfColor(
                    index=idx,
                    red=int(match.group(1)),
                    green=int(match.group(2)),
                    blue=int(match.group(3)),
                )
            )
            idx += 1

    def _extract_styles(self, text: str) -> None:
        """Extract stylesheet from RTF."""
        stylesheet_match = re.search(
            r"\{\\stylesheet(.*?)\}\s*(?=\{|\\)", text, re.DOTALL
        )
        if not stylesheet_match:
            return

        stylesheet = stylesheet_match.group(1)

        # Parse style definitions
        # Pattern: {\s<id>...<style name>;}
        style_pattern = re.compile(
            r"\{(?:\\[*]?)?\\s(\d+)"  # Style ID
            r"[^}]*?"  # Style properties
            r"(?:\\sbasedon(\d+))?"  # Based on
            r"(?:\\snext(\d+))?"  # Next style
            r"[^}]*?"
            r"\s+([^;}]+)"  # Style name
            r";?\}",
            re.IGNORECASE,
        )

        for match in style_pattern.finditer(stylesheet):
            style_id = int(match.group(1))
            based_on = int(match.group(2)) if match.group(2) else None
            next_style = int(match.group(3)) if match.group(3) else None
            style_name = match.group(4).strip()

            # Determine style type
            style_type = "paragraph"
            if "\\cs" in match.group(0):
                style_type = "character"
            elif "\\ts" in match.group(0):
                style_type = "table"

            self.styles.append(
                RtfStyle(
                    style_id=style_id,
                    style_type=style_type,
                    style_name=style_name,
                    based_on=based_on,
                    next_style=next_style,
                )
            )

    def _extract_metadata(self, text: str) -> None:
        """Extract document info/metadata from RTF."""
        info_match = re.search(
            r"\{\\info\s*(.*?)\}(?=\s*\\|\s*\{[^\\])", text, re.DOTALL
        )
        if not info_match:
            # Try alternate pattern
            info_match = re.search(r"\{\\info\s*((?:\{[^}]*\}\s*)+)\}", text, re.DOTALL)

        if not info_match:
            return

        info = info_match.group(1)

        # Extract individual metadata fields
        def get_info_value(pattern: str) -> str:
            match = re.search(pattern, info, re.IGNORECASE | re.DOTALL)
            if match:
                value = match.group(1).strip()
                # Handle hex escapes first (before removing control sequences)
                value = re.sub(
                    r"\\'([0-9a-fA-F]{2})",
                    lambda m: chr(int(m.group(1), 16)),
                    value,
                )
                # Handle non-breaking space
                value = value.replace("\\~", " ")
                # Remove RTF control sequences from the value
                value = re.sub(r"\\[a-z]+\d*\s*", "", value)
                # Remove braces
                value = value.replace("{", "").replace("}", "")
                return value.strip()
            return ""

        self.metadata.title = get_info_value(r"\{\\title\s+([^}]*)\}")
        self.metadata.subject = get_info_value(r"\{\\subject\s+([^}]*)\}")
        self.metadata.author = get_info_value(r"\{\\author\s+([^}]*)\}")
        self.metadata.keywords = get_info_value(r"\{\\keywords\s+([^}]*)\}")
        self.metadata.comments = get_info_value(r"\{\\comment\s+([^}]*)\}")
        self.metadata.operator = get_info_value(r"\{\\operator\s+([^}]*)\}")
        self.metadata.category = get_info_value(r"\{\\[*]?\\?category\s+([^}]*)\}")
        self.metadata.manager = get_info_value(r"\{\\manager\s+([^}]*)\}")
        self.metadata.company = get_info_value(r"\{\\company\s+([^}]*)\}")
        self.metadata.doc_comment = get_info_value(r"\{\\doccomm\s+([^}]*)\}")

        # Extract numeric values
        version_match = re.search(r"\{\\version(\d+)\}", info)
        if version_match:
            self.metadata.version = int(version_match.group(1))

        revision_match = re.search(r"\{\\vern(\d+)\}", info)
        if revision_match:
            self.metadata.revision = int(revision_match.group(1))

        nofpages_match = re.search(r"\{\\nofpages(\d+)\}", info)
        if nofpages_match:
            self.metadata.num_pages = int(nofpages_match.group(1))

        nofwords_match = re.search(r"\{\\nofwords(\d+)\}", info)
        if nofwords_match:
            self.metadata.num_words = int(nofwords_match.group(1))

        nofchars_match = re.search(r"\{\\nofchars(\d+)\}", info)
        if nofchars_match:
            self.metadata.num_chars = int(nofchars_match.group(1))

        nofcharsws_match = re.search(r"\{\\nofcharsws(\d+)\}", info)
        if nofcharsws_match:
            self.metadata.num_chars_with_spaces = int(nofcharsws_match.group(1))

        # Extract dates
        def parse_rtf_date(pattern: str) -> str:
            match = re.search(pattern, info, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                yr_match = re.search(r"\\yr(\d+)", date_str)
                mo_match = re.search(r"\\mo(\d+)", date_str)
                dy_match = re.search(r"\\dy(\d+)", date_str)
                hr_match = re.search(r"\\hr(\d+)", date_str)
                min_match = re.search(r"\\min(\d+)", date_str)

                if yr_match and mo_match and dy_match:
                    year = int(yr_match.group(1))
                    month = int(mo_match.group(1))
                    day = int(dy_match.group(1))
                    hour = int(hr_match.group(1)) if hr_match else 0
                    minute = int(min_match.group(1)) if min_match else 0

                    return (
                        f"{year:04d}-{month:02d}-{day:02d}T{hour:02d}:{minute:02d}:00"
                    )
            return ""

        self.metadata.created = parse_rtf_date(r"\{\\creatim([^}]*)\}")
        self.metadata.modified = parse_rtf_date(r"\{\\revtim([^}]*)\}")

    def _extract_headers_footers(self, text: str) -> None:
        """Extract headers and footers from RTF."""
        header_types = [
            ("header", "header"),
            ("headerl", "headerl"),  # Left page header
            ("headerr", "headerr"),  # Right page header
            ("headerf", "headerf"),  # First page header
            ("footer", "footer"),
            ("footerl", "footerl"),  # Left page footer
            ("footerr", "footerr"),  # Right page footer
            ("footerf", "footerf"),  # First page footer
        ]

        for rtf_keyword, hf_type in header_types:
            pattern = re.compile(
                r"\{\\" + rtf_keyword + r"\s+(.*?)\}(?=\s*\{|\s*\\)",
                re.DOTALL | re.IGNORECASE,
            )

            for match in pattern.finditer(text):
                content = match.group(1)
                extracted_text = self._strip_rtf_simple(content)

                if extracted_text.strip():
                    hf = RtfHeaderFooter(type=hf_type, text=extracted_text.strip())
                    if "header" in hf_type:
                        self.headers.append(hf)
                    else:
                        self.footers.append(hf)

    def _extract_body_text(self, text: str) -> None:
        """Extract body text from RTF, building paragraphs and pages.

        Pages are split on explicit \\page breaks. If no page breaks exist,
        the entire document is treated as a single page.
        """
        # Remove destination groups we don't want in the body
        destinations_to_remove = [
            r"\{\\fonttbl[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
            r"\{\\colortbl[^{}]*\}",
            r"\{\\stylesheet[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
            r"\{\\info[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
            r"\{\\[*]?\\?header[lrf]?\s+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
            r"\{\\[*]?\\?footer[lrf]?\s+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
            r"\{\\\*\\[a-z]+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",  # Destination groups
        ]

        body = text
        for pattern in destinations_to_remove:
            body = re.sub(pattern, "", body, flags=re.DOTALL | re.IGNORECASE)

        # Now extract text from the body with page break detection
        extracted = self._strip_rtf_full_with_pages(body)

        # Split into paragraphs (for backward compatibility)
        para_texts = re.split(r"\n+", extracted)
        for para_text in para_texts:
            cleaned = para_text.strip()
            if cleaned:
                self.paragraphs.append(RtfParagraph(text=cleaned))
                self.raw_text_blocks.append(cleaned)

    def _extract_hyperlinks(self, text: str) -> None:
        """Extract hyperlinks from RTF."""
        # Pattern for HYPERLINK fields
        hyperlink_pattern = re.compile(
            r'\\field\s*\{[^}]*\\fldinst\s*\{[^}]*HYPERLINK\s+"([^"]+)"[^}]*\}[^}]*\{[^}]*\\fldrslt\s*\{([^}]*)\}',
            re.IGNORECASE | re.DOTALL,
        )

        for match in hyperlink_pattern.finditer(text):
            url = match.group(1).strip()
            link_text = self._strip_rtf_simple(match.group(2)).strip()

            if url:
                self.hyperlinks.append(RtfHyperlink(text=link_text, url=url))

    def _extract_fields(self, text: str) -> None:
        """Extract fields (page numbers, dates, etc.) from RTF."""
        field_pattern = re.compile(
            r"\\field\s*\{[^}]*\\fldinst\s*\{([^}]*)\}[^}]*\{[^}]*\\fldrslt\s*\{([^}]*)\}",
            re.IGNORECASE | re.DOTALL,
        )

        for match in field_pattern.finditer(text):
            instruction = self._strip_rtf_simple(match.group(1)).strip()
            result = self._strip_rtf_simple(match.group(2)).strip()

            # Skip hyperlinks (handled separately)
            if "HYPERLINK" in instruction.upper():
                continue

            # Determine field type
            field_type = "unknown"
            if "PAGE" in instruction.upper():
                field_type = "page"
            elif "DATE" in instruction.upper():
                field_type = "date"
            elif "TIME" in instruction.upper():
                field_type = "time"
            elif "STYLEREF" in instruction.upper():
                field_type = "styleref"
            elif "TOC" in instruction.upper():
                field_type = "toc"

            self.fields.append(
                RtfField(
                    field_type=field_type,
                    field_instruction=instruction,
                    field_result=result,
                )
            )

    def _extract_images(self, text: str) -> None:
        """Extract embedded images from RTF."""
        # Pattern for picture groups
        pict_pattern = re.compile(
            r"\{\\pict([^}]*(?:\{[^}]*\}[^}]*)*)\}", re.DOTALL | re.IGNORECASE
        )

        for match in pict_pattern.finditer(text):
            pict_content = match.group(1)

            # Determine image type
            image_type = "unknown"
            if "\\pngblip" in pict_content:
                image_type = "png"
            elif "\\jpegblip" in pict_content:
                image_type = "jpeg"
            elif "\\emfblip" in pict_content:
                image_type = "emf"
            elif "\\wmetafile" in pict_content:
                image_type = "wmf"

            # Extract dimensions
            width = 0
            height = 0
            width_match = re.search(r"\\picw(\d+)", pict_content)
            height_match = re.search(r"\\pich(\d+)", pict_content)
            if width_match:
                width = int(width_match.group(1))
            if height_match:
                height = int(height_match.group(1))

            # Extract hex data (simplified - just note that there's image data)
            hex_data_match = re.search(r"([0-9a-fA-F]{20,})", pict_content)
            data = None
            if hex_data_match:
                try:
                    hex_str = hex_data_match.group(1)
                    data = bytes.fromhex(hex_str)
                except ValueError:
                    pass

            self.images.append(
                RtfImage(
                    image_type=image_type,
                    width=width,
                    height=height,
                    data=data,
                )
            )

    def _extract_footnotes(self, text: str) -> None:
        """Extract footnotes from RTF."""
        # Pattern for footnotes
        footnote_pattern = re.compile(
            r"\{\\footnote\s*([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}",
            re.DOTALL | re.IGNORECASE,
        )

        footnote_id = 1
        for match in footnote_pattern.finditer(text):
            content = match.group(1)
            footnote_text = self._strip_rtf_simple(content).strip()

            if footnote_text:
                self.footnotes.append(RtfFootnote(id=footnote_id, text=footnote_text))
                footnote_id += 1

    def _strip_rtf_simple(self, text: str) -> str:
        """Simple RTF stripping - removes control words and groups."""
        result = text

        # Handle unicode escapes \u<decimal>?
        result = re.sub(
            r"\\u(-?\d+)\??",
            lambda m: chr(int(m.group(1)) & 0xFFFF),
            result,
        )

        # Handle hex escapes \'xx
        result = re.sub(
            r"\\'([0-9a-fA-F]{2})",
            lambda m: chr(int(m.group(1), 16)),
            result,
        )

        # Replace special RTF characters
        for keyword, char in self.SPECIAL_CHARS.items():
            result = re.sub(r"\\" + re.escape(keyword) + r"(?:\s|$)", char, result)

        # Remove control words with parameters
        result = re.sub(r"\\[a-z]+(-?\d+)?\s?", "", result, flags=re.IGNORECASE)

        # Remove braces
        result = result.replace("{", "").replace("}", "")

        # Normalize whitespace
        result = re.sub(r"[ \t]+", " ", result)
        result = re.sub(r"\n{3,}", "\n\n", result)

        return result.strip()

    def _is_skip_destination(self, ahead: str) -> bool:
        """Check if the lookahead text indicates a destination to skip."""
        if ahead.startswith("\\*"):
            return True
        for kw in self.SKIP_DESTINATIONS:
            if ahead.startswith("\\" + kw):
                return True
        return False

    def _strip_rtf_full_with_pages(self, text: str) -> str:
        """Full RTF stripping with page break detection.

        Detects \\page control words and splits content into pages.
        Populates self.pages with text content per page.
        Returns the full text as a single string.
        """
        result = []
        current_page = []
        i = 0
        n = len(text)
        group_depth = 0
        skip_group = False
        skip_depth = 0

        def flush_page():
            """Save current page content and start a new page."""
            page_text = "".join(current_page).strip()
            # Normalize whitespace
            page_text = re.sub(r"[ \t]+", " ", page_text)
            page_text = re.sub(r"\n{3,}", "\n\n", page_text)
            if page_text:
                self.pages.append(page_text)
            current_page.clear()

        while i < n:
            char = text[i]

            if char == "{":
                group_depth += 1
                # Check if this is a destination to skip
                if i + 1 < n and text[i + 1] == "\\":
                    ahead = text[i + 1 : i + 30]
                    if self._is_skip_destination(ahead):
                        skip_group = True
                        skip_depth = group_depth
                i += 1

            elif char == "}":
                if skip_group and group_depth == skip_depth:
                    skip_group = False
                group_depth -= 1
                i += 1

            elif skip_group:
                i += 1

            elif char == "\\":
                # Control word or escape
                if i + 1 >= n:
                    i += 1
                    continue

                next_char = text[i + 1]

                # Escape sequences
                if next_char in "\\{}":
                    current_page.append(next_char)
                    result.append(next_char)
                    i += 2

                # Unicode escape
                elif next_char == "u":
                    unicode_match = re.match(r"\\u(-?\d+)\??", text[i:])
                    if unicode_match:
                        code = int(unicode_match.group(1))
                        char_val = chr(code & 0xFFFF)
                        current_page.append(char_val)
                        result.append(char_val)
                        i += len(unicode_match.group(0))
                    else:
                        i += 2

                # Hex escape
                elif next_char == "'":
                    if i + 3 < n:
                        hex_val = text[i + 2 : i + 4]
                        try:
                            char_val = chr(int(hex_val, 16))
                            current_page.append(char_val)
                            result.append(char_val)
                        except ValueError:
                            pass
                        i += 4
                    else:
                        i += 2

                # Control word
                elif next_char.isalpha():
                    # Find end of control word
                    j = i + 1
                    while j < n and text[j].isalpha():
                        j += 1
                    # Skip optional numeric parameter
                    if j < n and (text[j].isdigit() or text[j] == "-"):
                        while j < n and (text[j].isdigit() or text[j] == "-"):
                            j += 1
                    # Skip optional trailing space
                    if j < n and text[j] == " ":
                        j += 1

                    control_word = text[i + 1 : j].rstrip()
                    # Remove numeric suffix for lookup
                    word_only = re.sub(r"-?\d+$", "", control_word)

                    # Check for page break
                    # \page is explicit page break
                    # \sect with \sbkpage is section break with page break
                    if word_only == "page" or word_only == "sbkpage":
                        flush_page()
                    elif word_only in self.SPECIAL_CHARS:
                        char_val = self.SPECIAL_CHARS[word_only]
                        current_page.append(char_val)
                        result.append(char_val)

                    i = j

                else:
                    # Handle \~ (non-breaking space), \- (optional hyphen), etc.
                    if next_char == "~":
                        current_page.append("\u00a0")
                        result.append("\u00a0")
                    elif next_char == "-":
                        pass  # Optional hyphen - don't add
                    elif next_char == "_":
                        current_page.append("\u00ad")
                        result.append("\u00ad")
                    i += 2

            else:
                # Regular character
                if char not in "\r":  # Skip carriage returns
                    current_page.append(char)
                    result.append(char)
                i += 1

        # Flush the last page
        flush_page()

        # If no page breaks were found, treat the entire document as one page
        if not self.pages:
            full_text = "".join(result).strip()
            full_text = re.sub(r"[ \t]+", " ", full_text)
            full_text = re.sub(r"\n{3,}", "\n\n", full_text)
            if full_text:
                self.pages.append(full_text)

        return "".join(result)


def read_rtf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[RtfContent, Any, None]:
    """
    Extract all relevant content from an RTF file.

    Primary entry point for RTF file extraction. Parses the RTF markup,
    extracts text content, and builds structured output including pages,
    paragraphs, and metadata.

    This function uses a generator pattern for API consistency with other
    extractors, even though RTF files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete RTF file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned RtfContent.metadata.

    Yields:
        RtfContent: Single RtfContent object containing:
            - full_text: Complete document text as single string
            - pages: List of text strings, one per page
            - paragraphs: List of RtfParagraph objects
            - metadata: RtfMetadata with title, author, dates, etc.
            - fonts: List of RtfFont definitions
            - colors: List of RtfColor definitions
            - styles: List of RtfStyle definitions
            - headers, footers: Header/footer content
            - hyperlinks, bookmarks, fields: Document elements
            - images: Extracted image references
            - footnotes: Footnote content

    Example:
        >>> import io
        >>> with open("report.rtf", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_rtf(data, path="report.rtf"):
        ...         print(f"Author: {doc.metadata.author}")
        ...         print(f"Pages: {len(doc.pages)}")
        ...         for i, page in enumerate(doc.pages, 1):
        ...             print(f"Page {i}: {page[:50]}...")

    Implementation Notes:
        - Entire file is loaded into memory
        - Parsing errors fall back to simple text extraction
        - Invalid RTF files return partial content rather than failing
    """
    logger.debug("Reading RTF file")
    file_like.seek(0)
    data = file_like.read()

    parser = _RtfParser(data)
    content = parser.parse()

    content.metadata.populate_from_path(path)

    yield content
