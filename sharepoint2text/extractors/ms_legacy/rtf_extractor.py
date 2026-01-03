"""
RTF Document Extractor

Extracts text content, metadata, and structure from RTF (Rich Text Format) files.
RTF is a cross-platform document format using plain text with control words.
"""

import io
import logging
import re
from typing import Any, Generator

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
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
    RtfTable,
)

logger = logging.getLogger(__name__)

# =============================================================================
# Pre-compiled regex patterns (performance optimization)
# =============================================================================

# Font table patterns
_RE_FONTTBL = re.compile(r"\{\\fonttbl(.*?)\}", re.DOTALL)
_RE_FONT = re.compile(
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

# Color table patterns
_RE_COLORTBL = re.compile(r"\{\\colortbl\s*;?(.*?)\}", re.DOTALL)
_RE_COLOR = re.compile(r"\\red(\d+)\\green(\d+)\\blue(\d+)", re.IGNORECASE)

# Stylesheet pattern
_RE_STYLESHEET = re.compile(r"\{\\stylesheet(.*?)\}\s*(?=\{|\\)", re.DOTALL)
_RE_STYLE = re.compile(
    r"\{(?:\\[*]?)?\\s(\d+)"  # Style ID
    r"[^}]*?"  # Style properties
    r"(?:\\sbasedon(\d+))?"  # Based on
    r"(?:\\snext(\d+))?"  # Next style
    r"[^}]*?"
    r"\s+([^;}]+)"  # Style name
    r";?\}",
    re.IGNORECASE,
)

# Metadata patterns
_RE_INFO = re.compile(r"\{\\info\s*(.*?)\}(?=\s*\\|\s*\{[^\\])", re.DOTALL)
_RE_INFO_ALT = re.compile(r"\{\\info\s*((?:\{[^}]*\}\s*)+)\}", re.DOTALL)
_RE_HEX_ESCAPE = re.compile(r"\\'([0-9a-fA-F]{2})")
_RE_CONTROL_SEQ = re.compile(r"\\[a-z]+\d*\s*", re.IGNORECASE)

# Date component patterns
_RE_YEAR = re.compile(r"\\yr(\d+)")
_RE_MONTH = re.compile(r"\\mo(\d+)")
_RE_DAY = re.compile(r"\\dy(\d+)")
_RE_HOUR = re.compile(r"\\hr(\d+)")
_RE_MINUTE = re.compile(r"\\min(\d+)")

# Hyperlink and field patterns
_RE_HYPERLINK = re.compile(
    r'\\field\s*\{[^}]*\\fldinst\s*\{[^}]*HYPERLINK\s+"([^"]+)"[^}]*\}'
    r"[^}]*\{[^}]*\\fldrslt\s*\{([^}]*)\}",
    re.IGNORECASE | re.DOTALL,
)
_RE_FIELD = re.compile(
    r"\\field\s*\{[^}]*\\fldinst\s*\{([^}]*)\}[^}]*\{[^}]*\\fldrslt\s*\{([^}]*)\}",
    re.IGNORECASE | re.DOTALL,
)

# Image and table patterns
_RE_PICT = re.compile(
    r"\{\\pict([^}]*(?:\{[^}]*\}[^}]*)*)\}", re.DOTALL | re.IGNORECASE
)
_RE_PAGE_BREAK = re.compile(r"\\page\b")
_RE_TROWD = re.compile(r"\\trowd\b")
_RE_ROW = re.compile(r"\\row\b")
_RE_FOOTNOTE = re.compile(
    r"\{\\footnote\s*([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}",
    re.DOTALL | re.IGNORECASE,
)

# Image dimension patterns
_RE_PICW = re.compile(r"\\picw(\d+)")
_RE_PICH = re.compile(r"\\pich(\d+)")
_RE_HEX_DATA = re.compile(r"([0-9a-fA-F]{20,})")

# Text cleaning patterns
_RE_UNICODE = re.compile(r"\\u(-?\d+)\??")
_RE_CONTROL_WORD = re.compile(r"\\[a-z]+(-?\d+)?\s?", re.IGNORECASE)
_RE_MULTI_SPACE = re.compile(r"[ \t]+")
_RE_MULTI_NEWLINE = re.compile(r"\n{3,}")
_RE_HEX_RUN = re.compile(r"[0-9a-fA-F]{64,}")
_RE_CELL_SPACE = re.compile(r"[ \t\f\v]+")
_RE_CELL_NEWLINE = re.compile(r" *\n *")

# Destination removal patterns (pre-compiled for body extraction)
_DEST_PATTERNS = [
    re.compile(r"\{\\fonttbl[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", re.DOTALL | re.IGNORECASE),
    re.compile(r"\{\\colortbl[^{}]*\}", re.DOTALL | re.IGNORECASE),
    re.compile(
        r"\{\\stylesheet[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", re.DOTALL | re.IGNORECASE
    ),
    re.compile(r"\{\\info[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", re.DOTALL | re.IGNORECASE),
    re.compile(
        r"\{\\[*]?\\?header[lrf]?\s+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
        re.DOTALL | re.IGNORECASE,
    ),
    re.compile(
        r"\{\\[*]?\\?footer[lrf]?\s+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}",
        re.DOTALL | re.IGNORECASE,
    ),
    re.compile(
        r"\{\\\*\\[a-z]+[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", re.DOTALL | re.IGNORECASE
    ),
]


# =============================================================================
# Helper functions
# =============================================================================


def _get_page_for_position(position: int, page_breaks: list[int]) -> int:
    """Determine page number (1-based) for a given position."""
    page = 1
    for bp in page_breaks:
        if bp < position:
            page += 1
        else:
            break
    return page


def _parse_rtf_date(date_str: str) -> str:
    """Parse RTF date components into ISO format."""
    yr = _RE_YEAR.search(date_str)
    mo = _RE_MONTH.search(date_str)
    dy = _RE_DAY.search(date_str)

    if not (yr and mo and dy):
        return ""

    year = int(yr.group(1))
    month = int(mo.group(1))
    day = int(dy.group(1))

    hr = _RE_HOUR.search(date_str)
    mn = _RE_MINUTE.search(date_str)
    hour = int(hr.group(1)) if hr else 0
    minute = int(mn.group(1)) if mn else 0

    return f"{year:04d}-{month:02d}-{day:02d}T{hour:02d}:{minute:02d}:00"


# =============================================================================
# RTF Parser
# =============================================================================


class _RtfParser:
    """Low-level RTF parser for text and metadata extraction."""

    # Destinations to skip during text extraction
    SKIP_DESTINATIONS = frozenset(
        {
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
        }
    )

    # RTF special characters mapping
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
        "~": "\u00a0",
        "_": "\u00ad",
        "-": "\u00ad",
        "enspace": "\u2002",
        "emspace": "\u2003",
        "qmspace": "\u2005",
    }

    def __init__(self, data: bytes):
        self.data = data

        # Extracted data
        self.fonts: list[RtfFont] = []
        self.colors: list[RtfColor] = []
        self.styles: list[RtfStyle] = []
        self.metadata = RtfMetadata()
        self.paragraphs: list[RtfParagraph] = []
        self.headers: list[RtfHeaderFooter] = []
        self.footers: list[RtfHeaderFooter] = []
        self.hyperlinks: list[RtfHyperlink] = []
        self.bookmarks: list[RtfBookmark] = []
        self.fields: list[RtfField] = []
        self.images: list[RtfImage] = []
        self.tables: list[RtfTable] = []
        self.footnotes: list[RtfFootnote] = []
        self.annotations: list[RtfAnnotation] = []
        self.pages: list[str] = []
        self.raw_text_blocks: list[str] = []

    def parse(self) -> RtfContent:
        """Parse the RTF document and return extracted content."""
        try:
            text = self._decode_rtf()

            if not text.startswith("{\\rtf"):
                logger.warning("Not a valid RTF file - missing RTF header")
                return RtfContent(full_text=text)

            self._extract_fonts(text)
            self._extract_colors(text)
            self._extract_styles(text)
            self._extract_metadata(text)
            self._extract_headers_footers(text)
            self._extract_body_text(text)
            self._extract_hyperlinks(text)
            self._extract_fields(text)
            self._extract_images(text)
            self._extract_tables(text)
            self._extract_footnotes(text)

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
                tables=self.tables,
                footnotes=self.footnotes,
                annotations=self.annotations,
                pages=self.pages,
                full_text=full_text,
                raw_text_blocks=self.raw_text_blocks,
            )
        except Exception as e:
            logger.error(f"RTF parsing failed: {e}")
            try:
                text = self._decode_rtf()
                plain = self._strip_rtf_simple(text)
                return RtfContent(full_text=plain)
            except Exception:
                return RtfContent()

    def _decode_rtf(self) -> str:
        """Decode RTF bytes to string, handling various encodings."""
        for encoding in ("utf-8", "cp1252", "latin-1"):
            try:
                return self.data.decode(encoding)
            except UnicodeDecodeError:
                continue
        return self.data.decode("cp1252", errors="replace")

    def _extract_fonts(self, text: str) -> None:
        """Extract font table from RTF."""
        match = _RE_FONTTBL.search(text)
        if not match:
            return

        for m in _RE_FONT.finditer(match.group(1)):
            self.fonts.append(
                RtfFont(
                    font_id=int(m.group(1)),
                    font_family=m.group(2) or "",
                    font_name=m.group(5).strip(),
                    charset=int(m.group(3)) if m.group(3) else 0,
                    pitch=int(m.group(4)) if m.group(4) else 0,
                )
            )

    def _extract_colors(self, text: str) -> None:
        """Extract color table from RTF."""
        match = _RE_COLORTBL.search(text)
        if not match:
            return

        for idx, m in enumerate(_RE_COLOR.finditer(match.group(1))):
            self.colors.append(
                RtfColor(
                    index=idx,
                    red=int(m.group(1)),
                    green=int(m.group(2)),
                    blue=int(m.group(3)),
                )
            )

    def _extract_styles(self, text: str) -> None:
        """Extract stylesheet from RTF."""
        match = _RE_STYLESHEET.search(text)
        if not match:
            return

        for m in _RE_STYLE.finditer(match.group(1)):
            style_type = "paragraph"
            if "\\cs" in m.group(0):
                style_type = "character"
            elif "\\ts" in m.group(0):
                style_type = "table"

            self.styles.append(
                RtfStyle(
                    style_id=int(m.group(1)),
                    style_type=style_type,
                    style_name=m.group(4).strip(),
                    based_on=int(m.group(2)) if m.group(2) else None,
                    next_style=int(m.group(3)) if m.group(3) else None,
                )
            )

    def _extract_metadata(self, text: str) -> None:
        """Extract document metadata from RTF."""
        match = _RE_INFO.search(text)
        if not match:
            match = _RE_INFO_ALT.search(text)
        if not match:
            return

        info = match.group(1)

        def get_value(pattern: str) -> str:
            m = re.search(pattern, info, re.IGNORECASE | re.DOTALL)
            if not m:
                return ""
            val = m.group(1).strip()
            val = _RE_HEX_ESCAPE.sub(lambda x: chr(int(x.group(1), 16)), val)
            val = val.replace("\\~", " ")
            val = _RE_CONTROL_SEQ.sub("", val)
            return val.replace("{", "").replace("}", "").strip()

        # String metadata
        self.metadata.title = get_value(r"\{\\title\s+([^}]*)\}")
        self.metadata.subject = get_value(r"\{\\subject\s+([^}]*)\}")
        self.metadata.author = get_value(r"\{\\author\s+([^}]*)\}")
        self.metadata.keywords = get_value(r"\{\\keywords\s+([^}]*)\}")
        self.metadata.comments = get_value(r"\{\\comment\s+([^}]*)\}")
        self.metadata.operator = get_value(r"\{\\operator\s+([^}]*)\}")
        self.metadata.category = get_value(r"\{\\[*]?\\?category\s+([^}]*)\}")
        self.metadata.manager = get_value(r"\{\\manager\s+([^}]*)\}")
        self.metadata.company = get_value(r"\{\\company\s+([^}]*)\}")
        self.metadata.doc_comment = get_value(r"\{\\doccomm\s+([^}]*)\}")

        # Numeric metadata
        for pattern, attr in [
            (r"\{\\version(\d+)\}", "version"),
            (r"\{\\vern(\d+)\}", "revision"),
            (r"\{\\nofpages(\d+)\}", "num_pages"),
            (r"\{\\nofwords(\d+)\}", "num_words"),
            (r"\{\\nofchars(\d+)\}", "num_chars"),
            (r"\{\\nofcharsws(\d+)\}", "num_chars_with_spaces"),
        ]:
            m = re.search(pattern, info)
            if m:
                setattr(self.metadata, attr, int(m.group(1)))

        # Dates
        for pattern, attr in [
            (r"\{\\creatim([^}]*)\}", "created"),
            (r"\{\\revtim([^}]*)\}", "modified"),
        ]:
            m = re.search(pattern, info, re.IGNORECASE)
            if m:
                setattr(self.metadata, attr, _parse_rtf_date(m.group(1)))

    def _extract_headers_footers(self, text: str) -> None:
        """Extract headers and footers from RTF."""
        hf_types = [
            "header",
            "headerl",
            "headerr",
            "headerf",
            "footer",
            "footerl",
            "footerr",
            "footerf",
        ]

        for hf_type in hf_types:
            pattern = re.compile(
                r"\{\\" + hf_type + r"\s+(.*?)\}(?=\s*\{|\s*\\)",
                re.DOTALL | re.IGNORECASE,
            )
            for match in pattern.finditer(text):
                extracted = self._strip_rtf_simple(match.group(1)).strip()
                if extracted:
                    hf = RtfHeaderFooter(type=hf_type, text=extracted)
                    if "header" in hf_type:
                        self.headers.append(hf)
                    else:
                        self.footers.append(hf)

    def _extract_body_text(self, text: str) -> None:
        """Extract body text from RTF with page break detection."""
        body = text
        for pattern in _DEST_PATTERNS:
            body = pattern.sub("", body)

        extracted = self._strip_rtf_full_with_pages(body)

        for para_text in extracted.split("\n"):
            cleaned = para_text.strip()
            if cleaned:
                self.paragraphs.append(RtfParagraph(text=cleaned))
                self.raw_text_blocks.append(cleaned)

    def _extract_hyperlinks(self, text: str) -> None:
        """Extract hyperlinks from RTF."""
        for match in _RE_HYPERLINK.finditer(text):
            url = match.group(1).strip()
            link_text = self._strip_rtf_simple(match.group(2)).strip()
            if url:
                self.hyperlinks.append(RtfHyperlink(text=link_text, url=url))

    def _extract_fields(self, text: str) -> None:
        """Extract fields (page numbers, dates, etc.) from RTF."""
        for match in _RE_FIELD.finditer(text):
            instruction = self._strip_rtf_simple(match.group(1)).strip()
            result = self._strip_rtf_simple(match.group(2)).strip()

            if "HYPERLINK" in instruction.upper():
                continue

            instr_upper = instruction.upper()
            field_type = "unknown"
            if "PAGE" in instr_upper:
                field_type = "page"
            elif "DATE" in instr_upper:
                field_type = "date"
            elif "TIME" in instr_upper:
                field_type = "time"
            elif "STYLEREF" in instr_upper:
                field_type = "styleref"
            elif "TOC" in instr_upper:
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
        page_breaks = [m.start() for m in _RE_PAGE_BREAK.finditer(text)]

        for idx, match in enumerate(_RE_PICT.finditer(text), 1):
            content = match.group(1)
            page_num = _get_page_for_position(match.start(), page_breaks)

            # Determine image type
            img_type = "unknown"
            if "\\pngblip" in content:
                img_type = "png"
            elif "\\jpegblip" in content:
                img_type = "jpeg"
            elif "\\emfblip" in content:
                img_type = "emf"
            elif "\\wmetafile" in content:
                img_type = "wmf"

            # Extract dimensions
            w_match = _RE_PICW.search(content)
            h_match = _RE_PICH.search(content)
            width = int(w_match.group(1)) if w_match else 0
            height = int(h_match.group(1)) if h_match else 0

            # Extract hex data
            data = None
            hex_match = _RE_HEX_DATA.search(content)
            if hex_match:
                try:
                    data = bytes.fromhex(hex_match.group(1))
                except ValueError:
                    pass

            self.images.append(
                RtfImage(
                    image_type=img_type,
                    width=width,
                    height=height,
                    data=data,
                    image_index=idx,
                    page_number=page_num,
                )
            )

        if self.images:
            logger.debug(f"Extracted {len(self.images)} images from RTF")

    def _extract_tables(self, text: str) -> None:
        """Extract tables from RTF."""
        page_breaks = [m.start() for m in _RE_PAGE_BREAK.finditer(text)]
        trowd_positions = [m.start() for m in _RE_TROWD.finditer(text)]
        row_positions = [m.end() for m in _RE_ROW.finditer(text)]

        if not trowd_positions or not row_positions:
            return

        # Match \trowd with corresponding \row
        table_rows: list[tuple[int, int, str]] = []
        for trowd_pos in trowd_positions:
            for rpos in row_positions:
                if rpos > trowd_pos:
                    table_rows.append((trowd_pos, rpos, text[trowd_pos:rpos]))
                    break

        # Group consecutive rows into tables
        current_rows: list[list[str]] = []
        current_start: int | None = None
        last_end = -1
        table_idx = 0

        for row_start, row_end, row_content in table_rows:
            # Check for table break (significant gap between rows)
            if current_rows and row_start - last_end > 100:
                between = self._strip_rtf_simple(text[last_end:row_start]).strip()
                if len(between) > 20:
                    if current_rows:
                        self._save_table(
                            current_rows, current_start or 0, page_breaks, table_idx + 1
                        )
                        table_idx += 1
                    current_rows = []
                    current_start = None

            cells = self._extract_table_cells(row_content)
            if cells:
                if current_start is None:
                    current_start = row_start
                current_rows.append(cells)
            last_end = row_end

        # Save last table
        if current_rows:
            self._save_table(
                current_rows, current_start or 0, page_breaks, table_idx + 1
            )

        if self.tables:
            logger.debug(f"Extracted {len(self.tables)} tables from RTF")

    def _save_table(
        self, rows: list[list[str]], start_pos: int, page_breaks: list[int], idx: int
    ) -> None:
        """Normalize and save a table."""
        if not rows:
            return
        max_cols = max(len(r) for r in rows)
        for row in rows:
            row.extend([""] * (max_cols - len(row)))
        self.tables.append(
            RtfTable(
                data=rows,
                table_index=idx,
                page_number=_get_page_for_position(start_pos, page_breaks),
            )
        )

    def _extract_table_cells(self, row_content: str) -> list[str]:
        """Extract cell contents from a table row."""
        cells = []
        for part in re.split(r"\\cell\b", row_content)[:-1]:
            cell_text = self._strip_rtf_simple(part)
            cell_text = _RE_HEX_RUN.sub("", cell_text)
            cell_text = _RE_CELL_SPACE.sub(" ", cell_text)
            cell_text = _RE_CELL_NEWLINE.sub("\n", cell_text)
            cell_text = _RE_MULTI_NEWLINE.sub("\n\n", cell_text)
            cells.append(cell_text.strip())
        return cells

    def _extract_footnotes(self, text: str) -> None:
        """Extract footnotes from RTF."""
        for idx, match in enumerate(_RE_FOOTNOTE.finditer(text), 1):
            footnote_text = self._strip_rtf_simple(match.group(1)).strip()
            if footnote_text:
                self.footnotes.append(RtfFootnote(id=idx, text=footnote_text))

    def _remove_ignorable_groups(self, text: str) -> str:
        """Remove ignorable/binary groups for simple text extraction."""
        lower = text.lower()
        prefixes = ("{\\pict", "{\\object", "{\\*")

        out: list[str] = []
        i = 0
        n = len(text)

        while i < n:
            if text[i] != "{":
                out.append(text[i])
                i += 1
                continue

            if any(lower.startswith(pfx, i) for pfx in prefixes):
                depth = 0
                while i < n:
                    if text[i] == "{":
                        depth += 1
                    elif text[i] == "}":
                        depth -= 1
                        if depth == 0:
                            i += 1
                            break
                    i += 1
                continue

            out.append(text[i])
            i += 1

        return "".join(out)

    def _strip_rtf_simple(self, text: str) -> str:
        """Simple RTF stripping - removes control words and groups."""
        result = self._remove_ignorable_groups(text)

        # Unicode escapes
        result = _RE_UNICODE.sub(lambda m: chr(int(m.group(1)) & 0xFFFF), result)

        # Hex escapes
        result = _RE_HEX_ESCAPE.sub(lambda m: chr(int(m.group(1), 16)), result)

        # Special characters
        for keyword, char in self.SPECIAL_CHARS.items():
            pattern = r"\\" + re.escape(keyword) + r"(?:(?:\s+)|(?=\\)|(?=\{)|(?=\})|$)"
            result = re.sub(pattern, char, result)

        # Control words
        result = _RE_CONTROL_WORD.sub("", result)

        # Braces and whitespace
        result = result.replace("{", "").replace("}", "")
        result = _RE_MULTI_SPACE.sub(" ", result)
        result = _RE_MULTI_NEWLINE.sub("\n\n", result)

        return result.strip()

    def _is_skip_destination(self, ahead: str) -> bool:
        """Check if lookahead indicates a destination to skip."""
        if ahead.startswith("\\*"):
            return True
        for kw in self.SKIP_DESTINATIONS:
            if ahead.startswith("\\" + kw):
                return True
        return False

    def _strip_rtf_full_with_pages(self, text: str) -> str:
        """Full RTF stripping with page break detection."""
        result: list[str] = []
        current_page: list[str] = []
        i = 0
        n = len(text)
        group_depth = 0
        skip_group = False
        skip_depth = 0

        def flush_page() -> None:
            page_text = "".join(current_page).strip()
            page_text = _RE_MULTI_SPACE.sub(" ", page_text)
            page_text = _RE_MULTI_NEWLINE.sub("\n\n", page_text)
            if page_text:
                self.pages.append(page_text)
            current_page.clear()

        while i < n:
            char = text[i]

            if char == "{":
                group_depth += 1
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
                if i + 1 >= n:
                    i += 1
                    continue

                next_char = text[i + 1]

                if next_char in "\\{}":
                    current_page.append(next_char)
                    result.append(next_char)
                    i += 2

                elif next_char == "u":
                    m = _RE_UNICODE.match(text, i)
                    if m:
                        char_val = chr(int(m.group(1)) & 0xFFFF)
                        current_page.append(char_val)
                        result.append(char_val)
                        i += len(m.group(0))
                    else:
                        i += 2

                elif next_char == "'":
                    if i + 3 < n:
                        try:
                            char_val = chr(int(text[i + 2 : i + 4], 16))
                            current_page.append(char_val)
                            result.append(char_val)
                        except ValueError:
                            pass
                        i += 4
                    else:
                        i += 2

                elif next_char.isalpha():
                    j = i + 1
                    while j < n and text[j].isalpha():
                        j += 1
                    if j < n and (text[j].isdigit() or text[j] == "-"):
                        while j < n and (text[j].isdigit() or text[j] == "-"):
                            j += 1
                    if j < n and text[j] == " ":
                        j += 1

                    control_word = text[i + 1 : j].rstrip()
                    word_only = re.sub(r"-?\d+$", "", control_word)

                    if word_only in ("page", "sbkpage"):
                        flush_page()
                    elif word_only in self.SPECIAL_CHARS:
                        char_val = self.SPECIAL_CHARS[word_only]
                        current_page.append(char_val)
                        result.append(char_val)

                    i = j

                else:
                    if next_char == "~":
                        current_page.append("\u00a0")
                        result.append("\u00a0")
                    elif next_char == "_":
                        current_page.append("\u00ad")
                        result.append("\u00ad")
                    i += 2

            else:
                if char != "\r":
                    current_page.append(char)
                    result.append(char)
                i += 1

        flush_page()

        if not self.pages:
            full_text = "".join(result).strip()
            full_text = _RE_MULTI_SPACE.sub(" ", full_text)
            full_text = _RE_MULTI_NEWLINE.sub("\n\n", full_text)
            if full_text:
                self.pages.append(full_text)

        return "".join(result)


# =============================================================================
# Main entry point
# =============================================================================


def read_rtf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[RtfContent, Any, None]:
    """
    Extract content from an RTF file.

    Uses a generator pattern for API consistency. RTF files yield exactly one
    RtfContent object containing text, pages, metadata, and document elements.
    """
    try:
        logger.debug("Reading RTF file")
        file_like.seek(0)
        data = file_like.read()

        parser = _RtfParser(data)
        content = parser.parse()
        content.metadata.populate_from_path(path)

        yield content
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract RTF file", cause=exc) from exc
