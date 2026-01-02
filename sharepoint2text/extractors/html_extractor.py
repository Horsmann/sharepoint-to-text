"""
HTML Content Extractor
======================

Extracts text content, metadata, and structure from HTML documents using
Python's standard library html.parser for parsing.

File Format Background
----------------------
HTML (HyperText Markup Language) is the standard markup language for web
documents. This extractor handles:
    - HTML5 documents with modern semantic elements
    - XHTML documents with strict XML syntax
    - Legacy HTML (4.x and earlier) with lenient parsing
    - HTML fragments without full document structure

Encoding Detection
------------------
The extractor detects character encoding in priority order:
    1. BOM (Byte Order Mark): UTF-8, UTF-16 LE/BE
    2. Meta charset tag: <meta charset="...">
    3. Content-Type meta: <meta http-equiv="Content-Type" content="...;charset=...">
    4. Fallback: UTF-8 with error replacement

Element Handling
----------------
REMOVE_TAGS: Elements removed entirely (including content):
    - script: JavaScript code
    - style: CSS stylesheets
    - noscript: No-script fallback content
    - iframe, object, embed, applet: Embedded content

BLOCK_TAGS: Block-level elements that receive newline formatting:
    - Structural: div, section, article, header, footer, nav, aside, main
    - Headings: h1-h6
    - Text blocks: p, blockquote, pre, address
    - Figures: figure, figcaption
    - Lists: ul, ol, li, dl, dt, dd
    - Tables: table, tr
    - Misc: hr, br, form, fieldset

Dependencies
------------
No external dependencies. Uses Python standard library:
    - html.parser: HTML parsing
    - html.entities: HTML entity decoding

Extracted Content
-----------------
The extractor produces:
    - content: Full text with preserved structure (newlines, lists, tables)
    - tables: Nested lists [table[row[cell]]] for each table
    - headings: List of {level, text} for h1-h6 elements
    - links: List of {text, href} for anchor elements
    - metadata: Title, charset, language, description, keywords, author

Table Formatting
----------------
Tables are extracted both as structured data and formatted text:
    - Structured: List of rows, each row a list of cell strings
    - Text: Column-aligned with pipe separators for readability

Known Limitations
-----------------
- JavaScript-generated content is not captured (no JS execution)
- CSS-hidden content is extracted (no CSS interpretation)
- Very deeply nested structures may lose some formatting
- Inline SVG content is not specially handled
- Form field values are not extracted
- Shadow DOM content is not accessible

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.html_extractor import read_html
    >>>
    >>> with open("page.html", "rb") as f:
    ...     for doc in read_html(io.BytesIO(f.read()), path="page.html"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(f"Tables: {len(doc.tables)}")
    ...         print(doc.content[:500])

See Also
--------
- HTML Living Standard: https://html.spec.whatwg.org/

Maintenance Notes
-----------------
- Uses html.parser from stdlib for parsing
- Parser is lenient and handles malformed HTML
- Tables are processed separately to preserve structure
- List items are indented with bullet markers for readability
- Multiple consecutive newlines are collapsed to maximum of 2
"""

import io
import logging
import re
from html.parser import HTMLParser
from typing import Any, Dict, Generator, List, Optional, Tuple

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.extractors.data_types import (
    HtmlContent,
    HtmlMetadata,
)

logger = logging.getLogger(__name__)

# Tags to remove entirely (including their content)
REMOVE_TAGS = {"script", "style", "noscript", "iframe", "object", "embed", "applet"}

# Block-level tags that should have newlines around them
BLOCK_TAGS = {
    "p",
    "div",
    "section",
    "article",
    "header",
    "footer",
    "nav",
    "aside",
    "main",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "blockquote",
    "pre",
    "address",
    "figure",
    "figcaption",
    "form",
    "fieldset",
    "ul",
    "ol",
    "li",
    "dl",
    "dt",
    "dd",
    "table",
    "tr",
    "hr",
    "br",
}


class _HtmlTreeBuilder(HTMLParser):
    """
    Build a simple tree structure from HTML for processing.

    Creates a lightweight DOM-like structure that can be traversed
    for text extraction without external dependencies.

    Each node has:
        - tag: element tag name
        - attrs: dictionary of attributes
        - children: list of child nodes
        - text: text content before children
        - tail: text content after this element (before next sibling)
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        # Root node
        self.root: Dict = {
            "tag": "root",
            "attrs": {},
            "children": [],
            "text": "",
            "tail": "",
        }
        # Stack for building tree
        self.stack: List[Dict] = [self.root]
        # Track if we're inside a tag whose content should be ignored
        self.skip_depth = 0
        # Track the last closed element for tail text
        self.last_closed: Optional[Dict] = None

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]):
        tag = tag.lower()
        attrs_dict = {k: v for k, v in attrs if v is not None}

        node = {"tag": tag, "attrs": attrs_dict, "children": [], "text": "", "tail": ""}

        if self.skip_depth > 0:
            self.skip_depth += 1
            return

        if tag in REMOVE_TAGS:
            self.skip_depth = 1
            return

        # Clear last_closed since we're starting a new element
        self.last_closed = None

        # Add to parent's children
        self.stack[-1]["children"].append(node)
        # Push onto stack (for non-void elements)
        if tag not in {
            "br",
            "hr",
            "img",
            "input",
            "meta",
            "link",
            "area",
            "base",
            "col",
            "embed",
            "param",
            "source",
            "track",
            "wbr",
        }:
            self.stack.append(node)
        else:
            # For void elements, they become the "last closed" element
            self.last_closed = node

    def handle_endtag(self, tag: str):
        tag = tag.lower()

        if self.skip_depth > 0:
            self.skip_depth -= 1
            return

        # Pop from stack if this closes the current element
        if len(self.stack) > 1 and self.stack[-1]["tag"] == tag:
            self.last_closed = self.stack.pop()

    def handle_data(self, data: str):
        if self.skip_depth > 0:
            return

        if self.last_closed is not None:
            # Text after a closed element goes to its tail
            self.last_closed["tail"] += data
        elif self.stack:
            # Text inside an element goes to its text
            self.stack[-1]["text"] += data

    def handle_comment(self, data: str):
        # Ignore comments
        pass

    def get_tree(self) -> Dict:
        return self.root


class _HtmlTextExtractor:
    """Helper class to extract text from the parsed HTML tree."""

    def __init__(self, root: Dict):
        self.root = root
        self.tables: List[List[List[str]]] = []
        self.headings: List[Dict[str, str]] = []
        self.links: List[Dict[str, str]] = []
        self.metadata = HtmlMetadata()

    def _get_node_text(
        self, node: Dict, include_children: bool = True, include_tail: bool = False
    ) -> str:
        """Get all text content from a node and its descendants."""
        parts = []
        if node.get("text"):
            parts.append(node["text"])

        if include_children:
            for child in node.get("children", []):
                parts.append(self._get_node_text(child, True, include_tail=True))

        # Include tail text if requested (text after this element)
        if include_tail and node.get("tail"):
            parts.append(node["tail"])

        return "".join(parts)

    def _find_nodes(self, node: Dict, tag: str) -> List[Dict]:
        """Find all descendant nodes with the given tag."""
        result = []
        if node.get("tag") == tag:
            result.append(node)
        for child in node.get("children", []):
            result.extend(self._find_nodes(child, tag))
        return result

    def _find_node(self, node: Dict, tag: str) -> Optional[Dict]:
        """Find first descendant node with the given tag."""
        if node.get("tag") == tag:
            return node
        for child in node.get("children", []):
            found = self._find_node(child, tag)
            if found:
                return found
        return None

    def _extract_table(self, table_node: Dict) -> List[List[str]]:
        """Extract table as a list of rows, each row being a list of cell values."""
        rows = []
        for tr in self._find_nodes(table_node, "tr"):
            row = []
            for child in tr.get("children", []):
                if child.get("tag") in ("th", "td"):
                    cell_text = self._get_node_text(child).strip()
                    cell_text = re.sub(r"\s+", " ", cell_text)
                    row.append(cell_text)
            if row:
                rows.append(row)
        return rows

    def _format_table_as_text(self, table_data: List[List[str]]) -> str:
        """Format a table as readable text with aligned columns."""
        if not table_data:
            return ""

        # Calculate column widths
        num_cols = max(len(row) for row in table_data) if table_data else 0
        col_widths = [0] * num_cols

        for row in table_data:
            for i, cell in enumerate(row):
                if i < num_cols:
                    col_widths[i] = max(col_widths[i], len(cell))

        # Build text representation
        lines = []
        for row in table_data:
            # Pad row to full width
            padded_row = row + [""] * (num_cols - len(row))
            formatted_cells = []
            for i, cell in enumerate(padded_row):
                formatted_cells.append(cell.ljust(col_widths[i]))
            lines.append(" | ".join(formatted_cells).rstrip())

        return "\n".join(lines)

    def _extract_headings(self) -> None:
        """Extract all headings with their level."""
        for level in range(1, 7):
            for h in self._find_nodes(self.root, f"h{level}"):
                text = self._get_node_text(h).strip()
                text = re.sub(r"\s+", " ", text)
                if text:
                    self.headings.append({"level": f"h{level}", "text": text})

    def _extract_links(self) -> None:
        """Extract all links with their text and href."""
        for a in self._find_nodes(self.root, "a"):
            href = a.get("attrs", {}).get("href", "")
            text = self._get_node_text(a).strip()
            text = re.sub(r"\s+", " ", text)
            if href and text:
                self.links.append({"text": text, "href": href})

    def _extract_metadata(self, path: Optional[str]) -> None:
        """Extract metadata from HTML document."""
        self.metadata.populate_from_path(path)

        # Extract title
        title_node = self._find_node(self.root, "title")
        if title_node:
            self.metadata.title = self._get_node_text(title_node).strip()

        # Extract language from html tag
        html_node = self._find_node(self.root, "html")
        if html_node:
            lang = html_node.get("attrs", {}).get("lang", "")
            if lang:
                self.metadata.language = lang

        # Extract from meta tags
        for meta in self._find_nodes(self.root, "meta"):
            attrs = meta.get("attrs", {})

            # charset
            if "charset" in attrs:
                self.metadata.charset = attrs["charset"]

            # http-equiv content-type
            if attrs.get("http-equiv", "").lower() == "content-type":
                content = attrs.get("content", "")
                match = re.search(r"charset=([^\s;]+)", content)
                if match:
                    self.metadata.charset = match.group(1)

            # name-based meta tags
            name = attrs.get("name", "").lower()
            content = attrs.get("content", "")
            if name == "description" and content:
                self.metadata.description = content
            elif name == "keywords" and content:
                self.metadata.keywords = content
            elif name == "author" and content:
                self.metadata.author = content

    def _process_node(
        self, node: Dict, depth: int = 0, include_tail: bool = False
    ) -> str:
        """Recursively process a node and return its text representation."""
        tag = node.get("tag", "")

        if tag in REMOVE_TAGS:
            return ""

        result_parts = []

        # Handle tables specially
        if tag == "table":
            table_data = self._extract_table(node)
            self.tables.append(table_data)
            table_text = self._format_table_as_text(table_data)
            result = "\n" + table_text + "\n"
            # Add tail text
            if include_tail and node.get("tail"):
                result += node["tail"]
            return result

        # Handle list items
        if tag == "li":
            text = node.get("text", "")
            for child in node.get("children", []):
                text += self._process_node(child, depth + 1, include_tail=True)
            text = re.sub(r"\s+", " ", text.strip())
            result = "  " * depth + "- " + text + "\n"
            # Add tail text
            if include_tail and node.get("tail"):
                result += node["tail"]
            return result

        # Handle headings
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            text = self._get_node_text(node).strip()
            text = re.sub(r"\s+", " ", text)
            result = "\n" + text + "\n"
            # Add tail text
            if include_tail and node.get("tail"):
                result += node["tail"]
            return result

        # Handle line breaks
        if tag == "br":
            result = "\n"
            if include_tail and node.get("tail"):
                result += node["tail"]
            return result

        # Handle horizontal rules
        if tag == "hr":
            result = "\n---\n"
            if include_tail and node.get("tail"):
                result += node["tail"]
            return result

        # Add node's own text
        if node.get("text"):
            result_parts.append(node["text"])

        # Process children (include tail for all children)
        for child in node.get("children", []):
            child_text = self._process_node(child, depth, include_tail=True)
            result_parts.append(child_text)

        result = "".join(result_parts)

        # Add newlines around block elements
        if tag in BLOCK_TAGS:
            result = "\n" + result.strip() + "\n"

        # Add tail text for this node
        if include_tail and node.get("tail"):
            result += node["tail"]

        return result

    def extract(self, path: Optional[str] = None) -> str:
        """Extract and return the full text content."""
        self._extract_metadata(path)
        self._extract_headings()
        self._extract_links()

        # Find the body or use root
        body = self._find_node(self.root, "body")
        if body:
            text = self._process_node(body)
        else:
            text = self._process_node(self.root)

        # Clean up the text
        # Collapse multiple newlines to maximum of 2
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Remove leading/trailing whitespace from lines
        lines = [line.strip() for line in text.split("\n")]
        text = "\n".join(lines)
        # Remove leading/trailing newlines
        text = text.strip()

        return text


def read_html(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[HtmlContent, Any, None]:
    """
    Extract all relevant content from an HTML document.

    Primary entry point for HTML extraction. Detects encoding, parses the
    document with Python's html.parser, removes unwanted elements (scripts,
    styles), and extracts text with preserved structure.

    This function uses a generator pattern for API consistency with other
    extractors, even though HTML files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete HTML file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned HtmlContent.metadata.

    Yields:
        HtmlContent: Single HtmlContent object containing:
            - content: Full extracted text with formatting
            - tables: Structured table data as nested lists
            - headings: List of heading elements with level
            - links: List of anchor elements with href
            - metadata: HtmlMetadata with title, charset, etc.

    Note:
        On parse failure, yields empty HtmlContent rather than raising.
        A warning is logged when parsing fails completely.

    Example:
        >>> import io
        >>> with open("report.html", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_html(data, path="report.html"):
        ...         print(f"Title: {doc.metadata.title}")
        ...         for heading in doc.headings:
        ...             print(f"  {heading['level']}: {heading['text']}")
    """
    try:
        logger.debug("Reading HTML file")
        file_like.seek(0)

        content = file_like.read()

        # Detect encoding from content if possible, default to utf-8
        encoding = "utf-8"

        # Try to detect encoding from meta tag or BOM
        if content.startswith(b"\xef\xbb\xbf"):
            encoding = "utf-8"
            content = content[3:]
        elif content.startswith(b"\xff\xfe"):
            encoding = "utf-16-le"
        elif content.startswith(b"\xfe\xff"):
            encoding = "utf-16-be"
        else:
            # Try to find charset in meta tag
            charset_match = re.search(
                rb'<meta[^>]+charset=["\']?([^"\'\s>]+)', content, re.IGNORECASE
            )
            if charset_match:
                encoding = charset_match.group(1).decode("ascii", errors="ignore")

        try:
            html_text = content.decode(encoding, errors="replace")
        except (UnicodeDecodeError, LookupError):
            html_text = content.decode("utf-8", errors="replace")

        # Parse the HTML
        try:
            parser = _HtmlTreeBuilder()
            parser.feed(html_text)
            root = parser.get_tree()
        except Exception:
            # Last resort: return empty content
            logger.warning("Failed to parse HTML content")
            metadata = HtmlMetadata()
            metadata.populate_from_path(path)
            yield HtmlContent(content="", metadata=metadata)
            return

        # Extract text content
        extractor = _HtmlTextExtractor(root)
        text = extractor.extract(path)

        logger.info(
            "Extracted HTML: %d characters, %d tables, %d links",
            len(text),
            len(extractor.tables),
            len(extractor.links),
        )

        yield HtmlContent(
            content=text,
            tables=extractor.tables,
            headings=extractor.headings,
            links=extractor.links,
            metadata=extractor.metadata,
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract HTML file", cause=exc) from exc
