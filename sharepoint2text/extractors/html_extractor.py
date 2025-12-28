"""
HTML content extractor using lxml.

Extracts text content from HTML files while:
- Removing all JavaScript and CSS
- Preserving structural elements (tables, headings, lists)
- Extracting metadata from <meta> tags
"""

import io
import logging
import re
from typing import Any, Dict, Generator, List

from lxml import etree
from lxml.html import HtmlElement, fromstring

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


class _HtmlTextExtractor:
    """Helper class to extract text from HTML while preserving structure."""

    def __init__(self, root: HtmlElement):
        self.root = root
        self.tables: List[List[List[str]]] = []
        self.headings: List[Dict[str, str]] = []
        self.links: List[Dict[str, str]] = []

    def _remove_unwanted_elements(self) -> None:
        """Remove script, style, and other unwanted elements."""
        for tag in REMOVE_TAGS:
            for element in self.root.xpath(f"//{tag}"):
                element.getparent().remove(element)

        # Also remove HTML comments
        for comment in self.root.xpath("//comment()"):
            parent = comment.getparent()
            if parent is not None:
                parent.remove(comment)

    def _extract_table(self, table_elem: HtmlElement) -> List[List[str]]:
        """Extract table as a list of rows, each row being a list of cell values."""
        rows = []
        for tr in table_elem.xpath(".//tr"):
            row = []
            for cell in tr.xpath(".//th | .//td"):
                # Get text content of the cell, stripping whitespace
                cell_text = self._get_element_text(cell).strip()
                # Normalize whitespace within cell
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
            for h in self.root.xpath(f"//h{level}"):
                text = self._get_element_text(h).strip()
                text = re.sub(r"\s+", " ", text)
                if text:
                    self.headings.append({"level": f"h{level}", "text": text})

    def _extract_links(self) -> None:
        """Extract all links with their text and href."""
        for a in self.root.xpath("//a[@href]"):
            href = a.get("href", "")
            text = self._get_element_text(a).strip()
            text = re.sub(r"\s+", " ", text)
            if href and text:
                self.links.append({"text": text, "href": href})

    def _get_element_text(self, element: HtmlElement) -> str:
        """Get text content of an element, including tail text of children."""
        return "".join(element.itertext())

    def _process_element(self, element: HtmlElement, depth: int = 0) -> str:
        """Recursively process an element and return its text representation."""
        if element.tag in REMOVE_TAGS:
            return ""

        tag = element.tag if isinstance(element.tag, str) else ""
        result_parts = []

        # Handle tables specially
        if tag == "table":
            table_data = self._extract_table(element)
            self.tables.append(table_data)
            table_text = self._format_table_as_text(table_data)
            return "\n" + table_text + "\n"

        # Handle list items
        if tag == "li":
            # Get the text content
            text = ""
            if element.text:
                text += element.text
            for child in element:
                text += self._process_element(child, depth + 1)
                if child.tail:
                    text += child.tail
            text = re.sub(r"\s+", " ", text.strip())
            return "  " * depth + "- " + text + "\n"

        # Handle headings
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            text = self._get_element_text(element).strip()
            text = re.sub(r"\s+", " ", text)
            return "\n" + text + "\n"

        # Handle line breaks
        if tag == "br":
            return "\n"

        # Handle horizontal rules
        if tag == "hr":
            return "\n---\n"

        # Add element's own text
        if element.text:
            result_parts.append(element.text)

        # Process children
        for child in element:
            child_text = self._process_element(child, depth)
            result_parts.append(child_text)
            # Add tail text (text after the child element)
            if child.tail:
                result_parts.append(child.tail)

        result = "".join(result_parts)

        # Add newlines around block elements
        if tag in BLOCK_TAGS:
            result = "\n" + result.strip() + "\n"

        return result

    def extract(self) -> str:
        """Extract and return the full text content."""
        self._remove_unwanted_elements()
        self._extract_headings()
        self._extract_links()

        # Find the body or use root
        body = self.root.xpath("//body")
        if body:
            text = self._process_element(body[0])
        else:
            text = self._process_element(self.root)

        # Clean up the text
        # Collapse multiple newlines to maximum of 2
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Remove leading/trailing whitespace from lines
        lines = [line.strip() for line in text.split("\n")]
        text = "\n".join(lines)
        # Remove leading/trailing newlines
        text = text.strip()

        return text


def _extract_metadata(root: HtmlElement, path: str | None) -> HtmlMetadata:
    """Extract metadata from HTML document."""
    metadata = HtmlMetadata()
    metadata.populate_from_path(path)

    # Extract title
    title_elem = root.xpath("//title")
    if title_elem:
        metadata.title = title_elem[0].text_content().strip()

    # Extract language from html tag
    html_elem = root.xpath("//html[@lang]")
    if html_elem:
        metadata.language = html_elem[0].get("lang", "")

    # Extract charset
    charset_meta = root.xpath("//meta[@charset]")
    if charset_meta:
        metadata.charset = charset_meta[0].get("charset", "")
    else:
        # Try content-type meta
        content_type_meta = root.xpath('//meta[@http-equiv="Content-Type"]/@content')
        if content_type_meta:
            match = re.search(r"charset=([^\s;]+)", content_type_meta[0])
            if match:
                metadata.charset = match.group(1)

    # Extract common meta tags
    meta_mappings = {
        "description": "description",
        "keywords": "keywords",
        "author": "author",
    }

    for attr_name, meta_name in meta_mappings.items():
        meta_elem = root.xpath(f'//meta[@name="{meta_name}"]/@content')
        if meta_elem:
            setattr(metadata, attr_name, meta_elem[0])

    return metadata


def read_html(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[HtmlContent, Any, None]:
    """
    Extract content from an HTML file.

    Args:
        file_like: A BytesIO object containing the HTML file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        HtmlContent dataclass with the extracted content.
    """
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
        root = fromstring(html_text)
    except etree.ParserError:
        # If parsing fails, try wrapping in html tags
        try:
            root = fromstring(f"<html><body>{html_text}</body></html>")
        except etree.ParserError:
            # Last resort: return empty content
            logger.warning("Failed to parse HTML content")
            metadata = HtmlMetadata()
            metadata.populate_from_path(path)
            yield HtmlContent(content="", metadata=metadata)
            return

    # Extract metadata
    metadata = _extract_metadata(root, path)

    # Extract text content
    extractor = _HtmlTextExtractor(root)
    text = extractor.extract()

    yield HtmlContent(
        content=text,
        tables=extractor.tables,
        headings=extractor.headings,
        links=extractor.links,
        metadata=metadata,
    )
