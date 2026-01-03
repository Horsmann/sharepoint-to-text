"""
EPUB eBook Extractor
====================

Extracts text content, metadata, and structure from EPUB eBook files.

File Format Background
----------------------
EPUB (Electronic Publication) is a widely-adopted open eBook format defined
by the IDPF (International Digital Publishing Forum). It is essentially a
ZIP archive containing:

    META-INF/container.xml: Points to the OPF (Open Packaging Format) file
    content.opf (or similar): Package document with metadata, manifest, spine
    *.xhtml/*.html: Content documents (chapters)
    *.css: Stylesheets
    images/: Embedded images
    toc.ncx or nav.xhtml: Navigation/table of contents

EPUB Structure
--------------
1. Container: META-INF/container.xml specifies the root file (OPF)
2. Package Document (OPF): Contains three main sections:
   - <metadata>: Dublin Core metadata (title, author, language, etc.)
   - <manifest>: Lists all files in the EPUB with IDs and media-types
   - <spine>: Defines reading order by referencing manifest item IDs
3. Content Documents: XHTML files containing the actual book content

EPUB Versions
-------------
- EPUB 2.0: Uses NCX for navigation, simpler metadata
- EPUB 3.0: Uses HTML5 nav element, richer metadata, multimedia support

This extractor supports both EPUB 2.x and 3.x formats.

Dependencies
------------
Python Standard Library only:
    - zipfile: ZIP archive handling
    - xml.etree.ElementTree: XML parsing
    - html.parser: HTML/XHTML content extraction
    - mimetypes: Content type detection

Extracted Content
-----------------
- metadata: Title, author, language, publisher, date, description, etc.
- chapters: Text content from each content document in reading order
- images: Embedded images with binary data
- toc: Table of contents structure

Known Limitations
-----------------
- DRM-protected EPUBs are not supported
- Embedded fonts are not extracted
- Audio/video content is not extracted
- JavaScript-enhanced EPUBs may not extract dynamic content
- Very complex CSS layouts may affect text ordering

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.epub_extractor import read_epub
    >>>
    >>> with open("book.epub", "rb") as f:
    ...     for epub in read_epub(io.BytesIO(f.read()), path="book.epub"):
    ...         print(f"Title: {epub.metadata.title}")
    ...         print(f"Author: {epub.metadata.creator}")
    ...         for chapter in epub.chapters:
    ...             print(f"Chapter: {chapter.title}")
    ...             print(chapter.text[:200])

See Also
--------
- EPUB 3.3 Specification: https://www.w3.org/TR/epub-33/
- html_extractor: For HTML parsing utilities

Maintenance Notes
-----------------
- Container.xml namespace varies between EPUB versions
- OPF namespace also varies (OPF 2.0 vs 3.0)
- NCX navigation (EPUB 2) vs nav.xhtml (EPUB 3)
"""

import io
import logging
import mimetypes
import re
from functools import lru_cache
from html.parser import HTMLParser
from typing import Any, Dict, Generator, List, Optional, Tuple
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    EpubChapter,
    EpubContent,
    EpubImage,
    EpubMetadata,
)
from sharepoint2text.extractors.util.zip_context import ZipContext

logger = logging.getLogger(__name__)

# Namespaces used in EPUB files
NS = {
    "container": "urn:oasis:names:tc:opendocument:xmlns:container",
    "opf": "http://www.idpf.org/2007/opf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "ncx": "http://www.daisy.org/z3986/2005/ncx/",
    "xhtml": "http://www.w3.org/1999/xhtml",
    "epub": "http://www.idpf.org/2007/ops",
}

# Tags to remove entirely when extracting text from XHTML
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


@lru_cache(maxsize=256)
def _guess_content_type(path: str) -> str:
    """Guess content type from file path."""
    return mimetypes.guess_type(path)[0] or "application/octet-stream"


class _XhtmlTextExtractor(HTMLParser):
    """
    Extract text content from XHTML/HTML documents.

    This is a simplified version of the HTML extractor optimized for EPUB
    content documents. It extracts text while preserving basic structure
    (paragraphs, headings, lists) and also extracts tables.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.text_parts: List[str] = []
        self.skip_depth = 0
        self.in_block = False
        self.tables: List[List[List[str]]] = []
        self._current_table: List[List[str]] = []
        self._current_row: List[str] = []
        self._current_cell: List[str] = []
        self._in_table = False
        self._in_cell = False
        self._title: str = ""
        self._in_title = False

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]):
        tag = tag.lower()

        if self.skip_depth > 0:
            self.skip_depth += 1
            return

        if tag in REMOVE_TAGS:
            self.skip_depth = 1
            return

        if tag == "title":
            self._in_title = True
            return

        if tag == "table":
            self._in_table = True
            self._current_table = []
            return

        if self._in_table:
            if tag == "tr":
                self._current_row = []
            elif tag in ("td", "th"):
                self._in_cell = True
                self._current_cell = []

        if tag in BLOCK_TAGS:
            self.text_parts.append("\n")
            self.in_block = True

        if tag == "br":
            self.text_parts.append("\n")

    def handle_endtag(self, tag: str):
        tag = tag.lower()

        if self.skip_depth > 0:
            self.skip_depth -= 1
            return

        if tag == "title":
            self._in_title = False
            return

        if tag == "table":
            if self._current_table:
                self.tables.append(self._current_table)
            self._current_table = []
            self._in_table = False
            return

        if self._in_table:
            if tag == "tr":
                if self._current_row:
                    self._current_table.append(self._current_row)
                self._current_row = []
            elif tag in ("td", "th"):
                cell_text = " ".join(self._current_cell).strip()
                cell_text = re.sub(r"\s+", " ", cell_text)
                self._current_row.append(cell_text)
                self._current_cell = []
                self._in_cell = False

        if tag in BLOCK_TAGS:
            self.text_parts.append("\n")
            self.in_block = False

    def handle_data(self, data: str):
        if self.skip_depth > 0:
            return

        if self._in_title:
            self._title += data
            return

        if self._in_cell:
            self._current_cell.append(data)
            return

        self.text_parts.append(data)

    def get_text(self) -> str:
        """Get the extracted text, cleaned up."""
        text = "".join(self.text_parts)
        # Collapse multiple newlines to maximum of 2
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Collapse multiple spaces
        text = re.sub(r"[ \t]+", " ", text)
        # Remove leading/trailing whitespace from lines
        lines = [line.strip() for line in text.split("\n")]
        text = "\n".join(lines)
        return text.strip()

    def get_title(self) -> str:
        """Get the document title from <title> tag."""
        return self._title.strip()

    def get_tables(self) -> List[List[List[str]]]:
        """Get extracted tables."""
        return self.tables


class _EpubContext(ZipContext):
    """Context for EPUB extraction with cached OPF parsing."""

    def __init__(self, file_like: io.BytesIO):
        super().__init__(file_like)
        self._opf_path: str = ""
        self._opf_dir: str = ""
        self._opf_root: Optional[ET.Element] = None
        self._manifest: Dict[str, Dict[str, str]] = {}  # id -> {href, media-type}
        self._spine: List[str] = []  # List of manifest item IDs in reading order
        self._metadata: EpubMetadata = EpubMetadata()

        self._parse_container()
        if self._opf_path:
            self._parse_opf()

    def _parse_container(self) -> None:
        """Parse META-INF/container.xml to find the OPF file."""
        container_path = "META-INF/container.xml"
        if not self.exists(container_path):
            logger.warning("No container.xml found in EPUB")
            return

        try:
            root = self.read_xml_root(container_path)

            # Find rootfile element - try with and without namespace
            rootfile = root.find(".//container:rootfile", NS)
            if rootfile is None:
                # Try without namespace
                rootfile = root.find(".//{*}rootfile")
            if rootfile is None:
                # Try bare
                rootfile = root.find(".//rootfile")

            if rootfile is not None:
                self._opf_path = rootfile.get("full-path", "")
                # Get the directory containing the OPF for resolving relative paths
                if "/" in self._opf_path:
                    self._opf_dir = self._opf_path.rsplit("/", 1)[0] + "/"
                else:
                    self._opf_dir = ""
        except Exception as e:
            logger.warning("Failed to parse container.xml: %s", e)

    def _parse_opf(self) -> None:
        """Parse the OPF (Open Packaging Format) file."""
        if not self._opf_path or not self.exists(self._opf_path):
            return

        try:
            self._opf_root = self.read_xml_root(self._opf_path)
            self._parse_metadata()
            self._parse_manifest()
            self._parse_spine()
        except Exception as e:
            logger.warning("Failed to parse OPF file: %s", e)

    def _parse_metadata(self) -> None:
        """Extract Dublin Core and EPUB metadata from OPF."""
        if self._opf_root is None:
            return

        # Find metadata element
        metadata_elem = self._opf_root.find("opf:metadata", NS)
        if metadata_elem is None:
            metadata_elem = self._opf_root.find("{*}metadata")
        if metadata_elem is None:
            return

        # Helper to get text from first matching element
        def get_dc(name: str) -> str:
            elem = metadata_elem.find(f"dc:{name}", NS)
            if elem is None:
                elem = metadata_elem.find(f"{{http://purl.org/dc/elements/1.1/}}{name}")
            return (elem.text or "").strip() if elem is not None else ""

        self._metadata.title = get_dc("title")
        self._metadata.creator = get_dc("creator")
        self._metadata.language = get_dc("language")
        self._metadata.identifier = get_dc("identifier")
        self._metadata.publisher = get_dc("publisher")
        self._metadata.date = get_dc("date")
        self._metadata.description = get_dc("description")
        self._metadata.subject = get_dc("subject")
        self._metadata.rights = get_dc("rights")
        self._metadata.contributor = get_dc("contributor")

        # Get EPUB version from package element
        package = self._opf_root
        if package.tag.endswith("package") or package.tag == "package":
            self._metadata.epub_version = package.get("version", "")

    def _parse_manifest(self) -> None:
        """Parse the manifest section of OPF."""
        if self._opf_root is None:
            return

        manifest_elem = self._opf_root.find("opf:manifest", NS)
        if manifest_elem is None:
            manifest_elem = self._opf_root.find("{*}manifest")
        if manifest_elem is None:
            return

        for item in manifest_elem.findall("opf:item", NS):
            item_id = item.get("id", "")
            href = item.get("href", "")
            media_type = item.get("media-type", "")
            if item_id and href:
                self._manifest[item_id] = {
                    "href": href,
                    "media-type": media_type,
                }

        # Also try without namespace
        if not self._manifest:
            for item in manifest_elem.findall("{*}item"):
                item_id = item.get("id", "")
                href = item.get("href", "")
                media_type = item.get("media-type", "")
                if item_id and href:
                    self._manifest[item_id] = {
                        "href": href,
                        "media-type": media_type,
                    }

    def _parse_spine(self) -> None:
        """Parse the spine section of OPF to get reading order."""
        if self._opf_root is None:
            return

        spine_elem = self._opf_root.find("opf:spine", NS)
        if spine_elem is None:
            spine_elem = self._opf_root.find("{*}spine")
        if spine_elem is None:
            return

        for itemref in spine_elem.findall("opf:itemref", NS):
            idref = itemref.get("idref", "")
            if idref:
                self._spine.append(idref)

        # Also try without namespace
        if not self._spine:
            for itemref in spine_elem.findall("{*}itemref"):
                idref = itemref.get("idref", "")
                if idref:
                    self._spine.append(idref)

    def resolve_href(self, href: str) -> str:
        """Resolve a relative href to its full path in the ZIP."""
        if href.startswith("/"):
            return href[1:]
        return self._opf_dir + href

    @property
    def opf_dir(self) -> str:
        return self._opf_dir

    @property
    def manifest(self) -> Dict[str, Dict[str, str]]:
        return self._manifest

    @property
    def spine(self) -> List[str]:
        return self._spine

    @property
    def metadata(self) -> EpubMetadata:
        return self._metadata


def _extract_chapter(
    ctx: _EpubContext,
    item_id: str,
    chapter_number: int,
    image_counter: int,
) -> Tuple[Optional[EpubChapter], int, List[EpubImage]]:
    """
    Extract content from a single chapter (content document).

    Args:
        ctx: The EPUB context
        item_id: Manifest item ID for this chapter
        chapter_number: 1-based chapter number
        image_counter: Current global image counter

    Returns:
        Tuple of (EpubChapter, updated_image_counter, list of images)
    """
    if item_id not in ctx.manifest:
        return None, image_counter, []

    item = ctx.manifest[item_id]
    href = ctx.resolve_href(item["href"])
    media_type = item.get("media-type", "")

    # Only process XHTML/HTML content
    if not (
        media_type.startswith("application/xhtml")
        or media_type.startswith("text/html")
        or href.endswith((".xhtml", ".html", ".htm"))
    ):
        return None, image_counter, []

    if not ctx.exists(href):
        logger.debug("Content document not found: %s", href)
        return None, image_counter, []

    try:
        content = ctx.read_text(href)
    except Exception as e:
        logger.debug("Failed to read content document %s: %s", href, e)
        return None, image_counter, []

    # Parse the XHTML content
    parser = _XhtmlTextExtractor()
    try:
        parser.feed(content)
    except Exception as e:
        logger.debug("Failed to parse content document %s: %s", href, e)
        return None, image_counter, []

    text = parser.get_text()
    title = parser.get_title()
    tables = parser.get_tables()

    # If no title from <title> tag, try to extract from first heading
    if not title:
        # Look for first h1 or h2 in text
        heading_match = re.search(r"^([^\n]+)", text)
        if heading_match:
            first_line = heading_match.group(1).strip()
            if len(first_line) < 100:  # Reasonable title length
                title = first_line

    # Extract image references from this chapter
    # (We'll collect the actual image data separately from manifest)
    chapter_images: List[EpubImage] = []

    chapter = EpubChapter(
        chapter_number=chapter_number,
        href=href,
        title=title,
        text=text,
        images=chapter_images,
        tables=tables,
    )

    return chapter, image_counter, chapter_images


def _extract_images(ctx: _EpubContext) -> List[EpubImage]:
    """Extract all images from the EPUB manifest."""
    images: List[EpubImage] = []
    image_counter = 0

    for item_id, item in ctx.manifest.items():
        media_type = item.get("media-type", "")
        if not media_type.startswith("image/"):
            continue

        href = ctx.resolve_href(item["href"])
        if not ctx.exists(href):
            continue

        try:
            image_counter += 1
            data = ctx.read_bytes(href)
            images.append(
                EpubImage(
                    image_index=image_counter,
                    href=href,
                    content_type=media_type,
                    data=io.BytesIO(data),
                    size_bytes=len(data),
                )
            )
        except Exception as e:
            logger.debug("Failed to extract image %s: %s", href, e)

    return images


def _extract_toc(ctx: _EpubContext) -> List[Dict[str, str]]:
    """
    Extract table of contents from EPUB.

    Tries EPUB 3 nav.xhtml first, falls back to EPUB 2 NCX.
    """
    toc: List[Dict[str, str]] = []

    # Look for nav document in manifest (EPUB 3)
    for item_id, item in ctx.manifest.items():
        if item.get("media-type") == "application/xhtml+xml":
            href = ctx.resolve_href(item["href"])
            # Check if this is a nav document by looking at properties
            # or by checking if it contains nav element
            # For simplicity, we'll check common nav document names
            if "nav" in href.lower() or "toc" in href.lower():
                toc = _parse_nav_document(ctx, href)
                if toc:
                    return toc

    # Fall back to NCX (EPUB 2)
    for item_id, item in ctx.manifest.items():
        if item.get("media-type") == "application/x-dtbncx+xml":
            href = ctx.resolve_href(item["href"])
            toc = _parse_ncx(ctx, href)
            if toc:
                return toc

    return toc


def _parse_nav_document(ctx: _EpubContext, href: str) -> List[Dict[str, str]]:
    """Parse EPUB 3 nav document."""
    if not ctx.exists(href):
        return []

    try:
        content = ctx.read_text(href)
        # Simple extraction of nav items using regex
        # Look for <a href="...">title</a> patterns
        toc: List[Dict[str, str]] = []
        pattern = re.compile(r'<a[^>]+href="([^"]+)"[^>]*>([^<]+)</a>', re.IGNORECASE)
        for match in pattern.finditer(content):
            link_href = match.group(1)
            title = match.group(2).strip()
            if title:
                toc.append({"title": title, "href": link_href})
        return toc
    except Exception as e:
        logger.debug("Failed to parse nav document: %s", e)
        return []


def _parse_ncx(ctx: _EpubContext, href: str) -> List[Dict[str, str]]:
    """Parse EPUB 2 NCX navigation document."""
    if not ctx.exists(href):
        return []

    try:
        root = ctx.read_xml_root(href)
        toc: List[Dict[str, str]] = []

        # Find navMap and navPoints
        nav_map = root.find("ncx:navMap", NS)
        if nav_map is None:
            nav_map = root.find("{*}navMap")
        if nav_map is None:
            return []

        def extract_nav_points(parent: ET.Element) -> None:
            for nav_point in parent.findall("ncx:navPoint", NS):
                label_elem = nav_point.find("ncx:navLabel/ncx:text", NS)
                content_elem = nav_point.find("ncx:content", NS)

                if label_elem is None:
                    label_elem = nav_point.find("{*}navLabel/{*}text")
                if content_elem is None:
                    content_elem = nav_point.find("{*}content")

                title = (
                    (label_elem.text or "").strip() if label_elem is not None else ""
                )
                link_href = (
                    content_elem.get("src", "") if content_elem is not None else ""
                )

                if title:
                    toc.append({"title": title, "href": link_href})

                # Recurse for nested nav points
                extract_nav_points(nav_point)

            # Also try without namespace
            for nav_point in parent.findall("{*}navPoint"):
                if nav_point in parent.findall("ncx:navPoint", NS):
                    continue  # Already processed

                label_elem = nav_point.find("{*}navLabel/{*}text")
                content_elem = nav_point.find("{*}content")

                title = (
                    (label_elem.text or "").strip() if label_elem is not None else ""
                )
                link_href = (
                    content_elem.get("src", "") if content_elem is not None else ""
                )

                if title:
                    toc.append({"title": title, "href": link_href})

                extract_nav_points(nav_point)

        extract_nav_points(nav_map)
        return toc
    except Exception as e:
        logger.debug("Failed to parse NCX: %s", e)
        return []


def _is_epub_encrypted(ctx: _EpubContext) -> bool:
    """Check if the EPUB has DRM encryption."""
    # Check for encryption.xml which indicates DRM
    if ctx.exists("META-INF/encryption.xml"):
        try:
            root = ctx.read_xml_root("META-INF/encryption.xml")
            # If there are any EncryptedData elements, the EPUB is encrypted
            encrypted = root.findall(
                ".//{http://www.w3.org/2001/04/xmlenc#}EncryptedData"
            )
            if encrypted:
                return True
        except Exception:
            pass

    # Check for rights.xml (Adobe DRM)
    if ctx.exists("META-INF/rights.xml"):
        return True

    return False


def read_epub(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EpubContent, Any, None]:
    """
    Extract all relevant content from an EPUB eBook file.

    Primary entry point for EPUB file extraction. Opens the ZIP archive,
    parses the OPF package document, and extracts content from each chapter
    in reading order as defined by the spine.

    This function uses a generator pattern for API consistency with other
    extractors, even though EPUB files contain exactly one book.

    Args:
        file_like: BytesIO object containing the complete EPUB file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned EpubContent.metadata.

    Yields:
        EpubContent: Single EpubContent object containing:
            - metadata: EpubMetadata with title, author, language, etc.
            - chapters: List of EpubChapter objects in reading order
            - images: List of EpubImage objects with binary data
            - toc: Table of contents as list of {title, href} dicts

    Raises:
        ExtractionFileEncryptedError: If the EPUB has DRM protection.
        ExtractionFailedError: If extraction fails for other reasons.

    Example:
        >>> import io
        >>> with open("book.epub", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for epub in read_epub(data, path="book.epub"):
        ...         print(f"Title: {epub.metadata.title}")
        ...         print(f"Chapters: {len(epub.chapters)}")
        ...         for chapter in epub.chapters:
        ...             print(f"  {chapter.chapter_number}: {chapter.title}")
    """
    try:
        file_like.seek(0)

        ctx = _EpubContext(file_like)
        try:
            # Check for DRM encryption
            if _is_epub_encrypted(ctx):
                raise ExtractionFileEncryptedError(
                    "EPUB is DRM-protected and cannot be extracted"
                )

            metadata = ctx.metadata
            metadata.populate_from_path(path)

            # Extract chapters in spine order
            chapters: List[EpubChapter] = []
            image_counter = 0
            chapter_number = 0

            for item_id in ctx.spine:
                chapter_number += 1
                chapter, image_counter, _ = _extract_chapter(
                    ctx, item_id, chapter_number, image_counter
                )
                if chapter is not None:
                    chapters.append(chapter)

            # Extract images from manifest
            images = _extract_images(ctx)

            # Extract table of contents
            toc = _extract_toc(ctx)

        finally:
            ctx.close()

        yield EpubContent(
            metadata=metadata,
            chapters=chapters,
            images=images,
            toc=toc,
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract EPUB file", cause=exc) from exc
