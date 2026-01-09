"""
ODF Formula Extractor
=====================

Extracts text content and metadata from OpenDocument Formula (.odf) files
created by LibreOffice Math / OpenOffice Math.

ODF formula documents are ZIP archives containing XML files. The formula
payload is typically stored under:

    content.xml -> office:body/office:formula

Formula markup often uses MathML elements. This extractor prioritizes:
    - math:annotation (e.g., StarMath source)
    - text:p / text:h elements (captions or surrounding text)
and falls back to a best-effort itertext() representation when needed.
"""

import io
import logging
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    OdfContent,
    OpenDocumentMetadata,
)
from sharepoint2text.parsing.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.parsing.extractors.util.zip_context import ZipContext

logger = logging.getLogger(__name__)


NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "math": "http://www.w3.org/1998/Math/MathML",
}

_TEXT_SPACE_TAG = f"{{{NS['text']}}}s"
_TEXT_TAB_TAG = f"{{{NS['text']}}}tab"
_TEXT_LINE_BREAK_TAG = f"{{{NS['text']}}}line-break"
_OFFICE_ANNOTATION_TAG = f"{{{NS['office']}}}annotation"

_TEXT_P_TAG = f"{{{NS['text']}}}p"
_TEXT_H_TAG = f"{{{NS['text']}}}h"
_MATH_ANNOTATION_TAG = f"{{{NS['math']}}}annotation"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"


def _get_text_recursive(element: ET.Element) -> str:
    """Recursively extract all text from an element and its children."""
    parts: list[str] = []

    text = element.text
    if text:
        parts.append(text)

    for child in element:
        tag = child.tag
        if tag == _TEXT_SPACE_TAG:
            count = int(child.get(_ATTR_TEXT_C, "1"))
            parts.append(" " * count)
        elif tag == _TEXT_TAB_TAG:
            parts.append("\t")
        elif tag == _TEXT_LINE_BREAK_TAG:
            parts.append("\n")
        elif tag == _OFFICE_ANNOTATION_TAG:
            # Skip annotations in main text extraction.
            pass
        else:
            parts.append(_get_text_recursive(child))

        tail = child.tail
        if tail:
            parts.append(tail)

    return "".join(parts)


def _extract_metadata(meta_root: ET.Element | None) -> OpenDocumentMetadata:
    """Extract metadata from meta.xml."""
    metadata = OpenDocumentMetadata()

    if meta_root is None:
        return metadata

    meta_elem = meta_root.find(".//office:meta", NS)
    if meta_elem is None:
        return metadata

    title = meta_elem.find("dc:title", NS)
    if title is not None and title.text:
        metadata.title = title.text

    description = meta_elem.find("dc:description", NS)
    if description is not None and description.text:
        metadata.description = description.text

    subject = meta_elem.find("dc:subject", NS)
    if subject is not None and subject.text:
        metadata.subject = subject.text

    creator = meta_elem.find("dc:creator", NS)
    if creator is not None and creator.text:
        metadata.creator = creator.text

    date = meta_elem.find("dc:date", NS)
    if date is not None and date.text:
        metadata.date = date.text

    language = meta_elem.find("dc:language", NS)
    if language is not None and language.text:
        metadata.language = language.text

    keywords = meta_elem.find("meta:keyword", NS)
    if keywords is not None and keywords.text:
        metadata.keywords = keywords.text

    initial_creator = meta_elem.find("meta:initial-creator", NS)
    if initial_creator is not None and initial_creator.text:
        metadata.initial_creator = initial_creator.text

    creation_date = meta_elem.find("meta:creation-date", NS)
    if creation_date is not None and creation_date.text:
        metadata.creation_date = creation_date.text

    editing_cycles = meta_elem.find("meta:editing-cycles", NS)
    if editing_cycles is not None and editing_cycles.text:
        try:
            metadata.editing_cycles = int(editing_cycles.text)
        except ValueError:
            pass

    editing_duration = meta_elem.find("meta:editing-duration", NS)
    if editing_duration is not None and editing_duration.text:
        metadata.editing_duration = editing_duration.text

    generator = meta_elem.find("meta:generator", NS)
    if generator is not None and generator.text:
        metadata.generator = generator.text

    return metadata


def _extract_full_text(formula_root: ET.Element) -> str:
    lines: list[str] = []

    for elem in formula_root.iter():
        if elem.tag in (_TEXT_H_TAG, _TEXT_P_TAG):
            value = _get_text_recursive(elem).strip()
            if value:
                lines.append(value)
        elif elem.tag == _MATH_ANNOTATION_TAG:
            value = (elem.text or "").strip()
            if value:
                lines.append(value)

    if lines:
        return "\n".join(lines).strip()

    # Best-effort fallback for purely MathML documents.
    text = " ".join(token.strip() for token in formula_root.itertext() if token.strip())
    return text.strip()


def read_odf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdfContent, Any, None]:
    """Extract text and metadata from an ODF formula file."""
    try:
        file_like.seek(0)
        if is_odf_encrypted(file_like):
            raise ExtractionFileEncryptedError("ODF is encrypted or password-protected")

        ctx = ZipContext(file_like)
        try:
            meta_root = (
                ctx.read_xml_root("meta.xml") if ctx.exists("meta.xml") else None
            )
            content_root = (
                ctx.read_xml_root("content.xml") if ctx.exists("content.xml") else None
            )
            if content_root is None:
                raise ExtractionFailedError("Invalid ODF file: content.xml not found")

            formula = content_root.find(".//office:body/office:formula", NS)
            if formula is None:
                raise ExtractionFailedError("Invalid ODF file: formula body not found")

            metadata = _extract_metadata(meta_root)
            full_text = _extract_full_text(formula)
        finally:
            ctx.close()

        metadata.populate_from_path(path)
        yield OdfContent(metadata=metadata, full_text=full_text)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract ODF file", cause=exc) from exc
