from __future__ import annotations

import mimetypes
from functools import lru_cache
from xml.etree import ElementTree as ET

from sharepoint2text.parsing.extractors.data_types import OpenDocumentMetadata


@lru_cache(maxsize=512)
def guess_content_type(path: str) -> str:
    return mimetypes.guess_type(path)[0] or "application/octet-stream"


def extract_odf_metadata(
    meta_root: ET.Element | None, ns: dict[str, str]
) -> OpenDocumentMetadata:
    metadata = OpenDocumentMetadata()
    if meta_root is None:
        return metadata

    meta_elem = meta_root.find(".//office:meta", ns)
    if meta_elem is None:
        return metadata

    title = meta_elem.find("dc:title", ns)
    if title is not None and title.text:
        metadata.title = title.text

    description = meta_elem.find("dc:description", ns)
    if description is not None and description.text:
        metadata.description = description.text

    subject = meta_elem.find("dc:subject", ns)
    if subject is not None and subject.text:
        metadata.subject = subject.text

    creator = meta_elem.find("dc:creator", ns)
    if creator is not None and creator.text:
        metadata.creator = creator.text

    date = meta_elem.find("dc:date", ns)
    if date is not None and date.text:
        metadata.date = date.text

    language = meta_elem.find("dc:language", ns)
    if language is not None and language.text:
        metadata.language = language.text

    keywords = meta_elem.find("meta:keyword", ns)
    if keywords is not None and keywords.text:
        metadata.keywords = keywords.text

    initial_creator = meta_elem.find("meta:initial-creator", ns)
    if initial_creator is not None and initial_creator.text:
        metadata.initial_creator = initial_creator.text

    creation_date = meta_elem.find("meta:creation-date", ns)
    if creation_date is not None and creation_date.text:
        metadata.creation_date = creation_date.text

    editing_cycles = meta_elem.find("meta:editing-cycles", ns)
    if editing_cycles is not None and editing_cycles.text:
        try:
            metadata.editing_cycles = int(editing_cycles.text)
        except ValueError:
            pass

    editing_duration = meta_elem.find("meta:editing-duration", ns)
    if editing_duration is not None and editing_duration.text:
        metadata.editing_duration = editing_duration.text

    generator = meta_elem.find("meta:generator", ns)
    if generator is not None and generator.text:
        metadata.generator = generator.text

    return metadata


def element_text(
    element: ET.Element,
    *,
    text_space_tag: str,
    text_tab_tag: str,
    text_line_break_tag: str,
    attr_text_c: str,
    skip_tags: set[str] | None = None,
) -> str:
    parts: list[str] = []
    _append_element_text(
        element,
        parts,
        text_space_tag=text_space_tag,
        text_tab_tag=text_tab_tag,
        text_line_break_tag=text_line_break_tag,
        attr_text_c=attr_text_c,
        skip_tags=skip_tags or set(),
    )
    return "".join(parts)


def _append_element_text(
    element: ET.Element,
    parts: list[str],
    *,
    text_space_tag: str,
    text_tab_tag: str,
    text_line_break_tag: str,
    attr_text_c: str,
    skip_tags: set[str],
) -> None:
    text = element.text
    if text:
        parts.append(text)

    for child in element:
        tag = child.tag
        if tag in skip_tags:
            pass
        elif tag == text_space_tag:
            raw_count = child.get(attr_text_c, "1")
            try:
                count = int(raw_count)
            except ValueError:
                count = 1
            if count > 0:
                parts.append(" " * count)
        elif tag == text_tab_tag:
            parts.append("\t")
        elif tag == text_line_break_tag:
            parts.append("\n")
        else:
            _append_element_text(
                child,
                parts,
                text_space_tag=text_space_tag,
                text_tab_tag=text_tab_tag,
                text_line_break_tag=text_line_break_tag,
                attr_text_c=attr_text_c,
                skip_tags=skip_tags,
            )

        tail = child.tail
        if tail:
            parts.append(tail)
