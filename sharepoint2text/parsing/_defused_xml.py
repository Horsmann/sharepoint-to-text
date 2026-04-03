"""
Safe XML parsing module.

Re-exports defusedxml's safe parsing functions (``fromstring``, ``parse``,
``iterparse``) together with stdlib's ``Element`` and ``ParseError`` so that
extractor modules can ``import _defused_xml as ET`` and use ``ET.Element``,
``ET.fromstring``, ``ET.ParseError``, etc. without any XXE risk.
"""

from xml.etree.ElementTree import Element, ParseError  # noqa: F401

from defusedxml.ElementTree import (  # noqa: F401  # type: ignore[import-untyped]
    fromstring,
    iterparse,
    parse,
    tostring,
)
