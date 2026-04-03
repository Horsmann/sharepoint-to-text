"""
Shared OOXML text and element processing utilities.

Consolidates common text extraction and XML element processing logic used
across DOCX, PPTX, and other OOXML extractors, reducing code duplication.
"""

from typing import Callable

from sharepoint2text.parsing import _defused_xml as ET
from sharepoint2text.parsing.extractors.ms_modern.ooxml_namespaces import (
    M_OMATH,
    M_OMATHPARA,
    MC_NS,
    W_NS,
    W_T,
)

# =============================================================================
# Text extraction from elements
# =============================================================================


def collect_text_from_element(element: ET.Element) -> str:
    """
    Extract all text from text elements within an element.

    Recursively finds all text content, typically used for collecting text
    from runs (w:r in DOCX) or similar text container elements.

    Args:
        element: The element to extract text from.

    Returns:
        Concatenated text content from all text elements.
    """
    return "".join(t.text for t in element.iter(W_T) if t.text)


def get_element_immediate_text(element: ET.Element) -> str:
    """
    Extract text from direct child text elements only (not recursive).

    Used when you want text from immediate children but not deeply nested text.

    Args:
        element: The element to extract text from.

    Returns:
        Concatenated text from immediate text children.
    """
    text_parts: list[str] = []
    for child in element:
        if child.tag == W_T and child.text:
            text_parts.append(child.text)
    return "".join(text_parts)


# =============================================================================
# Element tree traversal and processing
# =============================================================================


def process_text_element(
    elem: ET.Element,
    parts: list[str],
    include_formulas: bool,
    *,
    omath_tag: str = M_OMATH,
    omath_para_tag: str = M_OMATHPARA,
    formula_formatter: Callable[[ET.Element], str] | None = None,
) -> None:
    """
    Recursively process a text element and append extracted text to parts list.

    Handles special cases like:
    - AlternateContent (markup compatibility)
    - Fallback elements
    - Text runs (w:r in DOCX)
    - Math formulas (OMML)

    Args:
        elem: The element to process.
        parts: List to append extracted text to.
        include_formulas: If True, format formulas and append them.
        omath_tag: Tag for inline math elements (e.g., m:oMath).
        omath_para_tag: Tag for display math elements (e.g., m:oMathPara).
        formula_formatter: Callable to convert formula element to string
                          (e.g., omml_to_latex). Required if include_formulas=True.
    """
    tag = elem.tag

    # Handle markup compatibility AlternateContent
    if tag.endswith("}AlternateContent"):
        mc_choice_tag = f"{{{MC_NS}}}Choice"
        choice = elem.find(mc_choice_tag)
        if choice is not None:
            for child in choice:
                process_text_element(
                    child,
                    parts,
                    include_formulas,
                    omath_tag=omath_tag,
                    omath_para_tag=omath_para_tag,
                    formula_formatter=formula_formatter,
                )
        return

    # Skip Fallback elements
    if tag.endswith("}Fallback"):
        return

    # Handle text runs (w:r in DOCX)
    w_r_tag = f"{{{W_NS}}}r"
    if tag == w_r_tag:
        for child in elem:
            if child.tag == W_T:
                if child.text:
                    parts.append(child.text)
            elif child.tag.endswith("}AlternateContent"):
                process_text_element(
                    child,
                    parts,
                    include_formulas,
                    omath_tag=omath_tag,
                    omath_para_tag=omath_para_tag,
                    formula_formatter=formula_formatter,
                )
        return

    # Handle inline math (oMath)
    if tag == omath_tag:
        if include_formulas and formula_formatter:
            latex = formula_formatter(elem)
            if latex.strip():
                parts.append(f"${latex}$")
        return

    # Handle display math (oMathPara)
    if tag == omath_para_tag:
        if include_formulas and formula_formatter:
            omath = elem.find(omath_tag)
            if omath is not None:
                latex = formula_formatter(omath)
                if latex.strip():
                    parts.append(f"$${latex}$$")
        return

    # Recurse into children
    for child in elem:
        process_text_element(
            child,
            parts,
            include_formulas,
            omath_tag=omath_tag,
            omath_para_tag=omath_para_tag,
            formula_formatter=formula_formatter,
        )


def extract_text_with_formulas(
    element: ET.Element,
    include_formulas: bool = True,
    *,
    omath_tag: str = M_OMATH,
    omath_para_tag: str = M_OMATHPARA,
    formula_formatter: Callable[[ET.Element], str] | None = None,
) -> str:
    """
    Extract text from an element, optionally including formatted formulas.

    Convenience wrapper around process_text_element that handles text collection.

    Args:
        element: The element to extract text from.
        include_formulas: If True, include formula elements formatted via formula_formatter.
        omath_tag: Tag for inline math elements.
        omath_para_tag: Tag for display math elements.
        formula_formatter: Callable to format formula elements (e.g., omml_to_latex).

    Returns:
        Extracted text with optional formula markers.
    """
    parts: list[str] = []
    process_text_element(
        element,
        parts,
        include_formulas,
        omath_tag=omath_tag,
        omath_para_tag=omath_para_tag,
        formula_formatter=formula_formatter,
    )
    return "".join(parts)


# =============================================================================
# Element filtering and searching
# =============================================================================


def find_elements_by_tag(
    element: ET.Element, tag: str, recursive: bool = True
) -> list[ET.Element]:
    """
    Find all elements matching a tag, either recursively or direct children.

    Args:
        element: The parent element to search.
        tag: The tag to match (should be fully qualified with namespace).
        recursive: If True, search recursively (iter). If False, only direct children.

    Returns:
        List of matching elements.
    """
    if recursive:
        return list(element.iter(tag))
    return element.findall(tag)


def element_has_child(element: ET.Element, tag: str) -> bool:
    """
    Check if an element has a direct child with the given tag.

    Args:
        element: The element to check.
        tag: The tag to search for.

    Returns:
        True if a matching child exists, False otherwise.
    """
    return element.find(tag) is not None


# =============================================================================
# Attribute extraction helpers
# =============================================================================


def get_attribute(element: ET.Element, attr_tag: str, default: str = "") -> str:
    """
    Get an attribute value from an element, with a default fallback.

    Args:
        element: The element to extract from.
        attr_tag: The attribute to extract (can be a tag like W_VAL).
        default: Default value if attribute is not found.

    Returns:
        The attribute value or default.
    """
    if element is None:
        return default
    return element.get(attr_tag, default)


def get_child_attribute(
    element: ET.Element, child_tag: str, attr: str, default: str = ""
) -> str:
    """
    Get an attribute from a child element.

    Args:
        element: The parent element.
        child_tag: The child element tag.
        attr: The attribute name.
        default: Default if child or attribute not found.

    Returns:
        The attribute value or default.
    """
    child = element.find(child_tag) if element is not None else None
    if child is None:
        return default
    return child.get(attr, default)


def get_first_attribute(element: ET.Element | None, *attrs: str) -> str | None:
    """
    Get the first non-empty attribute from a list of attribute names.

    Useful for attributes with fallbacks (e.g., W_ASCII, W_HANSI, W_CS for fonts).

    Args:
        element: The element to extract from.
        *attrs: Attribute names to try in order.

    Returns:
        First non-empty attribute value, or None if all are empty.
    """
    if element is None:
        return None
    for attr in attrs:
        if value := element.get(attr):
            return value
    return None


# =============================================================================
# Boolean and value parsing
# =============================================================================


def parse_boolean_element(
    element: ET.Element | None, val_attr: str | None = None
) -> bool | None:
    """
    Parse a boolean element (like w:b for bold).

    In OOXML, boolean elements can be:
    - Present without attribute: True
    - Attribute val="true" or val="1": True
    - Attribute val="false" or val="0": False
    - Absent: None (unspecified)

    Args:
        element: The element to parse.
        val_attr: Optional value attribute (e.g., "{...}val"). If None, checks presence.

    Returns:
        True, False, or None if unspecified.
    """
    if element is None:
        return None
    if val_attr is None:
        return True  # Element presence = True
    val = element.get(val_attr)
    if val is None:
        return True  # No attribute specified = True
    return val.lower() not in ("false", "0")


def parse_int_attribute(element: ET.Element | None, attr: str) -> int | None:
    """
    Parse an integer attribute, returning None on failure.

    Args:
        element: The element to extract from.
        attr: The attribute name.

    Returns:
        Parsed integer or None if attribute missing or invalid.
    """
    if element is None:
        return None
    val = element.get(attr)
    if not val:
        return None
    try:
        return int(val)
    except ValueError:
        return None


def parse_float_attribute(element: ET.Element | None, attr: str) -> float | None:
    """
    Parse a float attribute, returning None on failure.

    Args:
        element: The element to extract from.
        attr: The attribute name.

    Returns:
        Parsed float or None if attribute missing or invalid.
    """
    if element is None:
        return None
    val = element.get(attr)
    if not val:
        return None
    try:
        return float(val)
    except ValueError:
        return None


# =============================================================================
# Unit conversion helpers
# =============================================================================


def twips_to_inches(twips_value: str | int | None) -> float | None:
    """
    Convert OOXML twips (1/20th of a point) to inches.

    Args:
        twips_value: Value in twips, as string or int.

    Returns:
        Value in inches, or None if conversion fails.
    """
    if twips_value is None or twips_value == "":
        return None
    try:
        twips_int = int(twips_value) if isinstance(twips_value, str) else twips_value
        return twips_int / 1440  # 1440 twips per inch
    except (ValueError, TypeError):
        return None


def half_points_to_points(half_points: str | int | None) -> float | None:
    """
    Convert OOXML half-points to points (for font size).

    OOXML font sizes are specified in half-points (e.g., 24 = 12pt).

    Args:
        half_points: Font size in half-points.

    Returns:
        Font size in points, or None if conversion fails.
    """
    if half_points is None or half_points == "":
        return None
    try:
        hp = int(half_points) if isinstance(half_points, str) else half_points
        return hp / 2
    except (ValueError, TypeError):
        return None


def emu_to_pixels(emu_value: int, emu_per_pixel: int = 9525) -> int | None:
    """
    Convert English Metric Units (EMU) to pixels.

    Used in XLSX for image dimensions.

    Args:
        emu_value: Value in EMU.
        emu_per_pixel: Conversion factor (default for screen: 9525 EMU/pixel).

    Returns:
        Value in pixels, or None on failure.
    """
    if emu_value is None:
        return None
    try:
        return int(emu_value / emu_per_pixel)
    except (ValueError, ZeroDivisionError, TypeError):
        return None
