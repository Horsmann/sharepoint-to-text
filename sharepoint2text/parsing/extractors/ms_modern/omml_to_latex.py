"""
OMML to LaTeX Converter
=======================

Converts Office Math Markup Language (OMML) XML elements to LaTeX notation.

OMML is the XML format Microsoft Office uses to store mathematical formulas
in documents. This module provides utilities to convert OMML elements to
human-readable LaTeX strings.

Supported OMML Elements
-----------------------
- m:f (fraction) -> \\frac{num}{den}
- m:sSup/m:sSub (super/subscript) -> base^{sup} / base_{sub}
- m:sSubSup (sub-superscript) -> base_{sub}^{sup}
- m:rad (radical/root) -> \\sqrt{content} or \\sqrt[n]{content}
- m:nary (n-ary operators) -> \\sum, \\int, \\prod with limits
- m:d (delimiter) -> parentheses, brackets
- m:m (matrix) -> \\begin{matrix}...\\end{matrix}
- m:func (functions) -> \\sin, \\cos, etc.
- m:bar (overline) -> \\overline{content}
- m:acc (accent) -> \\hat, \\tilde, etc.

Greek Letter and Symbol Conversion
----------------------------------
The converter also handles:
- Greek letters (alpha, beta, gamma, etc.)
- Math symbols (infinity, partial, nabla, etc.)
- Set notation symbols
- Logical operators
- Arrows

Malformed Input Handling
------------------------
Some Word-generated OMML has malformed sqrt elements where the radical
contains only an opening bracket with the content following. This module
detects this pattern and consumes content until the matching closing
bracket is found.

Usage
-----
    >>> from xml.etree import ElementTree as ET
    >>> from sharepoint2text.parsing.extractors.util.omml_to_latex import omml_to_latex
    >>>
    >>> # Parse OMML from document XML
    >>> omath = ET.fromstring(omml_xml_string)
    >>> latex = omml_to_latex(omath)
    >>> print(latex)  # e.g., "\\frac{a}{b}"

See Also
--------
- docx_extractor: Uses this module for formula extraction
- pptx_extractor: Uses this module for formula extraction
"""

from xml.etree import ElementTree as ET

# OMML namespace
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# Greek letter and symbol mapping for LaTeX conversion
# Lowercase and uppercase Greek, plus common mathematical symbols
GREEK_TO_LATEX: dict[str, str] = {
    # Lowercase Greek
    "\u03b1": "\\alpha",
    "\u03b2": "\\beta",
    "\u03b3": "\\gamma",
    "\u03b4": "\\delta",
    "\u03b5": "\\epsilon",
    "\u03b6": "\\zeta",
    "\u03b7": "\\eta",
    "\u03b8": "\\theta",
    "\u03b9": "\\iota",
    "\u03ba": "\\kappa",
    "\u03bb": "\\lambda",
    "\u03bc": "\\mu",
    "\u03bd": "\\nu",
    "\u03be": "\\xi",
    "\u03bf": "o",  # omicron is just 'o' in LaTeX
    "\u03c0": "\\pi",
    "\u03c1": "\\rho",
    "\u03c3": "\\sigma",
    "\u03c2": "\\varsigma",
    "\u03c4": "\\tau",
    "\u03c5": "\\upsilon",
    "\u03c6": "\\phi",
    "\u03c7": "\\chi",
    "\u03c8": "\\psi",
    "\u03c9": "\\omega",
    # Uppercase Greek
    "\u0391": "A",
    "\u0392": "B",
    "\u0393": "\\Gamma",
    "\u0394": "\\Delta",
    "\u0395": "E",
    "\u0396": "Z",
    "\u0397": "H",
    "\u0398": "\\Theta",
    "\u0399": "I",
    "\u039a": "K",
    "\u039b": "\\Lambda",
    "\u039c": "M",
    "\u039d": "N",
    "\u039e": "\\Xi",
    "\u039f": "O",
    "\u03a0": "\\Pi",
    "\u03a1": "P",
    "\u03a3": "\\Sigma",
    "\u03a4": "T",
    "\u03a5": "\\Upsilon",
    "\u03a6": "\\Phi",
    "\u03a7": "X",
    "\u03a8": "\\Psi",
    "\u03a9": "\\Omega",
    # Common math symbols
    "\u221e": "\\infty",
    "\u2202": "\\partial",
    "\u2207": "\\nabla",
    "\u00b1": "\\pm",
    "\u2213": "\\mp",
    "\u00d7": "\\times",
    "\u00f7": "\\div",
    "\u00b7": "\\cdot",
    "\u2264": "\\leq",
    "\u2265": "\\geq",
    "\u2260": "\\neq",
    "\u2248": "\\approx",
    "\u2261": "\\equiv",
    "\u2208": "\\in",
    "\u2209": "\\notin",
    "\u2282": "\\subset",
    "\u2283": "\\supset",
    "\u2286": "\\subseteq",
    "\u2287": "\\supseteq",
    "\u222a": "\\cup",
    "\u2229": "\\cap",
    "\u2227": "\\land",
    "\u2228": "\\lor",
    "\u00ac": "\\neg",
    "\u2192": "\\rightarrow",
    "\u2190": "\\leftarrow",
    "\u2194": "\\leftrightarrow",
    "\u21d2": "\\Rightarrow",
    "\u21d0": "\\Leftarrow",
    "\u21d4": "\\Leftrightarrow",
    "\u2200": "\\forall",
    "\u2203": "\\exists",
    "\u2205": "\\emptyset",
    "\u2115": "\\mathbb{N}",
    "\u2124": "\\mathbb{Z}",
    "\u211a": "\\mathbb{Q}",
    "\u211d": "\\mathbb{R}",
    "\u2102": "\\mathbb{C}",
}

# Property elements to skip during conversion
_SKIP_TAGS = frozenset(
    {
        "rPr",
        "fPr",
        "radPr",
        "ctrlPr",
        "oMathParaPr",
        "degHide",
        "type",
        "rFonts",
        "i",
        "color",
        "sz",
        "szCs",
        "jc",
        "solidFill",
        "srgbClr",
        "latin",
    }
)


def convert_greek_and_symbols(text: str) -> str:
    """
    Convert Greek letters and math symbols to LaTeX equivalents.

    Args:
        text: Input string potentially containing Unicode Greek/math chars.

    Returns:
        String with Greek letters and symbols replaced by LaTeX commands.

    Example:
        >>> convert_greek_and_symbols("αβγ")
        '\\\\alpha\\\\beta\\\\gamma'
        >>> convert_greek_and_symbols("x + y")
        'x + y'
    """
    result = []
    for char in text:
        if char in GREEK_TO_LATEX:
            result.append(GREEK_TO_LATEX[char])
        else:
            result.append(char)
    return "".join(result)


def omml_to_latex(omath_element: ET.Element | None) -> str:
    """
    Convert an OMML (Office Math Markup Language) element to LaTeX notation.

    This function recursively processes the OMML XML structure and produces
    a LaTeX string representation suitable for rendering or display.

    Args:
        omath_element: An ElementTree Element representing an m:oMath or m:oMathPara
            element from the document XML. Can be None.

    Returns:
        LaTeX string representation of the mathematical expression.
        Returns empty string if omath_element is None.

    Supported OMML Elements:
        - m:f -> \\frac{numerator}{denominator}
        - m:sSup -> base^{superscript}
        - m:sSub -> base_{subscript}
        - m:sSubSup -> base_{sub}^{sup}
        - m:rad -> \\sqrt{content} or \\sqrt[n]{content}
        - m:nary -> \\sum, \\int, \\prod with limits
        - m:d -> (content) or other delimiters
        - m:m -> \\begin{matrix}...\\end{matrix}
        - m:func -> \\sin, \\cos, etc.
        - m:bar -> \\overline{content}
        - m:acc -> \\hat, \\tilde, etc.

    Malformed Input Handling:
        Some Word-generated OMML has malformed sqrt elements where the
        radical contains only an opening bracket with the content following.
        This function detects this pattern and consumes content until the
        matching closing bracket is found.

    Example:
        >>> from xml.etree import ElementTree as ET
        >>> # Simple fraction
        >>> xml = '''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        ...   <m:f>
        ...     <m:num><m:r><m:t>a</m:t></m:r></m:num>
        ...     <m:den><m:r><m:t>b</m:t></m:r></m:den>
        ...   </m:f>
        ... </m:oMath>'''
        >>> omath = ET.fromstring(xml)
        >>> omml_to_latex(omath)
        '\\\\frac{a}{b}'
    """
    if omath_element is None:
        return ""

    parts: list[str] = []
    pending_sqrt_close: str | None = None  # Bracket needed to close current sqrt

    def process_element(elem: ET.Element | None) -> str:
        """Recursively process an element and return its LaTeX representation."""
        nonlocal pending_sqrt_close

        if elem is None:
            return ""

        tag = elem.tag.split("}")[-1]

        # Skip property elements
        if tag in _SKIP_TAGS:
            return ""

        # Text content (both w:t and m:t)
        if tag == "t":
            text = elem.text or ""
            converted = convert_greek_and_symbols(text)

            # Handle malformed sqrt: if we're waiting for a closing bracket
            if pending_sqrt_close and pending_sqrt_close in converted:
                idx = converted.index(pending_sqrt_close)
                inside = converted[:idx]  # Content inside sqrt
                outside = converted[idx + 1 :]  # Content after closing bracket
                pending_sqrt_close = None
                return inside + "}" + outside

            return converted

        # Fraction: m:f contains m:num (numerator) and m:den (denominator)
        if tag == "f":
            num = elem.find(f"{M_NS}num")
            den = elem.find(f"{M_NS}den")
            num_text = process_element(num)
            den_text = process_element(den)
            return f"\\frac{{{num_text}}}{{{den_text}}}"

        # Superscript: m:sSup contains m:e (base) and m:sup (superscript)
        if tag == "sSup":
            base = elem.find(f"{M_NS}e")
            sup = elem.find(f"{M_NS}sup")
            base_text = process_element(base)
            sup_text = process_element(sup)
            return f"{base_text}^{{{sup_text}}}"

        # Subscript: m:sSub contains m:e (base) and m:sub (subscript)
        if tag == "sSub":
            base = elem.find(f"{M_NS}e")
            sub = elem.find(f"{M_NS}sub")
            base_text = process_element(base)
            sub_text = process_element(sub)
            return f"{base_text}_{{{sub_text}}}"

        # Sub-superscript: m:sSubSup contains m:e, m:sub, and m:sup
        if tag == "sSubSup":
            base = elem.find(f"{M_NS}e")
            sub = elem.find(f"{M_NS}sub")
            sup = elem.find(f"{M_NS}sup")
            base_text = process_element(base)
            sub_text = process_element(sub)
            sup_text = process_element(sup)
            return f"{base_text}_{{{sub_text}}}^{{{sup_text}}}"

        # Square root: m:rad contains m:deg (degree, optional) and m:e (content)
        if tag == "rad":
            deg = elem.find(f"{M_NS}deg")
            content = elem.find(f"{M_NS}e")
            content_text = process_element(content)
            deg_text = process_element(deg).strip()

            # Handle malformed: lone opening bracket inside sqrt
            # Some OMML has sqrt containing just "(" with content after
            if content_text.strip() in ("(", "[", "{"):
                bracket_map = {"(": ")", "[": "]", "{": "}"}
                pending_sqrt_close = bracket_map.get(content_text.strip(), ")")
                if deg_text:
                    return f"\\sqrt[{deg_text}]{{"
                else:
                    return "\\sqrt{"
            else:
                if deg_text:
                    return f"\\sqrt[{deg_text}]{{{content_text}}}"
                else:
                    return f"\\sqrt{{{content_text}}}"

        # N-ary (sum, product, integral): m:nary
        if tag == "nary":
            chr_elem = elem.find(f".//{M_NS}chr")
            op = chr_elem.get(f"{M_NS}val") if chr_elem is not None else "\u2211"

            sub = elem.find(f"{M_NS}sub")
            sup = elem.find(f"{M_NS}sup")
            content = elem.find(f"{M_NS}e")

            op_map = {
                "\u2211": "\\sum",
                "\u220f": "\\prod",
                "\u222b": "\\int",
                "\u222c": "\\iint",
                "\u222d": "\\iiint",
            }
            latex_op = op_map.get(op, convert_greek_and_symbols(op))

            sub_text = process_element(sub)
            sup_text = process_element(sup)
            content_text = process_element(content)

            result = latex_op
            if sub_text.strip():
                result += f"_{{{sub_text}}}"
            if sup_text.strip():
                result += f"^{{{sup_text}}}"
            result += f" {content_text}"
            return result

        # Delimiter (parentheses, brackets): m:d
        if tag == "d":
            beg_chr = elem.find(f".//{M_NS}begChr")
            end_chr = elem.find(f".//{M_NS}endChr")
            left = beg_chr.get(f"{M_NS}val") if beg_chr is not None else "("
            right = end_chr.get(f"{M_NS}val") if end_chr is not None else ")"

            e_elements = elem.findall(f"{M_NS}e")
            content_parts = [process_element(e) for e in e_elements]
            content_text = ", ".join(content_parts)
            return f"{left}{content_text}{right}"

        # Matrix: m:m contains m:mr (rows) which contain m:e (elements)
        if tag == "m" and elem.find(f"{M_NS}mr") is not None:
            rows = []
            for mr in elem.findall(f"{M_NS}mr"):
                cells = [process_element(e) for e in mr.findall(f"{M_NS}e")]
                rows.append(" & ".join(cells))
            return "\\begin{matrix}" + " \\\\ ".join(rows) + "\\end{matrix}"

        # Function: m:func contains m:fName and m:e
        if tag == "func":
            fname = elem.find(f"{M_NS}fName")
            content = elem.find(f"{M_NS}e")
            fname_text = process_element(fname)
            content_text = process_element(content)
            func_map = {
                "sin": "\\sin",
                "cos": "\\cos",
                "tan": "\\tan",
                "log": "\\log",
                "ln": "\\ln",
                "lim": "\\lim",
                "exp": "\\exp",
                "max": "\\max",
                "min": "\\min",
            }
            latex_fname = func_map.get(fname_text.strip(), fname_text)
            return f"{latex_fname}{{{content_text}}}"

        # Bar/overline: m:bar
        if tag == "bar":
            content = elem.find(f"{M_NS}e")
            content_text = process_element(content)
            return f"\\overline{{{content_text}}}"

        # Accent (hat, tilde, etc.): m:acc
        if tag == "acc":
            chr_elem = elem.find(f".//{M_NS}chr")
            accent = chr_elem.get(f"{M_NS}val") if chr_elem is not None else "^"
            content = elem.find(f"{M_NS}e")
            content_text = process_element(content)

            accent_map = {
                "\u0302": "\\hat",
                "\u0303": "\\tilde",
                "\u0304": "\\bar",
                "\u20d7": "\\vec",
                "\u0307": "\\dot",
            }
            latex_accent = accent_map.get(accent, "\\hat")
            return f"{latex_accent}{{{content_text}}}"

        # Default: recurse into children and concatenate results
        result = []
        for child in elem:
            child_result = process_element(child)
            if child_result:
                result.append(child_result)
        return "".join(result)

    # Process all children of the omath element
    for child in omath_element:
        child_result = process_element(child)
        if child_result:
            parts.append(child_result)

    # If sqrt was never closed (no matching bracket found), close it now
    if pending_sqrt_close:
        parts.append("}")

    return "".join(parts)
