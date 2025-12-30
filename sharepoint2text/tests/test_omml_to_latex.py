"""
Unit tests for OMML to LaTeX conversion.

This module provides extensive tests for the omml_to_latex module which converts
Office Math Markup Language (OMML) XML elements to LaTeX notation.

Tests cover:
    - Greek letter and symbol conversion
    - Fractions
    - Superscripts and subscripts
    - Sub-superscripts
    - Square roots (with and without degrees)
    - N-ary operators (sum, product, integral)
    - Delimiters (parentheses, brackets)
    - Matrices
    - Functions (sin, cos, tan, log, etc.)
    - Bar/overline
    - Accents (hat, tilde, bar, vec, dot)
    - Complex nested expressions
    - Edge cases and malformed input
"""

from unittest import TestCase
from xml.etree import ElementTree as ET

from sharepoint2text.extractors.omml_to_latex import (
    GREEK_TO_LATEX,
    convert_greek_and_symbols,
    omml_to_latex,
)

tc = TestCase()


# Helper function to create OMML XML for testing
def make_omath(inner_xml: str) -> ET.Element:
    """Create an oMath element with the given inner XML content."""
    xml = f"""<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        {inner_xml}
    </m:oMath>"""
    return ET.fromstring(xml)


# =============================================================================
# Tests for Greek Letter and Symbol Conversion
# =============================================================================


def test_convert_greek_lowercase_alpha():
    """Test conversion of lowercase alpha."""
    result = convert_greek_and_symbols("α")
    tc.assertEqual("\\alpha", result)


def test_convert_greek_lowercase_beta():
    """Test conversion of lowercase beta."""
    result = convert_greek_and_symbols("β")
    tc.assertEqual("\\beta", result)


def test_convert_greek_lowercase_gamma():
    """Test conversion of lowercase gamma."""
    result = convert_greek_and_symbols("γ")
    tc.assertEqual("\\gamma", result)


def test_convert_greek_lowercase_all():
    """Test conversion of multiple lowercase Greek letters."""
    result = convert_greek_and_symbols("αβγδεζηθ")
    tc.assertEqual("\\alpha\\beta\\gamma\\delta\\epsilon\\zeta\\eta\\theta", result)


def test_convert_greek_uppercase_gamma():
    """Test conversion of uppercase Gamma."""
    result = convert_greek_and_symbols("Γ")
    tc.assertEqual("\\Gamma", result)


def test_convert_greek_uppercase_delta():
    """Test conversion of uppercase Delta."""
    result = convert_greek_and_symbols("Δ")
    tc.assertEqual("\\Delta", result)


def test_convert_greek_uppercase_omega():
    """Test conversion of uppercase Omega."""
    result = convert_greek_and_symbols("Ω")
    tc.assertEqual("\\Omega", result)


def test_convert_greek_omicron_lowercase():
    """Test that lowercase omicron converts to 'o'."""
    result = convert_greek_and_symbols("ο")
    tc.assertEqual("o", result)


def test_convert_symbol_infinity():
    """Test conversion of infinity symbol."""
    result = convert_greek_and_symbols("∞")
    tc.assertEqual("\\infty", result)


def test_convert_symbol_partial():
    """Test conversion of partial derivative symbol."""
    result = convert_greek_and_symbols("∂")
    tc.assertEqual("\\partial", result)


def test_convert_symbol_nabla():
    """Test conversion of nabla symbol."""
    result = convert_greek_and_symbols("∇")
    tc.assertEqual("\\nabla", result)


def test_convert_symbol_pm():
    """Test conversion of plus-minus symbol."""
    result = convert_greek_and_symbols("±")
    tc.assertEqual("\\pm", result)


def test_convert_symbol_times():
    """Test conversion of multiplication symbol."""
    result = convert_greek_and_symbols("×")
    tc.assertEqual("\\times", result)


def test_convert_symbol_leq():
    """Test conversion of less-than-or-equal symbol."""
    result = convert_greek_and_symbols("≤")
    tc.assertEqual("\\leq", result)


def test_convert_symbol_geq():
    """Test conversion of greater-than-or-equal symbol."""
    result = convert_greek_and_symbols("≥")
    tc.assertEqual("\\geq", result)


def test_convert_symbol_neq():
    """Test conversion of not-equal symbol."""
    result = convert_greek_and_symbols("≠")
    tc.assertEqual("\\neq", result)


def test_convert_symbol_in():
    """Test conversion of 'in' (element of) symbol."""
    result = convert_greek_and_symbols("∈")
    tc.assertEqual("\\in", result)


def test_convert_symbol_subset():
    """Test conversion of subset symbol."""
    result = convert_greek_and_symbols("⊂")
    tc.assertEqual("\\subset", result)


def test_convert_symbol_cup():
    """Test conversion of union symbol."""
    result = convert_greek_and_symbols("∪")
    tc.assertEqual("\\cup", result)


def test_convert_symbol_cap():
    """Test conversion of intersection symbol."""
    result = convert_greek_and_symbols("∩")
    tc.assertEqual("\\cap", result)


def test_convert_symbol_forall():
    """Test conversion of forall symbol."""
    result = convert_greek_and_symbols("∀")
    tc.assertEqual("\\forall", result)


def test_convert_symbol_exists():
    """Test conversion of exists symbol."""
    result = convert_greek_and_symbols("∃")
    tc.assertEqual("\\exists", result)


def test_convert_symbol_emptyset():
    """Test conversion of empty set symbol."""
    result = convert_greek_and_symbols("∅")
    tc.assertEqual("\\emptyset", result)


def test_convert_symbol_rightarrow():
    """Test conversion of right arrow symbol."""
    result = convert_greek_and_symbols("→")
    tc.assertEqual("\\rightarrow", result)


def test_convert_symbol_leftarrow():
    """Test conversion of left arrow symbol."""
    result = convert_greek_and_symbols("←")
    tc.assertEqual("\\leftarrow", result)


def test_convert_symbol_double_rightarrow():
    """Test conversion of double right arrow symbol."""
    result = convert_greek_and_symbols("⇒")
    tc.assertEqual("\\Rightarrow", result)


def test_convert_symbol_mathbb_naturals():
    """Test conversion of natural numbers symbol."""
    result = convert_greek_and_symbols("ℕ")
    tc.assertEqual("\\mathbb{N}", result)


def test_convert_symbol_mathbb_reals():
    """Test conversion of real numbers symbol."""
    result = convert_greek_and_symbols("ℝ")
    tc.assertEqual("\\mathbb{R}", result)


def test_convert_mixed_text_and_symbols():
    """Test conversion of mixed text with Greek letters and symbols."""
    result = convert_greek_and_symbols("x + α = β × γ")
    tc.assertEqual("x + \\alpha = \\beta \\times \\gamma", result)


def test_convert_no_special_characters():
    """Test that regular text passes through unchanged."""
    result = convert_greek_and_symbols("Hello World 123")
    tc.assertEqual("Hello World 123", result)


def test_convert_empty_string():
    """Test conversion of empty string."""
    result = convert_greek_and_symbols("")
    tc.assertEqual("", result)


# =============================================================================
# Tests for OMML to LaTeX - Fractions
# =============================================================================


def test_simple_fraction():
    """Test simple fraction a/b."""
    omath = make_omath(
        """
        <m:f>
            <m:num><m:r><m:t>a</m:t></m:r></m:num>
            <m:den><m:r><m:t>b</m:t></m:r></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{a}{b}", result)


def test_fraction_with_numbers():
    """Test fraction with numbers 3/4."""
    omath = make_omath(
        """
        <m:f>
            <m:num><m:r><m:t>3</m:t></m:r></m:num>
            <m:den><m:r><m:t>4</m:t></m:r></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{3}{4}", result)


def test_nested_fraction():
    """Test nested fraction (a/b)/(c/d)."""
    omath = make_omath(
        """
        <m:f>
            <m:num>
                <m:f>
                    <m:num><m:r><m:t>a</m:t></m:r></m:num>
                    <m:den><m:r><m:t>b</m:t></m:r></m:den>
                </m:f>
            </m:num>
            <m:den>
                <m:f>
                    <m:num><m:r><m:t>c</m:t></m:r></m:num>
                    <m:den><m:r><m:t>d</m:t></m:r></m:den>
                </m:f>
            </m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{\\frac{a}{b}}{\\frac{c}{d}}", result)


def test_fraction_with_greek_letters():
    """Test fraction with Greek letters α/β."""
    omath = make_omath(
        """
        <m:f>
            <m:num><m:r><m:t>α</m:t></m:r></m:num>
            <m:den><m:r><m:t>β</m:t></m:r></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{\\alpha}{\\beta}", result)


# =============================================================================
# Tests for OMML to LaTeX - Superscripts and Subscripts
# =============================================================================


def test_simple_superscript():
    """Test simple superscript x^2."""
    omath = make_omath(
        """
        <m:sSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("x^{2}", result)


def test_simple_subscript():
    """Test simple subscript x_i."""
    omath = make_omath(
        """
        <m:sSub>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
        </m:sSub>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("x_{i}", result)


def test_sub_superscript():
    """Test sub-superscript x_i^2."""
    omath = make_omath(
        """
        <m:sSubSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSubSup>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("x_{i}^{2}", result)


def test_nested_superscript():
    """Test nested superscript e^(x^2)."""
    omath = make_omath(
        """
        <m:sSup>
            <m:e><m:r><m:t>e</m:t></m:r></m:e>
            <m:sup>
                <m:sSup>
                    <m:e><m:r><m:t>x</m:t></m:r></m:e>
                    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
            </m:sup>
        </m:sSup>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("e^{x^{2}}", result)


def test_superscript_with_greek():
    """Test superscript with Greek letter σ^2."""
    omath = make_omath(
        """
        <m:sSup>
            <m:e><m:r><m:t>σ</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sigma^{2}", result)


# =============================================================================
# Tests for OMML to LaTeX - Square Roots
# =============================================================================


def test_simple_sqrt():
    """Test simple square root sqrt(x)."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:rad>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{x}", result)


def test_sqrt_with_expression():
    """Test square root with expression sqrt(a+b)."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>a+b</m:t></m:r></m:e>
        </m:rad>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{a+b}", result)


def test_nth_root():
    """Test nth root sqrt[3](x)."""
    omath = make_omath(
        """
        <m:rad>
            <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:rad>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt[3]{x}", result)


def test_nested_sqrt():
    """Test nested square root sqrt(sqrt(x))."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e>
                <m:rad>
                    <m:radPr><m:degHide m:val="1"/></m:radPr>
                    <m:deg/>
                    <m:e><m:r><m:t>x</m:t></m:r></m:e>
                </m:rad>
            </m:e>
        </m:rad>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{\\sqrt{x}}", result)


# =============================================================================
# Tests for OMML to LaTeX - N-ary Operators
# =============================================================================


def test_summation_simple():
    """Test simple summation without limits."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
            </m:naryPr>
            <m:sub/>
            <m:sup/>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sum x", result)


def test_summation_with_limits():
    """Test summation with limits sum_{i=1}^{n}."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
            </m:naryPr>
            <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sum_{i=1}^{n} x", result)


def test_product_with_limits():
    """Test product with limits prod_{i=1}^{n}."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∏"/>
            </m:naryPr>
            <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\prod_{i=1}^{n} a", result)


def test_integral_simple():
    """Test simple integral."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∫"/>
            </m:naryPr>
            <m:sub/>
            <m:sup/>
            <m:e><m:r><m:t>f(x)dx</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\int f(x)dx", result)


def test_integral_with_limits():
    """Test integral with limits int_{a}^{b}."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∫"/>
            </m:naryPr>
            <m:sub><m:r><m:t>a</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>b</m:t></m:r></m:sup>
            <m:e><m:r><m:t>f(x)dx</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\int_{a}^{b} f(x)dx", result)


def test_double_integral():
    """Test double integral."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∬"/>
            </m:naryPr>
            <m:sub/>
            <m:sup/>
            <m:e><m:r><m:t>f(x,y)dxdy</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\iint f(x,y)dxdy", result)


def test_triple_integral():
    """Test triple integral."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∭"/>
            </m:naryPr>
            <m:sub/>
            <m:sup/>
            <m:e><m:r><m:t>f(x,y,z)dxdydz</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\iiint f(x,y,z)dxdydz", result)


# =============================================================================
# Tests for OMML to LaTeX - Delimiters
# =============================================================================


def test_parentheses():
    """Test parentheses (x+y)."""
    omath = make_omath(
        """
        <m:d>
            <m:dPr>
                <m:begChr m:val="("/>
                <m:endChr m:val=")"/>
            </m:dPr>
            <m:e><m:r><m:t>x+y</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("(x+y)", result)


def test_square_brackets():
    """Test square brackets [x+y]."""
    omath = make_omath(
        """
        <m:d>
            <m:dPr>
                <m:begChr m:val="["/>
                <m:endChr m:val="]"/>
            </m:dPr>
            <m:e><m:r><m:t>x+y</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("[x+y]", result)


def test_curly_braces():
    """Test curly braces {x+y}."""
    omath = make_omath(
        """
        <m:d>
            <m:dPr>
                <m:begChr m:val="{"/>
                <m:endChr m:val="}"/>
            </m:dPr>
            <m:e><m:r><m:t>x+y</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("{x+y}", result)


def test_absolute_value():
    """Test absolute value |x|."""
    omath = make_omath(
        """
        <m:d>
            <m:dPr>
                <m:begChr m:val="|"/>
                <m:endChr m:val="|"/>
            </m:dPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("|x|", result)


def test_delimiter_multiple_elements():
    """Test delimiter with multiple elements (x, y, z)."""
    omath = make_omath(
        """
        <m:d>
            <m:dPr>
                <m:begChr m:val="("/>
                <m:endChr m:val=")"/>
            </m:dPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:e><m:r><m:t>y</m:t></m:r></m:e>
            <m:e><m:r><m:t>z</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("(x, y, z)", result)


# =============================================================================
# Tests for OMML to LaTeX - Matrices
# =============================================================================


def test_simple_matrix():
    """Test simple 2x2 matrix."""
    omath = make_omath(
        """
        <m:m>
            <m:mr>
                <m:e><m:r><m:t>a</m:t></m:r></m:e>
                <m:e><m:r><m:t>b</m:t></m:r></m:e>
            </m:mr>
            <m:mr>
                <m:e><m:r><m:t>c</m:t></m:r></m:e>
                <m:e><m:r><m:t>d</m:t></m:r></m:e>
            </m:mr>
        </m:m>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\begin{matrix}a & b \\\\ c & d\\end{matrix}", result)


def test_identity_matrix():
    """Test 3x3 identity matrix."""
    omath = make_omath(
        """
        <m:m>
            <m:mr>
                <m:e><m:r><m:t>1</m:t></m:r></m:e>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
            </m:mr>
            <m:mr>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
                <m:e><m:r><m:t>1</m:t></m:r></m:e>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
            </m:mr>
            <m:mr>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
                <m:e><m:r><m:t>0</m:t></m:r></m:e>
                <m:e><m:r><m:t>1</m:t></m:r></m:e>
            </m:mr>
        </m:m>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual(
        "\\begin{matrix}1 & 0 & 0 \\\\ 0 & 1 & 0 \\\\ 0 & 0 & 1\\end{matrix}", result
    )


# =============================================================================
# Tests for OMML to LaTeX - Functions
# =============================================================================


def test_sin_function():
    """Test sin function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sin{x}", result)


def test_cos_function():
    """Test cos function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>cos</m:t></m:r></m:fName>
            <m:e><m:r><m:t>θ</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\cos{\\theta}", result)


def test_tan_function():
    """Test tan function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>tan</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\tan{x}", result)


def test_log_function():
    """Test log function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>log</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\log{x}", result)


def test_ln_function():
    """Test ln (natural log) function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>ln</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\ln{x}", result)


def test_lim_function():
    """Test lim function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>lim</m:t></m:r></m:fName>
            <m:e><m:r><m:t>f(x)</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\lim{f(x)}", result)


def test_exp_function():
    """Test exp function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>exp</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\exp{x}", result)


def test_max_function():
    """Test max function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>max</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x,y</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\max{x,y}", result)


def test_min_function():
    """Test min function."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>min</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x,y</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\min{x,y}", result)


def test_unknown_function():
    """Test unknown function passes through."""
    omath = make_omath(
        """
        <m:func>
            <m:fName><m:r><m:t>myFunc</m:t></m:r></m:fName>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("myFunc{x}", result)


# =============================================================================
# Tests for OMML to LaTeX - Bar and Overline
# =============================================================================


def test_overline():
    """Test overline/bar."""
    omath = make_omath(
        """
        <m:bar>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:bar>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\overline{x}", result)


def test_overline_expression():
    """Test overline with expression."""
    omath = make_omath(
        """
        <m:bar>
            <m:e><m:r><m:t>AB</m:t></m:r></m:e>
        </m:bar>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\overline{AB}", result)


# =============================================================================
# Tests for OMML to LaTeX - Accents
# =============================================================================


def test_accent_hat():
    """Test hat accent."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="̂"/>
            </m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\hat{x}", result)


def test_accent_tilde():
    """Test tilde accent."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="̃"/>
            </m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\tilde{x}", result)


def test_accent_bar():
    """Test bar accent (different from overline element)."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="̄"/>
            </m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\bar{x}", result)


def test_accent_vec():
    """Test vector accent."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="⃗"/>
            </m:accPr>
            <m:e><m:r><m:t>v</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\vec{v}", result)


def test_accent_dot():
    """Test dot accent (time derivative)."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="̇"/>
            </m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\dot{x}", result)


def test_accent_default():
    """Test default accent (hat) for unknown accent character."""
    omath = make_omath(
        """
        <m:acc>
            <m:accPr>
                <m:chr m:val="?"/>
            </m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\hat{x}", result)


# =============================================================================
# Tests for OMML to LaTeX - Complex Expressions
# =============================================================================


def test_quadratic_formula():
    """Test quadratic formula: (-b ± sqrt(b^2 - 4ac)) / 2a."""
    # This is a simplified version focusing on the fraction and sqrt
    omath = make_omath(
        """
        <m:f>
            <m:num>
                <m:r><m:t>-b±</m:t></m:r>
                <m:rad>
                    <m:radPr><m:degHide m:val="1"/></m:radPr>
                    <m:deg/>
                    <m:e>
                        <m:sSup>
                            <m:e><m:r><m:t>b</m:t></m:r></m:e>
                            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                        </m:sSup>
                        <m:r><m:t>-4ac</m:t></m:r>
                    </m:e>
                </m:rad>
            </m:num>
            <m:den><m:r><m:t>2a</m:t></m:r></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}", result)


def test_gaussian_distribution():
    """Test Gaussian distribution formula component."""
    omath = make_omath(
        """
        <m:f>
            <m:num><m:r><m:t>1</m:t></m:r></m:num>
            <m:den>
                <m:rad>
                    <m:radPr><m:degHide m:val="1"/></m:radPr>
                    <m:deg/>
                    <m:e>
                        <m:r><m:t>2π</m:t></m:r>
                        <m:sSup>
                            <m:e><m:r><m:t>σ</m:t></m:r></m:e>
                            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                        </m:sSup>
                    </m:e>
                </m:rad>
            </m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}", result)


def test_summation_with_fraction():
    """Test summation with fraction in body."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
            </m:naryPr>
            <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
            <m:e>
                <m:f>
                    <m:num><m:r><m:t>1</m:t></m:r></m:num>
                    <m:den><m:r><m:t>i</m:t></m:r></m:den>
                </m:f>
            </m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sum_{i=1}^{n} \\frac{1}{i}", result)


# =============================================================================
# Tests for OMML to LaTeX - Edge Cases
# =============================================================================


def test_none_input():
    """Test that None input returns empty string."""
    result = omml_to_latex(None)
    tc.assertEqual("", result)


def test_empty_omath():
    """Test empty oMath element."""
    omath = make_omath("")
    result = omml_to_latex(omath)
    tc.assertEqual("", result)


def test_text_only():
    """Test oMath with just text."""
    omath = make_omath("<m:r><m:t>Hello World</m:t></m:r>")
    result = omml_to_latex(omath)
    tc.assertEqual("Hello World", result)


def test_skip_property_elements():
    """Test that property elements are skipped."""
    omath = make_omath(
        """
        <m:r>
            <m:rPr>
                <m:sty m:val="p"/>
            </m:rPr>
            <m:t>x</m:t>
        </m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("x", result)


def test_multiple_runs():
    """Test multiple runs concatenate."""
    omath = make_omath(
        """
        <m:r><m:t>a</m:t></m:r>
        <m:r><m:t>+</m:t></m:r>
        <m:r><m:t>b</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("a+b", result)


def test_greek_to_latex_dict_completeness():
    """Test that all expected Greek letters are in the mapping."""
    # Test lowercase Greek letters are present
    lowercase = "αβγδεζηθικλμνξοπρστυφχψω"
    for letter in lowercase:
        tc.assertIn(letter, GREEK_TO_LATEX)

    # Test uppercase Greek letters with unique LaTeX commands
    uppercase_with_commands = "ΓΔΘΛΞΠΣΥΦΨΩ"
    for letter in uppercase_with_commands:
        tc.assertIn(letter, GREEK_TO_LATEX)


def test_delimiter_default_parentheses():
    """Test that delimiter defaults to parentheses when not specified."""
    omath = make_omath(
        """
        <m:d>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:d>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("(x)", result)


def test_fraction_with_empty_numerator():
    """Test fraction with empty numerator."""
    omath = make_omath(
        """
        <m:f>
            <m:num></m:num>
            <m:den><m:r><m:t>b</m:t></m:r></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{}{b}", result)


def test_fraction_with_empty_denominator():
    """Test fraction with empty denominator."""
    omath = make_omath(
        """
        <m:f>
            <m:num><m:r><m:t>a</m:t></m:r></m:num>
            <m:den></m:den>
        </m:f>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\frac{a}{}", result)


def test_sub_superscript_empty_sub():
    """Test sub-superscript with empty subscript."""
    omath = make_omath(
        """
        <m:sSubSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sub></m:sub>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSubSup>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("x_{}^{2}", result)


def test_nary_only_subscript():
    """Test n-ary operator with only subscript."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
            </m:naryPr>
            <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
            <m:sup/>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sum_{i=1} x", result)


def test_nary_only_superscript():
    """Test n-ary operator with only superscript."""
    omath = make_omath(
        """
        <m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
            </m:naryPr>
            <m:sub/>
            <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:nary>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sum^{n} x", result)


# =============================================================================
# Tests for OMML to LaTeX - Malformed Input Handling
# =============================================================================


def test_malformed_sqrt_opening_paren():
    """Test malformed sqrt with just opening parenthesis."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>(</m:t></m:r></m:e>
        </m:rad>
        <m:r><m:t>x+y)</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{x+y}", result)


def test_malformed_sqrt_opening_bracket():
    """Test malformed sqrt with just opening bracket."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>[</m:t></m:r></m:e>
        </m:rad>
        <m:r><m:t>x+y]</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{x+y}", result)


def test_malformed_sqrt_opening_brace():
    """Test malformed sqrt with just opening brace."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>{</m:t></m:r></m:e>
        </m:rad>
        <m:r><m:t>x+y}</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt{x+y}", result)


def test_malformed_sqrt_unclosed():
    """Test malformed sqrt without closing bracket in content."""
    omath = make_omath(
        """
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>(</m:t></m:r></m:e>
        </m:rad>
        <m:r><m:t>x+y</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    # Should still close the sqrt even without matching bracket
    tc.assertEqual("\\sqrt{x+y}", result)


def test_malformed_nth_root_opening_paren():
    """Test malformed nth root with just opening parenthesis."""
    omath = make_omath(
        """
        <m:rad>
            <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
            <m:e><m:r><m:t>(</m:t></m:r></m:e>
        </m:rad>
        <m:r><m:t>x+y)</m:t></m:r>
    """
    )
    result = omml_to_latex(omath)
    tc.assertEqual("\\sqrt[3]{x+y}", result)
