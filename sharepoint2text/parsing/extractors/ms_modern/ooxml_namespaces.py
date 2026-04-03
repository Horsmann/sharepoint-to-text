"""
Centralized OOXML namespace definitions and pre-computed tag names.

This module consolidates all XML namespaces and pre-computed element tag names
used across DOCX, PPTX, and XLSX extractors. By centralizing these definitions,
we avoid duplication and reduce cognitive load.

All tag names are pre-computed (e.g., f"{NS}tag") to avoid repeated string
concatenation in hot code paths.
"""

# =============================================================================
# Common Namespaces (used across DOCX, PPTX, XLSX)
# =============================================================================

# DrawingML (shapes, images, text runs)
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Office relationships
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Math formulas (OMML)
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# Package-level core properties metadata
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

# Dublin Core metadata
DC_NS = "http://purl.org/dc/elements/1.1/"

# Dublin Core Terms
DCTERMS_NS = "http://purl.org/dc/terms/"

# Package-level relationships
PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# =============================================================================
# DOCX-Specific Namespaces
# =============================================================================

# WordprocessingML (main document format)
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Markup Compatibility (AlternateContent)
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# DrawingML Picture namespace
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# WordprocessingML Drawing Extensions (shapes)
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

# =============================================================================
# PPTX-Specific Namespaces
# =============================================================================

# PresentationML (main presentation format)
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# =============================================================================
# XLSX-Specific Namespaces
# =============================================================================

# Spreadsheet Drawing
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

# =============================================================================
# DOCX Pre-computed Tag Names
# =============================================================================

# Wrap namespace URIs in braces for ElementTree access
W_NS_B = f"{{{W_NS}}}"
M_NS_B = f"{{{M_NS}}}"
MC_NS_B = f"{{{MC_NS}}}"
R_NS_B = f"{{{R_NS}}}"
A_NS_B = f"{{{A_NS}}}"
CP_NS_B = f"{{{CP_NS}}}"
DC_NS_B = f"{{{DC_NS}}}"
DCTERMS_NS_B = f"{{{DCTERMS_NS}}}"
PIC_NS_B = f"{{{PIC_NS}}}"
WPS_NS_B = f"{{{WPS_NS}}}"

# WordprocessingML tags
W_T = f"{W_NS_B}t"
W_P = f"{W_NS_B}p"
W_R = f"{W_NS_B}r"
W_TBL = f"{W_NS_B}tbl"
W_TR = f"{W_NS_B}tr"
W_TC = f"{W_NS_B}tc"
W_PPR = f"{W_NS_B}pPr"
W_RPR = f"{W_NS_B}rPr"
W_PSTYLE = f"{W_NS_B}pStyle"
W_JC = f"{W_NS_B}jc"
W_VAL = f"{W_NS_B}val"
W_B = f"{W_NS_B}b"
W_I = f"{W_NS_B}i"
W_U = f"{W_NS_B}u"
W_SZ = f"{W_NS_B}sz"
W_COLOR = f"{W_NS_B}color"
W_RFONTS = f"{W_NS_B}rFonts"
W_DRAWING = f"{W_NS_B}drawing"
W_HYPERLINK = f"{W_NS_B}hyperlink"
W_FOOTNOTE = f"{W_NS_B}footnote"
W_ENDNOTE = f"{W_NS_B}endnote"
W_COMMENT = f"{W_NS_B}comment"
W_BODY = f"{W_NS_B}body"
W_SECTPR = f"{W_NS_B}sectPr"
W_BR = f"{W_NS_B}br"
W_TYPE = f"{W_NS_B}type"
W_LAST_RENDERED_PAGE_BREAK = f"{W_NS_B}lastRenderedPageBreak"
W_PGSZ = f"{W_NS_B}pgSz"
W_PGMAR = f"{W_NS_B}pgMar"
W_KEEPNEXT = f"{W_NS_B}keepNext"
W_STYLE = f"{W_NS_B}style"
W_STYLEID = f"{W_NS_B}styleId"
W_NAME = f"{W_NS_B}name"
W_ID = f"{W_NS_B}id"
W_AUTHOR = f"{W_NS_B}author"
W_DATE = f"{W_NS_B}date"
W_W = f"{W_NS_B}w"
W_H = f"{W_NS_B}h"
W_ORIENT = f"{W_NS_B}orient"
W_LEFT = f"{W_NS_B}left"
W_RIGHT = f"{W_NS_B}right"
W_TOP = f"{W_NS_B}top"
W_BOTTOM = f"{W_NS_B}bottom"
W_ASCII = f"{W_NS_B}ascii"
W_HANSI = f"{W_NS_B}hAnsi"
W_CS = f"{W_NS_B}cs"

# Math tags
M_OMATH = f"{{{M_NS}}}oMath"
M_OMATHPARA = f"{{{M_NS}}}oMathPara"

# Markup Compatibility
MC_CHOICE = f"{MC_NS_B}Choice"

# DrawingML tags
A_BLIP = f"{A_NS_B}blip"

# Relationships
R_ID = f"{R_NS_B}id"
R_EMBED = f"{R_NS_B}embed"

# Picture tags
PIC_CNVPR = f"{{{PIC_NS}}}cNvPr"

# WordprocessingML shapes
WPS_WSP = f"{WPS_NS_B}wsp"
WPS_TXBX = f"{WPS_NS_B}txbx"

# Metadata tags
_DC_TITLE = f"{DC_NS_B}title"
_DC_CREATOR = f"{DC_NS_B}creator"
_DC_SUBJECT = f"{DC_NS_B}subject"
_DC_DESCRIPTION = f"{DC_NS_B}description"
_CP_KEYWORDS = f"{CP_NS_B}keywords"
_CP_CATEGORY = f"{CP_NS_B}category"
_CP_LASTMODIFIEDBY = f"{CP_NS_B}lastModifiedBy"
_CP_REVISION = f"{CP_NS_B}revision"
_DCTERMS_CREATED = f"{DCTERMS_NS_B}created"
_DCTERMS_MODIFIED = f"{DCTERMS_NS_B}modified"

# =============================================================================
# PPTX Pre-computed Tag Names
# =============================================================================

P_NS_B = f"{{{P_NS}}}"
A_NS_B = f"{{{A_NS}}}"
M_NS_B_PPTX = f"{{{M_NS}}}"
R_NS_B_PPTX = f"{{{R_NS}}}"
CP_NS_B_PPTX = f"{{{CP_NS}}}"
DC_NS_B_PPTX = f"{{{DC_NS}}}"
DCTERMS_NS_B_PPTX = f"{{{DCTERMS_NS}}}"

# PresentationML tags
P_SP = f"{P_NS_B}sp"
P_PIC = f"{P_NS_B}pic"
P_SPTREE = f"{P_NS_B}spTree"
P_NVSPPR = f"{P_NS_B}nvSpPr"
P_NVPR = f"{P_NS_B}nvPr"
P_PH = f"{P_NS_B}ph"
P_TXBODY = f"{P_NS_B}txBody"
P_GRAPHICFRAME = f"{P_NS_B}graphicFrame"
P_CNVPR = f"{P_NS_B}cNvPr"
P_CM = f"{P_NS_B}cm"
P_TEXT = f"{P_NS_B}text"
P_SPPR = f"{P_NS_B}spPr"
P_XFRM = f"{P_NS_B}xfrm"
P_SLDID = f"{P_NS_B}sldId"
P_SLDIDLST = f"{P_NS_B}sldIdLst"

# DrawingML tags (shared with DOCX but used in PPTX context)
A_P = f"{A_NS_B}p"
A_R = f"{A_NS_B}r"
A_T = f"{A_NS_B}t"
A_BR = f"{A_NS_B}br"
A_FLD = f"{A_NS_B}fld"
A_BLIP_PPTX = f"{A_NS_B}blip"
A_XFRM = f"{A_NS_B}xfrm"
A_OFF = f"{A_NS_B}off"
A_TBL = f"{A_NS_B}tbl"
A_TR = f"{A_NS_B}tr"
A_TC = f"{A_NS_B}tc"
A_TXBODY = f"{A_NS_B}txBody"
A_GRAPHICDATA = f"{A_NS_B}graphicData"

# Math tags in PPTX context
M_OMATH_PPTX = f"{M_NS_B_PPTX}oMath"
M_OMATHPARA_PPTX = f"{M_NS_B_PPTX}oMathPara"

# Relationships in PPTX context
R_ID_PPTX = f"{R_NS_B_PPTX}id"
R_EMBED_PPTX = f"{R_NS_B_PPTX}embed"

# PPTX Metadata tags
_DC_TITLE_PPTX = f"{DC_NS_B_PPTX}title"
_DC_CREATOR_PPTX = f"{DC_NS_B_PPTX}creator"
_DC_SUBJECT_PPTX = f"{DC_NS_B_PPTX}subject"
_DC_DESCRIPTION_PPTX = f"{DC_NS_B_PPTX}description"
_CP_KEYWORDS_PPTX = f"{CP_NS_B_PPTX}keywords"
_CP_CATEGORY_PPTX = f"{CP_NS_B_PPTX}category"
_CP_LASTMODIFIEDBY_PPTX = f"{CP_NS_B_PPTX}lastModifiedBy"
_CP_REVISION_PPTX = f"{CP_NS_B_PPTX}revision"
_DCTERMS_CREATED_PPTX = f"{DCTERMS_NS_B_PPTX}created"
_DCTERMS_MODIFIED_PPTX = f"{DCTERMS_NS_B_PPTX}modified"

# =============================================================================
# XLSX Pre-computed Tag Names
# =============================================================================

XDR_NS_B = f"{{{XDR_NS}}}"

XDR_ONE_CELL_ANCHOR = f"{XDR_NS_B}oneCellAnchor"
XDR_TWO_CELL_ANCHOR = f"{XDR_NS_B}twoCellAnchor"
XDR_ABSOLUTE_ANCHOR = f"{XDR_NS_B}absoluteAnchor"
XDR_PIC = f"{XDR_NS_B}pic"
XDR_EXT = f"{XDR_NS_B}ext"
XDR_NVPICPR = f"{XDR_NS_B}nvPicPr"
XDR_CNVPR = f"{XDR_NS_B}cNvPr"
XDR_BLIPFILL = f"{XDR_NS_B}blipFill"
A_BLIP_XLSX = f"{A_NS_B}blip"
R_EMBED_XLSX = f"{R_NS_B}embed"

ANCHOR_TYPES = (XDR_ONE_CELL_ANCHOR, XDR_TWO_CELL_ANCHOR, XDR_ABSOLUTE_ANCHOR)

# =============================================================================
# DOCX Namespace Dictionary (for xmltodict-style access)
# =============================================================================

DOCX_NAMESPACES = {
    "w": W_NS,
    "m": M_NS,
    "mc": MC_NS,
    "r": R_NS,
    "a": A_NS,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "cp": CP_NS,
    "dc": DC_NS,
    "dcterms": DCTERMS_NS,
    "rel": PKG_RELS_NS,
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

# =============================================================================
# Constants for unit conversions and special values
# =============================================================================

# Unit conversions used in DOCX sections
EMU_PER_INCH = 914400  # English Metric Units per inch
TWIPS_PER_INCH = 1440  # Twips (1/20th of a point) per inch

# Unit conversions used in XLSX
EMU_PER_PIXEL = 9525  # English Metric Units per pixel (approximate)

# =============================================================================
# Caption and note filtering constants
# =============================================================================

# Caption style keywords for image caption detection
CAPTION_STYLE_KEYWORDS = ("caption", "bildunterschrift", "abbildung", "figure")

# Skip IDs for separator/continuation notes in DOCX
SKIP_NOTE_IDS = frozenset({"-1", "0"})

# =============================================================================
# Placeholder and content type categorization (PPTX)
# =============================================================================

# Title placeholder types
TITLE_TYPES = frozenset({"title", "ctrTitle"})

# Body/content placeholder types
BODY_TYPES = frozenset({"body", "subTitle", "obj", "tbl"})

# Footer-related placeholder types
FOOTER_TYPES = frozenset({"ftr"})

# Placeholder types to skip (not useful for text extraction)
# Note: sldNum (slide number) is NOT skipped - it goes to other_textboxes
SKIP_TYPES = frozenset({"dt", "sldImg", "hdr"})

# Table graphic data URI for PPTX
TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"

# =============================================================================
# XLSB parsing constants
# =============================================================================

# Record type for string shared strings table in XLSB
XLSB_SST_ITEM_RECORD = 19

# Datetime types for isinstance checking
DATETIME_TYPES = (
    int,
    float,
)  # Will be extended with actual datetime types in xlsx_extractor
