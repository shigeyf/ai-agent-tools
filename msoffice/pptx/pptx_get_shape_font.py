# pptx_get_shape_font.py
# -*- coding: utf-8 -*-

"""
Font resolution utilities for PowerPoint shapes.

This module provides functions to extract font information from PowerPoint
presentations and shapes, following the Open XML (ISO/IEC 29500-1) specification
for font inheritance and theme font resolution.
"""

from typing import Any, Dict, Literal, Optional, TypedDict

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn


# DrawingML namespace constant
_DRAWINGML_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

# Line spacing type constants
# Per ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing):
# "This can be specified in two different ways, percentage spacing and font point spacing."
LINE_SPACING_TYPE_RATIO = "ratio"  # spcPct: percentage-based spacing (§21.1.2.2.11)
LINE_SPACING_TYPE_FIXED = "fixed_pt"  # spcPts: fixed point spacing (§21.1.2.2.12)
LINE_SPACING_TYPE_DEFAULT = "default"  # No explicit line spacing set


# ---------------------------
# Sub functions
# ---------------------------


def _resolve_theme_font_reference(
    typeface: str, theme_fonts: Optional[Dict[str, Optional[str]]]
) -> Optional[str]:
    """
    Resolve theme font reference to actual font name.

    Theme font references use the format:
    - +mj-lt: major Latin font
    - +mn-lt: minor Latin font
    - +mj-ea: major East Asian font
    - +mn-ea: minor East Asian font

    Args:
        typeface: Font typeface string (may be a theme reference like "+mj-ea").
        theme_fonts: Dictionary containing theme font names.
    Returns:
        Resolved font name, or None if cannot resolve.
    """
    if not typeface or not theme_fonts:
        return None

    # Theme font reference mapping
    theme_ref_map = {
        "+mj-lt": "major_latin",
        "+mn-lt": "minor_latin",
        "+mj-ea": "major_ea",
        "+mn-ea": "minor_ea",
    }

    if typeface in theme_ref_map:
        theme_key = theme_ref_map[typeface]
        return theme_fonts.get(theme_key)

    return None


# ---------------------------
# Exportable functions
# ---------------------------


def get_theme_fonts(prs: Any) -> Dict[str, Optional[str]]:
    """
    Get theme fonts from PowerPoint presentation.

    Extracts the font scheme from the presentation's theme, which defines
    the default fonts for headings (major) and body text (minor) in both
    Latin and East Asian scripts.

    Args:
        prs: PowerPoint Presentation object.
    Returns:
        Dictionary with 'major_latin', 'major_ea', 'minor_latin', 'minor_ea' keys.
        Values are font names or None if not defined.
    """
    result: Dict[str, Optional[str]] = {
        "major_latin": None,
        "major_ea": None,
        "minor_latin": None,
        "minor_ea": None,
    }

    try:
        master = prs.slide_masters[0]
        master_part = master.part
        theme_part = master_part.part_related_by(RT.THEME)

        # Parse theme XML
        # pylint: disable=protected-access
        theme_xml = theme_part._blob
        # pylint: disable=c-extension-no-member
        root = etree.fromstring(theme_xml)

        # Find fontScheme
        font_scheme = root.find(".//a:fontScheme", _DRAWINGML_NS)

        if font_scheme is not None:
            # Major fonts (headings)
            major = font_scheme.find("a:majorFont", _DRAWINGML_NS)
            if major is not None:
                latin = major.find("a:latin", _DRAWINGML_NS)
                ea = major.find("a:ea", _DRAWINGML_NS)
                if latin is not None:
                    result["major_latin"] = latin.get("typeface")
                if ea is not None:
                    result["major_ea"] = ea.get("typeface")

            # Minor fonts (body)
            minor = font_scheme.find("a:minorFont", _DRAWINGML_NS)
            if minor is not None:
                latin = minor.find("a:latin", _DRAWINGML_NS)
                ea = minor.find("a:ea", _DRAWINGML_NS)
                if latin is not None:
                    result["minor_latin"] = latin.get("typeface")
                if ea is not None:
                    result["minor_ea"] = ea.get("typeface")
    # pylint: disable=broad-except
    except Exception as e:
        print(f"[WARN] Could not get theme fonts: {e}")

    return result


def get_shape_font(
    shape, theme_fonts: Optional[Dict[str, Optional[str]]] = None
) -> Optional[str]:
    """
    Get the font name from a shape's text frame.

    Checks font settings in the following order (highest priority first):
    1. First run's rPr (run properties) - ea then latin
    2. First paragraph's pPr > defRPr (default run properties) - ea then latin
    3. Text frame's lstStyle > lvl1pPr > defRPr - ea then latin
    4. Theme fonts fallback (minor_ea > major_ea > minor_latin > major_latin)

    Resolves theme font references (e.g., "+mj-ea", "+mn-ea") to actual font names.

    Per ISO/IEC 29500-1: when defRPr is omitted, the application uses theme defaults.

    Args:
        shape: The shape object containing the text frame.
        theme_fonts: Dictionary containing theme font names for resolving references.
    Returns:
        Font name string, or None if no font could be determined.
    """
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return None

    tf = shape.text_frame
    if not tf.paragraphs:
        return None

    def _get_font_from_rpr(rpr_elem) -> Optional[str]:
        """Extract font name from rPr element (ea first, then latin)."""
        if rpr_elem is None:
            return None
        # Try East Asian font first
        ea = rpr_elem.find("a:ea", _DRAWINGML_NS)
        if ea is not None:
            ea_typeface = ea.get("typeface")
            if ea_typeface:
                if ea_typeface.startswith("+"):
                    resolved = _resolve_theme_font_reference(ea_typeface, theme_fonts)
                    if resolved:
                        return resolved
                else:
                    return ea_typeface
        # Fall back to Latin font
        latin = rpr_elem.find("a:latin", _DRAWINGML_NS)
        if latin is not None:
            latin_typeface = latin.get("typeface")
            if latin_typeface:
                if latin_typeface.startswith("+"):
                    resolved = _resolve_theme_font_reference(
                        latin_typeface, theme_fonts
                    )
                    if resolved:
                        return resolved
                else:
                    return latin_typeface
        return None

    p = tf.paragraphs[0]

    # 1. Try first run's rPr (highest priority)
    if p.runs:
        run = p.runs[0]
        try:
            # pylint: disable=protected-access
            r_elem = run._r
            rpr = r_elem.find("a:rPr", _DRAWINGML_NS)
            font = _get_font_from_rpr(rpr)
            if font:
                return font
        except AttributeError as e:
            # python-pptx internal structure access failed (version compatibility issue)
            print(f"[DEBUG] Could not access run element: {e}")

        # Also check font.name via python-pptx API
        if run.font.name:
            if run.font.name.startswith("+"):
                resolved = _resolve_theme_font_reference(run.font.name, theme_fonts)
                if resolved:
                    return resolved
            else:
                return run.font.name

    # 2. Try paragraph's pPr > defRPr
    try:
        # pylint: disable=protected-access
        p_elem = p._p
        ppr = p_elem.find("a:pPr", _DRAWINGML_NS)
        if ppr is not None:
            def_rpr = ppr.find("a:defRPr", _DRAWINGML_NS)
            font = _get_font_from_rpr(def_rpr)
            if font:
                return font
    except AttributeError as e:
        # python-pptx internal structure access failed (version compatibility issue)
        print(f"[DEBUG] Could not access paragraph element: {e}")

    # 3. Try text frame's lstStyle > lvl1pPr > defRPr
    try:
        # pylint: disable=protected-access
        tx_body = tf._txBody
        lst_style = tx_body.find("a:lstStyle", _DRAWINGML_NS)
        if lst_style is not None:
            lvl1_ppr = lst_style.find("a:lvl1pPr", _DRAWINGML_NS)
            if lvl1_ppr is not None:
                def_rpr = lvl1_ppr.find("a:defRPr", _DRAWINGML_NS)
                font = _get_font_from_rpr(def_rpr)
                if font:
                    return font
    except AttributeError as e:
        # python-pptx internal structure access failed (version compatibility issue)
        print(f"[DEBUG] Could not access text body element: {e}")

    # 4. Fallback to theme fonts (minor_ea > major_ea > minor_latin > major_latin)
    # Per ISO/IEC 29500-1: when defRPr is omitted, application uses theme defaults
    if theme_fonts:
        return (
            theme_fonts.get("minor_ea")
            or theme_fonts.get("major_ea")
            or theme_fonts.get("minor_latin")
            or theme_fonts.get("major_latin")
        )

    return None


# Type definition for paragraph default settings
#
# Line spacing types per ISO/IEC 29500-1:
#   - LINE_SPACING_TYPE_RATIO ('ratio', spcPct): Percentage-based spacing relative to font size
#     Reference: ISO/IEC 29500-1 §21.1.2.2.11 (spcPct - Spacing Percent)
#     "This element specifies the amount of white space... in the form of a
#     percentage of the text size."
#     https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpercent
#
#   - LINE_SPACING_TYPE_FIXED ('fixed_pt', spcPts): Absolute spacing in points
#     Reference: ISO/IEC 29500-1 §21.1.2.2.12 (spcPts - Spacing Points)
#     "This element specifies the amount of white space... in the form of a
#     text point size. The size is specified using points where 100 is equal
#     to 1 point font and 1200 is equal to 12 point."
#     https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints
#
#   - LINE_SPACING_TYPE_DEFAULT ('default'): No explicit line spacing set (use application default)
#     Reference: ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing)
#     "If this element is omitted then the spacing between two lines of text
#     should be determined by the point size of the largest piece of text within a line."
#     https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing
LineSpacingType = Literal["ratio", "fixed_pt", "default"]


class ParagraphDefaults(TypedDict):
    """
    Type definition for paragraph default settings returned by get_placeholder_paragraph_defaults.

    References:
        - ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing):
          "This element specifies the vertical line spacing that is to be used within
          a paragraph. This can be specified in two different ways, percentage spacing
          and font point spacing."
          https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing
    """

    line_spacing: Optional[float]
    """
    Line spacing value. Interpretation depends on line_spacing_type:
        - 'ratio': multiplier (e.g., 1.0 for single, 1.5 for 1.5 lines)
        - 'fixed_pt': absolute line height in points
        - 'default': None (use application default)
    """

    line_spacing_type: LineSpacingType
    """
    The type of line spacing:
        - 'ratio': percentage-based (spcPct per ISO/IEC 29500-1 §21.1.2.2.11)
        - 'fixed_pt': fixed points (spcPts per ISO/IEC 29500-1 §21.1.2.2.12)
        - 'default': not explicitly set
    """

    space_before_pt: float
    """Space before paragraph in points (0.0 if not set)."""

    space_after_pt: float
    """Space after paragraph in points (0.0 if not set)."""


def get_placeholder_paragraph_defaults(shape: Any) -> ParagraphDefaults:
    """
    Get paragraph defaults from placeholder's inherited lstStyle.

    This function traverses the PowerPoint style inheritance chain:
    1. Shape pPr (explicit on paragraph)
    2. Shape lstStyle (rarely used)
    3. Layout placeholder lstStyle  <-- This is what we check
    4. Master placeholder lstStyle
    5. Theme defaults

    Line Spacing Interpretation (per ISO/IEC 29500-1):
        The lnSpc (Line Spacing) element can contain either:
        - spcPct (Spacing Percent): "the amount of white space... in the form of
          a percentage of the text size" (§21.1.2.2.11)
        - spcPts (Spacing Points): "the amount of white space... in the form of
          a text point size" (§21.1.2.2.12)

        For spcPts (fixed spacing), the value represents the total line height,
        not additional spacing. Per Microsoft Word documentation:
        "When you choose [Exactly], the line spacing is measured in points,
        just like font size."
        Reference: https://edu.gcfglobal.org/en/word/line-and-paragraph-spacing/1/

    Args:
        shape: A placeholder shape from a slide.

    Returns:
        ParagraphDefaults: A TypedDict containing:
            - line_spacing: Line spacing value.
              For 'ratio' type: multiplier (e.g., 1.0 for single spacing).
              For 'fixed_pt' type: absolute line height in points.
              None if using default (application determines from font size).
            - line_spacing_type: 'ratio' (spcPct), 'fixed_pt' (spcPts), or 'default'.
            - space_before_pt: Space before paragraph in points (0.0 if not set).
            - space_after_pt: Space after paragraph in points (0.0 if not set).
    """
    result: ParagraphDefaults = {
        "line_spacing": None,
        "line_spacing_type": LINE_SPACING_TYPE_DEFAULT,
        "space_before_pt": 0.0,
        "space_after_pt": 0.0,
    }

    # Check if shape has placeholder format
    if not hasattr(shape, "placeholder_format") or shape.placeholder_format is None:
        return result

    # Get the slide layout
    slide = shape.part.slide
    layout = slide.slide_layout

    # Find matching placeholder in layout by idx
    ph_idx = shape.placeholder_format.idx
    layout_shape = None
    for ph in layout.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            layout_shape = ph
            break

    if layout_shape is None:
        return result

    # Access txBody and lstStyle
    # Note: _element access is required for XML parsing, python-pptx doesn't expose this
    # Variable names reflect XML element names but use snake_case
    # pylint: disable=protected-access
    tx_body = layout_shape._element.find(qn("p:txBody"))
    if tx_body is None:
        return result

    lst_style = tx_body.find(qn("a:lstStyle"))
    if lst_style is None:
        return result

    # Get lvl1pPr (level 1 = bullet level 0)
    lvl1_ppr = lst_style.find(qn("a:lvl1pPr"))
    if lvl1_ppr is None:
        return result

    # Extract line spacing (lnSpc)
    # Per ISO/IEC 29500-1 §21.1.2.2.5:
    # "This element specifies the vertical line spacing... This can be specified
    # in two different ways, percentage spacing and font point spacing."
    ln_spc = lvl1_ppr.find(qn("a:lnSpc"))
    if ln_spc is not None:
        # Check for percentage-based spacing (spcPct)
        # Per ISO/IEC 29500-1 §21.1.2.2.11: value is in 1/100000 of a percent
        # Example: 100000 = 100% = single spacing
        spc_pct = ln_spc.find(qn("a:spcPct"))
        # Check for fixed point spacing (spcPts)
        # Per ISO/IEC 29500-1 §21.1.2.2.12: value is in 1/100 of a point
        # Example: 1400 = 14 points
        spc_pts = ln_spc.find(qn("a:spcPts"))
        if spc_pct is not None:
            # Convert from 1/100000 percent to ratio (100000 -> 1.0)
            result["line_spacing"] = int(spc_pct.get("val")) / 100000
            result["line_spacing_type"] = LINE_SPACING_TYPE_RATIO
        elif spc_pts is not None:
            # Convert from 1/100 points to points (1400 -> 14.0)
            # This value represents the total line height, not additional spacing
            result["line_spacing"] = int(spc_pts.get("val")) / 100
            result["line_spacing_type"] = LINE_SPACING_TYPE_FIXED

    # Extract space before (spcBef)
    # Note: Only spcPts (fixed points) is supported. spcPct (percentage) is not implemented
    # because it requires font size context which is not available at this stage.
    spc_bef = lvl1_ppr.find(qn("a:spcBef"))
    if spc_bef is not None:
        spc_pts = spc_bef.find(qn("a:spcPts"))
        if spc_pts is not None:
            result["space_before_pt"] = int(spc_pts.get("val")) / 100
        elif spc_bef.find(qn("a:spcPct")) is not None:
            print(
                "[WARN] spcBef with spcPct (percentage) is not supported; using default (0.0)"
            )

    # Extract space after (spcAft)
    # Note: Only spcPts (fixed points) is supported. spcPct (percentage) is not implemented
    # because it requires font size context which is not available at this stage.
    spc_aft = lvl1_ppr.find(qn("a:spcAft"))
    if spc_aft is not None:
        spc_pts = spc_aft.find(qn("a:spcPts"))
        if spc_pts is not None:
            result["space_after_pt"] = int(spc_pts.get("val")) / 100
        elif spc_aft.find(qn("a:spcPct")) is not None:
            print(
                "[WARN] spcAft with spcPct (percentage) is not supported; using default (0.0)"
            )

    return result
