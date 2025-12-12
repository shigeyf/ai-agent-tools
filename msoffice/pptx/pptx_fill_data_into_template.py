# pptx_fill_data_into_template.py
# -*- coding: utf-8 -*-

"""
Fill data into PowerPoint placeholders from JSON payload.
Supports text, image, and list types.

Expected JSON specification:
{
    "templatePptx": "template.pptx",
    "outputPptx": "output.pptx",
    "slideIndex": 0,
    "placeholders": {
        "ph1": {
            "name": "Title Placeholder",
            "type": "text",
            "isTitle": true,
            "maxFontSize": 36,
            "value": "This is the title"
        },
        "ph2": {
            "name": "Image Placeholder",
            "type": "image",
            "value": "path/to/image.jpg"
        },
        "ph3": {
            "name": "List Placeholder",
            "type": "list",
            "maxFontSize": 24,
            "value": [
                "Label1: Body text for item 1",
                "Label2: Body text for item 2"
            ]
        }
    }
}
"""

import os
from typing import Dict, List, Optional, Tuple, Any
from PIL import Image, UnidentifiedImageError
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Emu, Pt

from font_size_calculator import (
    calculate_fitting_font_size,
    clear_font_cache,
    initialize_font_system,
)
from pptx_get_shape_font import (
    get_theme_fonts,
    get_shape_font,
    get_placeholder_paragraph_defaults,
    LINE_SPACING_TYPE_FIXED,
)


MAX_TITLE_SIZE = 36
MAX_FONT_SIZE = 24
LABEL_SEPARATORS = ["：", ":"]

# PowerPoint's internal line height factor for single line spacing.
# This value is derived from PowerPoint's default behavior where the line height
# is approximately 1.2 times the font size (120% of font size).
# Reference: PowerPoint's line spacing "Single" corresponds to ~120% line height.
#
# This factor is ONLY used for ratio-based (percentage) line spacing (spcPct).
# For fixed line spacing (spcPts), the line_spacing value is used directly as
# the absolute line height, and this factor is ignored.
#
# Formula for ratio-based spacing:
#   line_height = font_size × line_height_factor × line_spacing_ratio
#   where line_spacing_ratio is the user-defined spacing (1.0 for single, 1.5 for 1.5 lines, etc.)
#
# References:
#   - ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing):
#     "This element specifies the vertical line spacing... This can be specified
#     in two different ways, percentage spacing and font point spacing."
#     https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing
#
#   - ISO/IEC 29500-1 §21.1.2.2.12 (spcPts - Spacing Points):
#     "This element specifies the amount of white space... in the form of a text point size."
#     https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints
POWERPOINT_LINE_HEIGHT_FACTOR = 1.2


# ---------------------------
# Sub functions
# ---------------------------


def _get_text_frame_dimensions(shape) -> Tuple[float, float]:
    """
    Get the effective text area dimensions from a shape in points.

    Calculates the available text area by subtracting margins from
    the shape's total dimensions.

    Args:
        shape: The shape object containing the text frame.
    Returns:
        Tuple[float, float]: A tuple of (width_pt, height_pt) representing
            the effective text area dimensions in points.
    """
    tf = shape.text_frame

    # Calculate effective dimensions (excluding margins)
    effective_width = shape.width - tf.margin_left - tf.margin_right
    effective_height = shape.height - tf.margin_top - tf.margin_bottom

    # Convert EMU to points
    width_pt = Emu(effective_width).pt
    height_pt = Emu(effective_height).pt

    print(
        f"[DEBUG] Shape '{shape.name}' "
        f"dimensions: width={shape.width} EMU ({width_pt:.1f} pt), "
        f"height={shape.height} EMU ({height_pt:.1f} pt)"
    )

    return width_pt, height_pt


def _get_pptx_shape_by_name(slide, name: str) -> Optional[Any]:
    """
    Get shape in slide by its name.
    Args:
        slide: PowerPoint slide object.
        name (str): Name of the shape to find.
    Returns:
        shape: The shape object if found; otherwise, None.
    """
    for s in slide.shapes:
        if s.name == name:
            return s
    return None


def _split_label_body(text: str) -> Tuple[str, str]:
    """
    Split text into label and body at the first colon (： or :).
    Args:
        text (str): The text to split.
    Returns:
        Tuple[str, str]: A tuple containing the label and body.
    """
    idx = text.find(LABEL_SEPARATORS[0])
    if idx < 0:
        idx = text.find(LABEL_SEPARATORS[1])
    if idx < 0:
        return text, ""
    label = text[:idx]
    body = text[idx + 1 :].lstrip()  # Remove colon and trim leading whitespace
    return label, body


def _fill_text(
    shape, text, is_title=False, max_font_size: Optional[int] = None
) -> None:
    """
    Fill text into given placeholder shape's text frame.

    Currently uses PowerPoint's built-in TEXT_TO_FIT_SHAPE auto-sizing.
    Parameters is_title and max_font_size are reserved for future implementation
    of custom font size optimization similar to _fill_list().

    Args:
        shape: The shape object to fill text into.
        text (str): The text to fill.
        is_title (bool): Whether the text is a title (affects max font size).
            Reserved for future implementation.
        max_font_size (Optional[int]): Maximum font size to use.
            Reserved for future implementation.
    Returns:
        None

    TODO: Implement custom font size calculation similar to _fill_list() to handle
          multi-byte characters properly. PowerPoint's fit_text() does not work
          correctly with Japanese and other multi-byte character sets.
    """
    # Parameters reserved for future font size optimization implementation
    # pylint: disable=unused-argument
    del is_title  # Reserved for future: differentiate title vs body text sizing
    del max_font_size  # Reserved for future: cap the calculated font size

    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        print(f"[WARN] Shape '{shape.name}' has no text_frame; skipped text injection.")
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True  # enable word wrap
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # shrink text to fit

    p = tf.paragraphs[0]
    p.text = text

    # Note: fit_text() does not work properly with multi-byte characters
    # No font settings are applied; template's default formatting is preserved
    # The lines below are kept commented out for reference
    # try:
    #    if max_font_size:
    #        max_size = max_font_size
    #    else:
    #        max_size = MAX_TITLE_SIZE if is_title else MAX_FONT_SIZE
    #    tf.fit_text(max_size=max_size)
    ## pylint: disable=broad-except
    # except Exception as e:
    #    print(f"[WARN] Shape '{shape.name}' Error at fit_text(): {e}")
    return


# pylint: disable=too-many-locals
def _fill_image(slide, shape, path) -> None:
    """
    Fill image into given placeholder shape,
    fitting it within the frame while maintaining aspect ratio.

    Args:
        slide: PowerPoint slide object.
        shape: The shape object to fill image into.
        path (str): Path to the image file.
    Returns:
        None
    """
    if not os.path.isfile(path):
        print(f"[WARN] Image file not found: {path}; skipped image injection.")
        return

    # Get image size (pixels)
    # Handle various image loading errors:
    # - UnidentifiedImageError: corrupted file or unsupported format
    # - PermissionError: no read permission
    # - OSError: general I/O errors
    try:
        with Image.open(path) as im:
            img_w_px, img_h_px = im.size
    except (UnidentifiedImageError, PermissionError, OSError) as e:
        print(f"[WARN] Could not open image file '{path}': {e}; skipped image injection.")
        return

    # Get frame size (EMU -> float)
    f_left = float(shape.left)
    f_top = float(shape.top)
    f_w = float(shape.width)
    f_h = float(shape.height)
    frame_ratio = f_w / f_h
    img_ratio = img_w_px / img_h_px

    if img_ratio >= frame_ratio:
        # Image is wider -> fit to frame width
        pic_w = f_w
        pic_h = f_w / img_ratio
    else:
        # Image is taller -> fit to frame height
        pic_h = f_h
        pic_w = f_h * img_ratio

    # Centering
    pic_left = f_left + (f_w - pic_w) / 2
    pic_top = f_top + (f_h - pic_h) / 2

    # Add image (EMU specified)
    pic = slide.shapes.add_picture(
        path,
        Emu(int(pic_left)),
        Emu(int(pic_top)),
        width=Emu(int(pic_w)),
        height=Emu(int(pic_h)),
    )
    pic.name = shape.name + "_Fit"

    # Remove original frame (comment out if not needed)
    try:
        shape.element.getparent().remove(shape.element)
        print(f"[INFO] Replaced frame by '{shape.name}_Fit' overlay for '{shape.name}'")
    # pylint: disable=broad-except
    except Exception as e:
        print(
            f"[INFO] Could not remove original frame '{shape.name}'. Kept overlay. Reason: {e}"
        )


def _fill_list(
    shape,
    items: List[str],
    max_font_size: Optional[int] = None,
    font_dir: Optional[str] = None,
    theme_fonts: Optional[Dict[str, Optional[str]]] = None,
) -> None:
    """
    Fill list of items into given placeholder shape's text frame.
    Each item is split into label and body at the first colon (： or :).

    Calculates appropriate font size for text and adjusts line spacing
    to ensure text fits within the shape bounds.

    Args:
        shape: The shape object to fill list into.
        items (List[str]): List of text items to fill.
        max_font_size (Optional[int]): Maximum font size to use.
        font_dir (Optional[str]): Directory containing font files.
        theme_fonts (Optional[Dict]): Theme fonts dictionary for fallback.
    Returns:
        None
    """
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        print(f"[WARN] Shape '{shape.name}' has no text_frame; skipped text injection.")
        return

    tf = shape.text_frame
    max_size = max_font_size if max_font_size else MAX_FONT_SIZE

    # Step 1: Determine font name from shape properties or theme fonts
    # get_shape_font handles the full priority chain including theme fallback
    font_name = get_shape_font(shape, theme_fonts)
    calculated_font_size: Optional[int] = None

    # Step 2: Calculate fitting font size
    if font_name and font_dir:
        print(f"[INFO] Using font: {font_name}")

        # Step 2.1: Get paragraph defaults from template
        para_defaults = get_placeholder_paragraph_defaults(shape)
        # Use direct key access for type safety with TypedDict
        line_spacing = para_defaults['line_spacing'] or 1.0
        line_spacing_type = para_defaults['line_spacing_type']
        space_before_pt = para_defaults['space_before_pt']
        space_after_pt = para_defaults['space_after_pt']

        # Determine if line spacing is fixed (absolute points) or ratio-based (percentage)
        # Per ISO/IEC 29500-1:
        #   - spcPct (Spacing Percent): line spacing as percentage of font size
        #   - spcPts (Spacing Points): line spacing as absolute point value
        # Reference:
        # https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing
        is_fixed_line_spacing = line_spacing_type == LINE_SPACING_TYPE_FIXED

        print(
            f"[INFO] Paragraph defaults: line_spacing={line_spacing}, "
            f"type={line_spacing_type}, is_fixed={is_fixed_line_spacing}, "
            f"space_before={space_before_pt}pt, space_after={space_after_pt}pt"
        )

        # Step 2.2: Get text frame dimensions and calculate appropriate font size
        width_pt, height_pt = _get_text_frame_dimensions(shape)
        try:
            calculated_font_size = calculate_fitting_font_size(
                width_pt=width_pt,
                height_pt=height_pt,
                items=items,
                max_size=max_size,
                font_name=font_name,
                font_dir=font_dir,
                line_spacing=line_spacing,
                space_before_pt=space_before_pt,
                space_after_pt=space_after_pt,
                line_height_factor=POWERPOINT_LINE_HEIGHT_FACTOR,
                is_fixed_line_spacing=is_fixed_line_spacing,
            )
            print(
                f"[INFO] Applied font size: {calculated_font_size} pt (max: {max_size} pt)"
            )
        except ValueError as e:
            print(f"[WARN] Font size calculation failed: {e}")
            print("[WARN] Falling back to auto-fit mode (TEXT_TO_FIT_SHAPE)")
            calculated_font_size = None
    else:
        if not font_name:
            print("[WARN] Could not determine font; skipping font size setting")
        elif not font_dir:
            print("[WARN] Font directory not specified; skipping font size setting")

    # Step 3: Fill with actual content and apply calculated font size (if available)
    tf.clear()
    tf.word_wrap = True  # enable word wrap
    tf.auto_size = (
        MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # shrink text to fit if calculated_font_size is None
        if calculated_font_size is None
        else MSO_AUTO_SIZE.NONE  # disable auto size, use manual font size
    )

    for i, line in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()

        # 2-run with label and body
        label, body = _split_label_body(line)
        # First run: label in bold
        run1 = p.add_run()
        run1.text = label.strip() if label else line.strip()
        run1.font.bold = True
        if calculated_font_size is not None:
            run1.font.size = Pt(calculated_font_size)
        # Second run: body in normal
        if body:
            run2 = p.add_run()
            run2.text = LABEL_SEPARATORS[0] + body
            run2.font.bold = False
            if calculated_font_size is not None:
                run2.font.size = Pt(calculated_font_size)

    # Note: fit_text() does not work properly with multi-byte characters
    # No font settings are applied; template's default formatting is preserved
    # The lines below are kept commented out for reference
    # try:
    #    print(max_font_size)
    #    max_size = max_font_size if max_font_size else MAX_FONT_SIZE
    #    tf.fit_text(max_size=max_size)
    ## pylint: disable=broad-except
    # except Exception as e:
    #    print(f"[WARN] Shape '{shape.name}' Error at fit_text(): {e}")
    return


# ---------------------------
# Main fill function
# ---------------------------
def pptx_fill_data_from_json(json_data: Dict) -> None:
    """
    Fill data into PowerPoint placeholders from JSON payload.
    Args:
        json_data (Dict): JSON payload with templatePptx, outputPptx, slideIndex, placeholders.
    Returns:
        None.
    Raises:
        FileNotFoundError: If the template PPTX file is not found.
        IndexError: If the slideIndex is out of range.
        ValueError: If the slideIndex is not an integer.
    """

    template_pptx = json_data.get("templatePptx", "")
    output_pptx = json_data.get("outputPptx", "output.pptx")
    try:
        slide_index = int(json_data.get("slideIndex", 0))
    except (ValueError, TypeError) as e:
        raise ValueError(f"[ERROR] slideIndex must be an integer {e}") from e

    phs = json_data.get("placeholders", {})
    font_dir = json_data.get("fontDir")  # Optional directory to search for fonts

    # Initialize font system early if fontDir is specified
    # This builds the font name mapping cache once, avoiding repeated scans
    if font_dir:
        initialize_font_system(font_dir)

    if not template_pptx or not os.path.isfile(template_pptx):
        raise FileNotFoundError(f"[ERROR] Template not found: {template_pptx}")

    prs = Presentation(template_pptx)
    print(f"[INFO] Loaded PowerPoint template file: {template_pptx}")

    # Get theme fonts for font resolution
    theme_fonts = get_theme_fonts(prs)

    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"[ERROR] slideIndex out of range: {slide_index} (0..{len(prs.slides)-1})"
        )
    slide = prs.slides[slide_index]
    print(f"[INFO] Loaded PowerPoint slide: Index[{slide_index}]")

    for _, ph_value in phs.items():
        name = ph_value.get("name")
        ph_type = ph_value.get("type")
        if ph_type is None or name is None:
            print("[WARN] Placeholder name/type missing; skipped.")
            continue
        is_title = ph_value.get("isTitle", False)
        max_font_size = ph_value.get("maxFontSize")
        if max_font_size is not None and not isinstance(max_font_size, int):
            print("[WARN] maxFontSize should be an integer; ignoring.")
            max_font_size = None

        print(f"[INFO] Filling placeholder '{name}' ...")
        shp = _get_pptx_shape_by_name(slide, name)
        if shp:
            ph_type = ph_type.lower()
            if ph_type == "text":
                _fill_text(shp, ph_value.get("value", ""), is_title, max_font_size)
            elif ph_type == "image":
                _fill_image(slide, shp, ph_value.get("value", ""))
            elif ph_type == "list":
                _fill_list(
                    shp, ph_value.get("value", []), max_font_size, font_dir, theme_fonts
                )
            else:
                print("[WARN] Unknown placeholder type; skipped.")
                continue

            print(f"[INFO] Filled data into '{name}'.")
        else:
            print(f"[WARN] Shape '{name}' not found.")

    prs.save(output_pptx)
    print(f"[INFO] Filled slide saved to : {output_pptx}")

    # Clear font cache to free memory after processing
    clear_font_cache()
    return
