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
from PIL import Image
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Emu


MAX_TITLE_SIZE = 36
MAX_FONT_SIZE = 24
LABEL_SEPARATORS = ["：", ":"]

# ---------------------------
# Sub functions
# ---------------------------


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


def _fill_text(shape, text, is_title=False, max_font_size: Optional[int] = None) -> None:
    """
    Fill text into given placeholder shape's text frame.
    Args:
        shape: The shape object to fill text into.
        text (str): The text to fill.
        is_title (bool): Whether the text is a title (affects max font size).
        max_font_size (Optional[int]): Maximum font size to use.
    Returns:
        None
    """
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        print(f"[WARN] Shape '{shape.name}' has no text_frame; skipped text injection.")
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True  # enable word wrap
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # shrink text to fit

    p = tf.paragraphs[0]
    p.text = text
    try:
        if max_font_size:
            max_size = max_font_size
        else:
            max_size = MAX_TITLE_SIZE if is_title else MAX_FONT_SIZE
        tf.fit_text(max_size=max_size)
    # pylint: disable=broad-except
    except Exception as e:
        print(f"[WARN] Shape '{shape.name}' Error at fit_text(): {e}")
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
        raise FileNotFoundError(f"[ERROR] Image file not found: {path}")

    # Get image size (pixels)
    with Image.open(path) as im:
        img_w_px, img_h_px = im.size

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


def _fill_list(shape, items: List[str], max_font_size: Optional[int] = None) -> None:
    """
    Fill list of items into given placeholder shape's text frame.
    Each item is split into label and body at the first colon (： or :).
    Args:
        shape: The shape object to fill list into.
        items (List[str]): List of text items to fill.
        max_font_size (Optional[int]): Maximum font size to use.
    Returns:
        None
    """
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        print(f"[WARN] Shape '{shape.name}' has no text_frame; skipped text injection.")
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True  # enable word wrap
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # shrink text to fit

    for i, line in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        # 2-run with label and body
        label, body = _split_label_body(line)
        # First run: label in bold
        run1 = p.add_run()
        run1.text = label.strip() if label else line.strip()
        run1.font.bold = True
        # Second run: body in normal
        if body:
            run2 = p.add_run()
            run2.text = LABEL_SEPARATORS[0] + body
            run2.font.bold = False
    try:
        max_size = max_font_size if max_font_size else MAX_FONT_SIZE
        tf.fit_text(max_size=max_size)
    # pylint: disable=broad-except
    except Exception as e:
        print(f"[WARN] Shape '{shape.name}' Error at fit_text(): {e}")
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

    if not template_pptx or not os.path.isfile(template_pptx):
        raise FileNotFoundError(f"[ERROR] Template not found: {template_pptx}")

    prs = Presentation(template_pptx)
    print(f"[INFO] Loaded PowerPoint template file: {template_pptx}")
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
                _fill_list(shp, ph_value.get("value", []), max_font_size)
            else:
                print("[WARN] Unknown placeholder type; skipped.")
                continue

            print(f"[INFO] Filled data into '{name}'.")
        else:
            print(f"[WARN] Shape '{name}' not found.")

    prs.save(output_pptx)
    print(f"[INFO] Filled slide saved to : {output_pptx}")
    return
