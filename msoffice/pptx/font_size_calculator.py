# font_size_calculator.py
# -*- coding: utf-8 -*-

"""
Font size calculation utilities for fitting text into bounded areas.

This module provides functions for calculating the optimal font size
that allows text to fit within specified dimensions. Text measurement
is performed using Pillow for precise width calculation.
"""

import os
from typing import Dict, List, Optional, Tuple

from PIL import ImageFont
from fontTools.ttLib import TTFont, TTCollection, TTLibError


# ---------------------------
# Constants
# ---------------------------

# Minimum font size in points
MIN_FONT_SIZE = 6

# DPI for internal point-to-pixel conversion in text measurement.
# This value is used for consistent text measurement calculations.
# The actual value doesn't affect results as long as it's used consistently
# throughout the calculations (both for text measurement and frame dimensions).
# Using 96 as it's a common standard, but any consistent value would work.
# Note: Pillow's truetype() takes size in pixels, PowerPoint uses points.
# Formula: pixels = points * DPI / 72
DEFAULT_DPI = 96

# Cached font object for performance
_font_cache: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}

# Cached dynamic font name to file mapping (populated by build_font_name_mapping)
_dynamic_font_name_to_file: Dict[str, List[str]] = {}


# ---------------------------
# Dynamic Font Name Mapping Functions
# ---------------------------


def _extract_names_from_font(font) -> Tuple[str, str, str]:
    """
    Extract name strings from a TTFont object.

    Name IDs:
        1 = Font Family (e.g., "Yu Gothic")
        2 = Font Subfamily (e.g., "Medium", "Bold")
        4 = Full Font Name (e.g., "Yu Gothic Medium")

    Args:
        font: TTFont object.

    Returns:
        Tuple of (full_name, family_name, subfamily_name)
    """
    name_table = font["name"]
    full_name = name_table.getDebugName(4) or ""
    family_name = name_table.getDebugName(1) or ""
    subfamily_name = name_table.getDebugName(2) or ""
    return (full_name, family_name, subfamily_name)


def _get_font_names_from_file(font_path: str) -> List[Tuple[str, str, str]]:
    """
    Extract font names from a TTF/TTC/OTF file.

    Args:
        font_path: Path to the font file.

    Returns:
        List of tuples: (full_name, family_name, subfamily_name)
    """
    results = []
    ttc = None
    font = None
    try:
        if font_path.lower().endswith(".ttc"):
            ttc = TTCollection(font_path)
            for ttc_font in ttc.fonts:
                names = _extract_names_from_font(ttc_font)
                if names[0] or names[1]:  # has full_name or family_name
                    results.append(names)
        else:
            font = TTFont(font_path)
            names = _extract_names_from_font(font)
            if names[0] or names[1]:
                results.append(names)
    except FileNotFoundError:
        print(f"[WARN] Font file not found: '{font_path}'")
    except PermissionError:
        print(f"[WARN] Permission denied reading font file: '{font_path}'")
    except TTLibError as e:
        print(f"[WARN] Invalid or corrupted font file '{font_path}': {e}")
    except KeyError as e:
        print(f"[WARN] Font file '{font_path}' missing required table: {e}")
    finally:
        # Explicitly close fontTools objects to release file handles
        if ttc is not None:
            ttc.close()
        if font is not None:
            font.close()

    return results


def initialize_font_system(font_dir: str) -> None:
    """
    Initialize the font system by building the font name mapping cache.

    This function should be called once at the beginning of processing,
    typically when fontDir is specified in the JSON payload. It builds
    the font name to file mapping and caches it for subsequent lookups.

    If the cache is already populated, this function does nothing.

    Args:
        font_dir: Directory containing font files.

    Example:
        >>> initialize_font_system("/path/to/fonts")
        [INFO] Built dynamic font mapping: 38 entries
        >>> initialize_font_system("/path/to/fonts")  # Already initialized
        [INFO] Font system already initialized (38 entries)
    """
    global _dynamic_font_name_to_file  # pylint: disable=global-statement

    # Skip if already initialized
    if _dynamic_font_name_to_file:
        print(
            f"[INFO] Font system already initialized ({len(_dynamic_font_name_to_file)} entries)"
        )
        return

    # Build and cache the mapping
    print(f"[INFO] Initializing font system from directory: {font_dir}")
    mapping = _build_font_name_mapping(font_dir)
    _dynamic_font_name_to_file = mapping


def _build_font_name_mapping(font_dir: str) -> Dict[str, List[str]]:
    """
    Build a font name to file mapping by scanning the font directory.

    This is a pure function that scans all TTF/TTC/OTF files in the directory
    and extracts font names using fontTools. Creates a mapping from lowercase
    font names (both family name and full name) to the corresponding font filenames.

    Note: This function does not modify any global state. Use initialize_font_system()
    to populate the module-level cache.

    Args:
        font_dir: Directory containing font files.

    Returns:
        Dictionary mapping lowercase font name -> list of filenames.
        Returns empty dict if directory is invalid.
    """
    if not font_dir or not os.path.isdir(font_dir):
        return {}

    font_extensions = (".ttf", ".ttc", ".otf")
    family_to_files: Dict[str, List[str]] = {}

    try:
        filenames = os.listdir(font_dir)
    except PermissionError:
        print(f"[WARN] Permission denied accessing font directory: '{font_dir}'")
        return {}
    except OSError as e:
        print(f"[WARN] Cannot access font directory '{font_dir}': {e}")
        return {}

    try:
        for filename in filenames:
            if filename.lower().endswith(font_extensions):
                font_path = os.path.join(font_dir, filename)
                font_names = _get_font_names_from_file(font_path)

                for full_name, family_name, _ in font_names:
                    # Add by family name (lowercase)
                    if family_name:
                        key = family_name.lower()
                        if key not in family_to_files:
                            family_to_files[key] = []
                        if filename not in family_to_files[key]:
                            family_to_files[key].append(filename)

                    # Add by full name (lowercase)
                    if full_name:
                        key_full = full_name.lower()
                        if key_full not in family_to_files:
                            family_to_files[key_full] = []
                        if filename not in family_to_files[key_full]:
                            family_to_files[key_full].append(filename)

        print(
            f"[INFO] Built font name and file mapping: {len(family_to_files)} entries"
        )

    except Exception as e:  # pylint: disable=broad-except
        # Catch any unexpected errors during font name extraction loop
        print(
            f"[WARN] Unexpected error while scanning font directory '{font_dir}': {e}"
        )

    return family_to_files


def get_font_name_mapping(font_dir: Optional[str] = None) -> Dict[str, List[str]]:
    """
    Get the font name to file mapping from cache.

    Returns the cached font name mapping. If no cache exists and font_dir is
    provided, initializes the font system first.

    Note: For explicit initialization, call initialize_font_system() directly.

    Args:
        font_dir: If provided and cache is empty, initializes the font system.

    Returns:
        Dictionary mapping font name -> list of filenames.
        Returns empty dict if not initialized.
    """
    # If cache is empty and font_dir is provided, initialize
    if not _dynamic_font_name_to_file and font_dir:
        initialize_font_system(font_dir)

    return _dynamic_font_name_to_file


# ---------------------------
# Font Loading Functions
# ---------------------------


def get_font(font_path: str, font_size_pt: int) -> Optional[ImageFont.FreeTypeFont]:
    """
    Get a cached font object for the given path and size.

    Args:
        font_path: Path to the font file (TTF/TTC).
        font_size_pt: Font size in points.

    Returns:
        ImageFont object or None if font cannot be loaded.
    """
    # Convert font size from points to pixels for Pillow
    # Pillow uses pixels, PowerPoint uses points
    font_size_px = int(font_size_pt * DEFAULT_DPI / 72)

    cache_key = (font_path, font_size_px)
    if cache_key in _font_cache:
        return _font_cache[cache_key]

    try:
        font = ImageFont.truetype(font_path, size=font_size_px)
        _font_cache[cache_key] = font
        return font
    except (IOError, OSError) as e:
        print(f"[WARN] Could not load font '{font_path}': {e}")
        return None


def clear_font_cache() -> None:
    """
    Clear all font caches to free memory.

    This function clears:
    - _font_cache: Pillow font objects (for text measurement)
    - _dynamic_font_name_to_file: Font name to filename mapping

    Call this function after processing is complete to release memory.

    Note: No 'global' statement needed here because .clear() modifies
    the dict in-place without reassigning the variable.
    """
    font_cache_count = len(_font_cache)
    mapping_count = len(_dynamic_font_name_to_file)
    _font_cache.clear()
    _dynamic_font_name_to_file.clear()
    print(
        f"[INFO] Cleared font cache: {font_cache_count} font objects,"
        f"{mapping_count} mapping entries"
    )


# ---------------------------
# Text Measurement Functions
# ---------------------------


def measure_text_width(text: str, font_path: str, font_size_pt: int) -> Optional[float]:
    """
    Measure the actual pixel width of text using the specified font.

    Args:
        text: The text to measure.
        font_path: Path to the font file (TTF/TTC).
        font_size_pt: Font size in points (will be converted to pixels internally).

    Returns:
        Width in pixels, or None if measurement failed.
    """
    font = get_font(font_path, font_size_pt)
    if font is None:
        return None

    bbox = font.getbbox(text)
    return bbox[2] - bbox[0]


def get_font_line_height(font_path: str, font_size_pt: int) -> Optional[float]:
    """
    Get the line height (ascent + descent) for the specified font.

    This function is currently not used in production code but is retained
    for potential future use (e.g., more precise line height calculations
    based on actual font metrics rather than PowerPoint's line_height_factor).

    Args:
        font_path: Path to the font file (TTF/TTC).
        font_size_pt: Font size in points.

    Returns:
        Line height in pixels, or None if measurement failed.
    """
    font = get_font(font_path, font_size_pt)
    if font is None:
        return None

    ascent, descent = font.getmetrics()
    return ascent + descent


# ---------------------------
# Font File Search Functions
# ---------------------------


def find_font_file(font_name: str, font_dir: str) -> Optional[str]:
    """
    Find a font file in the specified directory that matches the font name.

    Scans the font directory using fontTools to build a mapping from font
    names (extracted from font metadata) to font files.

    Args:
        font_name: Font name from PowerPoint (e.g., "Meiryo", "Yu Gothic").
        font_dir: Directory containing font files.

    Returns:
        Full path to the matching font file, or None if not found.
    """
    if not font_name or not font_dir or not os.path.isdir(font_dir):
        return None

    font_name_lower = font_name.lower()

    # Get dynamic mapping from font directory
    font_mapping = get_font_name_mapping(font_dir)

    # Check if font name is in our mapping
    if font_name_lower in font_mapping:
        possible_files = font_mapping[font_name_lower]
        for filename in possible_files:
            font_path = os.path.join(font_dir, filename)
            if os.path.isfile(font_path):
                return font_path

    # Fallback: Try to find a file that contains the font name
    # This is a best-effort heuristic when exact font name matching fails
    try:
        for filename in os.listdir(font_dir):
            if filename.lower().endswith((".ttf", ".ttc", ".otf")):
                # Check if font name is similar to filename
                name_parts = font_name_lower.replace(" ", "").replace("-", "")
                file_parts = (
                    filename.lower()
                    .replace(" ", "")
                    .replace("-", "")
                    .replace(".ttc", "")
                    .replace(".ttf", "")
                    .replace(".otf", "")
                )
                if name_parts in file_parts or file_parts in name_parts:
                    print(
                        f"[DEBUG] Font '{font_name}' matched by filename heuristic: {filename}"
                    )
                    return os.path.join(font_dir, filename)
    except OSError as e:
        # Directory access errors during fallback search (non-critical)
        print(f"[DEBUG] Fallback font search failed for '{font_name}': {e}")

    return None


# ---------------------------
# Font Size Calculation Functions
# ---------------------------


def calculate_fitting_font_size(
    width_pt: float,
    height_pt: float,
    items: List[str],
    max_size: int,
    font_name: str,
    font_dir: str,
    line_spacing: float,
    space_before_pt: float,
    space_after_pt: float,
    line_height_factor: float,
    is_fixed_line_spacing: bool = False,
) -> int:
    """
    Calculate the largest font size that allows all text items to fit within bounds.

    This function calculates the font size needed for text to fit within the
    specified width and height. It uses Pillow to measure actual text width
    for precise calculation.

    Line Height Calculation (per ISO/IEC 29500-1 and Microsoft documentation):

    The line_spacing parameter is interpreted differently based on is_fixed_line_spacing:

    1. Ratio-based spacing (is_fixed_line_spacing=False, default):
        line_height = font_size × line_height_factor × line_spacing
        Where line_spacing is a ratio (e.g., 1.0 for single, 1.5 for 1.5 lines)

    2. Fixed spacing (is_fixed_line_spacing=True):
        line_height = line_spacing (used directly as total line height in points)

    References:
        - ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing):
          "This element specifies the vertical line spacing... This can be specified
          in two different ways, percentage spacing and font point spacing."
          https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing

        - ISO/IEC 29500-1 §21.1.2.2.12 (spcPts - Spacing Points):
          "This element specifies the amount of white space... in the form of a
          text point size. The size is specified using points where 100 is equal
          to 1 point font and 1200 is equal to 12 point."
          https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints

        - Microsoft Word "Exactly" option behavior:
          "When you choose this option, the line spacing is measured in points,
          just like font size. For example, if you're using 12-point text,
          you could use 15-point spacing."
          https://edu.gcfglobal.org/en/word/line-and-paragraph-spacing/1/

    Height Calculation:
        paragraph_spacing = (space_before_pt + space_after_pt) × (num_paragraphs - 1)
        total_height = (total_lines × line_height) + paragraph_spacing

    Args:
        width_pt: Available width in points.
        height_pt: Available height in points.
        items: List of text items (paragraphs) to fit.
        max_size: Maximum font size in points.
        font_name: Font name to use for measurement.
        font_dir: Directory containing font files.
        line_spacing: Line spacing value from paragraph settings.
            - If is_fixed_line_spacing=False: ratio (e.g., 1.0 for single spacing)
            - If is_fixed_line_spacing=True: fixed line height in points
        space_before_pt: Space before each paragraph in points.
        space_after_pt: Space after each paragraph in points.
        line_height_factor: PowerPoint's internal line height factor (typically 1.2).
            Only used when is_fixed_line_spacing=False.
        is_fixed_line_spacing: If True, line_spacing is treated as absolute line
            height in points (ISO/IEC 29500-1 spcPts). If False (default),
            line_spacing is treated as a ratio multiplier (ISO/IEC 29500-1 spcPct).

    Returns:
        Calculated font size in points that fits all text.

    Raises:
        ValueError: If font file cannot be found for the given font_name.
    """

    # Resolve font path from font name and directory
    font_path = find_font_file(font_name, font_dir)
    if not font_path:
        raise ValueError(
            f"Font file not found for '{font_name}' in directory '{font_dir}'"
        )
    print(f"[INFO] Resolved font for precise measurement: {font_name} -> {font_path}")

    # Convert points to pixels for Pillow comparison
    # Pillow measures text in pixels, so we need consistent units
    pt_to_px = DEFAULT_DPI / 72
    width_px = width_pt * pt_to_px

    # Debug output: show text frame dimensions
    print(
        f"[DEBUG] Text frame dimensions: "
        f"width={width_pt:.1f}pt, height={height_pt:.1f}pt"
    )

    num_paragraphs = len(items)
    # Linear search for the largest font size that fits
    for font_size in range(max_size, MIN_FONT_SIZE - 1, -1):
        total_lines = 0
        lines_per_item = []  # Track lines per item for debug output

        # Use Pillow for precise text width measurement
        for item in items:
            text_width_px = measure_text_width(item, font_path, font_size)
            if text_width_px is None:
                # This should not happen if font_path is valid
                raise ValueError(f"Failed to measure text width for font '{font_path}'")
            # Calculate lines needed based on actual text width (in pixels)
            lines = max(1, int((text_width_px + width_px - 1) // width_px))
            total_lines += lines
            lines_per_item.append(lines)

        # Calculate line height based on spacing type
        # Reference: ISO/IEC 29500-1 §21.1.2.2.5 (lnSpc - Line Spacing)
        # "This can be specified in two different ways, percentage spacing and font point spacing."
        if is_fixed_line_spacing:
            # Fixed line spacing (spcPts): line_spacing is the absolute line height in points
            # Per ISO/IEC 29500-1 §21.1.2.2.12:
            #   spcPts specifies spacing "in the form of a text point size"
            # Per Microsoft Word UI: "Exactly" option means "line spacing is measured in points"
            line_height = line_spacing
        else:
            # Ratio-based line spacing (spcPct): calculate from font size
            # PowerPoint line height = font_size × line_height_factor × line_spacing
            # line_height_factor: PowerPoint's internal factor (typically 1.2 = 120%)
            # line_spacing: User-defined ratio from paragraph settings (e.g., 1.0 for single)
            line_height = font_size * line_height_factor * line_spacing

        # Calculate total height needed:
        # - Total text height = total_lines × line_height
        # - Paragraph spacing = (space_before + space_after) × (num_paragraphs - 1)
        total_text_height = total_lines * line_height
        total_para_spacing = (space_before_pt + space_after_pt) * (num_paragraphs - 1)
        total_height_needed = total_text_height + total_para_spacing

        # Debug output: show calculation details for each font size
        fit_status = "FIT" if total_height_needed <= height_pt else "OVERFLOW"
        print(
            f"[DEBUG] font_size={font_size}pt: "
            f"line_height={line_height:.1f}pt, "
            f"lines_per_item={lines_per_item}, total_lines={total_lines}, "
        )
        print(
            f"[DEBUG] text_height={total_text_height:.1f}pt, "
            f"para_spacing={total_para_spacing:.1f}pt, "
            f"total_height={total_height_needed:.1f}pt vs available={height_pt:.1f}pt "
            f"({fit_status})"
        )

        if total_height_needed <= height_pt:
            return font_size

    return MIN_FONT_SIZE
