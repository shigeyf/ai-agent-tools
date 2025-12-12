# PowerPoint Template Data Filler

A Python tool to fill JSON data into PowerPoint slide placeholders. Supports text, image, and list content types with automatic text fitting and image aspect ratio preservation.

## Features

- **Text Placeholders**: Fill text with auto-sizing to fit the placeholder shape
- **Image Placeholders**: Insert images while maintaining aspect ratio and centering within the frame
- **List Placeholders**: Create bulleted lists with bold labels (text before colon) and normal body text
- **Flexible Configuration**: Control font sizes and title formatting via JSON
- **Precise Font Size Calculation**: Calculate optimal font sizes using actual font metrics with Pillow and fontTools
- **Multi-byte Character Support**: Proper text measurement for CJK (Chinese, Japanese, Korean) and other multi-byte character sets (requires consistent language per placeholder)
- **Theme Font Resolution**: Automatically resolve PowerPoint theme fonts (major/minor, Latin/East Asian)

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [JSON Specification](#json-specification)
  - [JSON Fields](#json-fields)
  - [Placeholder Properties](#placeholder-properties)
  - [Placeholder Types](#placeholder-types)
- [Sample Input](#sample-input)
  - [Sample with Font Directory](#sample-with-font-directory)
- [How It Works](#how-it-works)
- [How to Name Shapes in PowerPoint](#how-to-name-shapes-in-powerpoint)
- [Notes](#notes)
- [Known Limitations](#known-limitations)
  - [Text Overflow and fit_text() Issues](#text-overflow-and-fit_text-issues)
  - [Solution: Font Directory (fontDir)](#solution-font-directory-fontdir)
  - [Multi-Language Limitations](#multi-language-limitations)
- [License](#license)

## Requirements

- Python 3.8+
- Dependencies listed in [requirements.txt](requirements.txt):
  - `python-pptx>=1.0.2` - PowerPoint file manipulation with type annotations
  - `Pillow>=10.0.0` - Image processing
  - `lxml>=4.9.3` - XML processing
  - `lxml-stubs>=0.5.1` - Type stubs for lxml
  - `fonttools>=4.0.0` - Font file parsing
  - `XlsxWriter>=0.5.7` - Required by python-pptx 1.0+
  - `typing_extensions>=4.9.0` - Required by python-pptx 1.0+

## Installation

### Using Virtual Environment (Recommended)

```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Direct Installation

```bash
pip install -r requirements.txt
```

## Usage

### Command Line

```bash
python cli.py -i <path_to_input_json>
```

**Options:**

- `-i, --input`: Path to input JSON file (required)

### Example

```bash
python cli.py -i samples/input.json
```

## JSON Specification

The input JSON file should follow this structure:

```json
{
  "templatePptx": "path/to/template.pptx",
  "outputPptx": "path/to/output.pptx",
  "fontDir": "path/to/fonts",
  "slideIndex": 0,
  "placeholders": {
    "placeholder_key": {
      "name": "Shape Name in PowerPoint",
      "type": "text|image|list",
      "isTitle": false,
      "maxFontSize": 24,
      "value": "content"
    }
  }
}
```

### JSON Fields

| Field          | Type    | Required | Description                                                                |
| -------------- | ------- | -------- | -------------------------------------------------------------------------- |
| `templatePptx` | string  | Yes      | Path to the PowerPoint template file                                       |
| `outputPptx`   | string  | No       | Output file path (default: `output.pptx`)                                  |
| `fontDir`      | string  | No       | Directory containing font files (TTF/TTC/OTF) for precise text measurement |
| `slideIndex`   | integer | No       | Zero-based slide index to fill (default: `0`)                              |
| `placeholders` | object  | Yes      | Dictionary of placeholder configurations                                   |

### Placeholder Properties

| Property      | Type         | Required | Description                                                         |
| ------------- | ------------ | -------- | ------------------------------------------------------------------- |
| `name`        | string       | Yes      | Name of the shape in PowerPoint (must match exactly)                |
| `type`        | string       | Yes      | Content type: `text`, `image`, or `list`                            |
| `value`       | string/array | Yes      | Content to fill (string for text/image, array for list)             |
| `isTitle`     | boolean      | No       | If `true`, uses larger max font size (default: `false`)             |
| `maxFontSize` | integer      | No       | Maximum font size in points (default: 36 for titles, 24 for others) |

### Placeholder Types

#### Text

```json
{
  "name": "Title",
  "type": "text",
  "isTitle": true,
  "maxFontSize": 36,
  "value": "Your title text here"
}
```

#### Image

```json
{
  "name": "GraphImage",
  "type": "image",
  "value": "path/to/image.png"
}
```

#### List

List items can include labels (text before `:` or `：`) which will be rendered in bold:

```json
{
  "name": "Insights",
  "type": "list",
  "maxFontSize": 24,
  "value": ["Label1: Body text for item 1", "Label2: Body text for item 2", "Plain text without label"]
}
```

## Sample Input

See [samples/input.json](samples/input.json) for a basic example without font size calculation:

> **Note:** This example does not include `fontDir`, so **no font size calculation or adjustment is performed by this tool**. The `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` property flag is set, but text may overflow the shape boundaries. To manually trigger font size adjustment in PowerPoint, click on the auto-fit options icon (appears at bottom-left of the shape when text overflows) and select **"AutoFit Text to Placeholder"**. For reliable automated text fitting without manual intervention, use `fontDir` to enable precise font size calculation.

```json
{
  "templatePptx": "samples/template_single.pptx",
  "outputPptx": "output_slides_from_json.pptx",
  "slideIndex": 0,
  "placeholders": {
    "title": {
      "name": "Title",
      "type": "text",
      "isTitle": true,
      "maxFontSize": 36,
      "value": "NYC Taxi Monthly Ridership Trends and Insights"
    },
    "graphImage": {
      "name": "GraphImage",
      "type": "image",
      "value": "samples/images/graph.png"
    },
    "insights": {
      "name": "Insights",
      "type": "list",
      "maxFontSize": 24,
      "value": [
        "Long-term Decline: Downward trend since before the pandemic.",
        "Spring 2020 Crash: Sharp drop due to lockdowns and tourism collapse.",
        "Weakening Seasonality: Peak season fluctuations have diminished."
      ]
    },
    "note": {
      "name": "Note",
      "type": "text",
      "maxFontSize": 24,
      "value": "Note: Combining with external indicators such as FHV usage, subway/bus ridership, tourism statistics, and office attendance rates can help clarify causal relationships."
    }
  }
}
```

### Sample with Font Directory

For precise font size calculation with multi-byte characters (Japanese, Chinese, etc.), use `fontDir` to specify a directory containing font files. See [samples/input-with-fonts.json](samples/input-with-fonts.json):

```json
{
  "templatePptx": "samples/template_single.pptx",
  "outputPptx": "output_slides_from_json.pptx",
  "fontDir": "samples/fonts",
  "slideIndex": 0,
  "placeholders": {
    "title": {
      "name": "Title",
      "type": "text",
      "isTitle": true,
      "maxFontSize": 36,
      "value": "NYCタクシー 月次乗車数の推移からの示唆"
    },
    "graphImage": {
      "name": "GraphImage",
      "type": "image",
      "value": "samples/images/graph.png"
    },
    "insights": {
      "name": "Insights",
      "type": "list",
      "maxFontSize": 28,
      "value": [
        "長期減少トレンド：パンデミック以前から右肩下がり、アプリ配車 (Uber/Lyft等) への需要シフトが背景",
        "2020年春の急落と低位推移：外出制限・観光蒸発・在宅勤務定着で乗車数が急減後、低位で横ばい",
        "季節性の弱まり：繁忙期の波が縮小、公共交通・マイクロモビリティ・配車プラットフォームへの分散が影響"
      ]
    },
    "note": {
      "name": "Note",
      "type": "text",
      "maxFontSize": 24,
      "value": "参考：FHV利用件数、地下鉄/バス乗車数、観光統計、オフィス出社率など外部指標と併用すると因果関係の整理が可能。"
    }
  }
}
```

## How It Works

1. **Load Template**: Opens the specified PowerPoint template file
2. **Initialize Font System**: If `fontDir` is specified, builds a font name mapping cache for precise text measurement
3. **Select Slide**: Navigates to the specified slide by index
4. **Get Theme Fonts**: Extracts theme font information for resolving font references
5. **Find Shapes**: Locates shapes by their exact name in the slide
6. **Fill Content**:
   - **Text**: Clears existing content, sets new text, and auto-fits font size
   - **Image**: Calculates optimal size maintaining aspect ratio, centers within frame, and replaces original shape
   - **List**: Resolves font from shape/theme, calculates optimal font size using Pillow for precise measurement, creates paragraphs for each item with bold labels separated by colons
7. **Save Output**: Saves the modified presentation to the output path
8. **Cleanup**: Clears font caches to free memory

## How to Name Shapes in PowerPoint

To set or find shape names in PowerPoint:

1. **Open the Selection Pane**:
   - Go to **Home** → **Select** → **Selection Pane**, or
   - Press **Alt + F10** (Windows) / **Option + F10** (Mac)

2. **View Shape Names**: The Selection Pane displays all shapes on the current slide with their names

3. **Rename a Shape**:
   - Click on the shape name in the Selection Pane
   - Click again (or press F2) to edit the name
   - Type the new name and press Enter

4. **Tips**:
   - Use descriptive, unique names (e.g., "Title", "GraphImage", "Insights")
   - Avoid spaces and special characters for easier JSON handling
   - Shape names are case-sensitive

## Notes

- Shape names in PowerPoint must match the `name` property exactly
- Images are centered within the placeholder bounds while preserving aspect ratio
- Text auto-sizing uses `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` for automatic fitting
- The original image placeholder shape is replaced with the fitted image

## Known Limitations

### Text Overflow and `fit_text()` Issues

The `python-pptx` library has a significant limitation with text fitting: **it does not have access to actual font metrics**. This means:

- **`fit_text()` uses approximations**: The library estimates text dimensions based on generic font metrics, not the actual fonts installed on your system or embedded in the presentation.
- **Text overflow may occur**: The calculated font size might still result in text overflowing the placeholder bounds, especially with:
  - Non-standard or custom fonts
  - Fonts with unusual character widths (e.g., CJK characters, condensed/extended fonts)
  - Long words or text without natural break points
- **No font rendering engine**: Unlike PowerPoint itself, `python-pptx` cannot render text to measure exact dimensions.

### Solution: Font Directory (`fontDir`)

This tool now includes a **precise font size calculation** feature that addresses the above limitations:

- **Specify `fontDir`**: Point to a directory containing TTF/TTC/OTF font files
- **Pillow-based measurement**: Uses Pillow's FreeType integration for actual text width measurement
- **fontTools integration**: Extracts font names from font files to match PowerPoint's font references
- **Theme font resolution**: Automatically resolves theme font references (e.g., `+mj-ea`, `+mn-ea`) to actual font names

**Important Constraint:** All text items within a single placeholder should use the **same language** (or languages covered by the same font). The tool uses one font per placeholder for text measurement.

**Example (Japanese text):**

```json
{
  "templatePptx": "template.pptx",
  "fontDir": "samples/fonts",
  "placeholders": {
    "insights": {
      "name": "Insights",
      "type": "list",
      "value": ["日本語テキスト1", "日本語テキスト2", "日本語テキスト3"]
    }
  }
}
```

**Note:** Mixing different languages (e.g., Japanese, Chinese, Korean) in the same placeholder may result in incorrect text measurement if the selected font doesn't cover all characters.

**Font Directory Setup:**

1. Create a `fonts/` directory in your project
2. Copy the required font files (TTF/TTC/OTF) to the directory
3. Font names are automatically extracted from font metadata

**Supported font file types:**

- `.ttf` - TrueType Font
- `.ttc` - TrueType Collection (multiple fonts in one file)
- `.otf` - OpenType Font

**Workarounds (when `fontDir` is not available):**

1. **Use conservative `maxFontSize` values**: Set smaller maximum font sizes than you might expect to need
2. **Test with your specific fonts**: Results vary significantly depending on which fonts are used in your template
3. **Keep text concise**: Shorter text is less likely to overflow
4. **Manual verification**: Always open the generated PPTX in PowerPoint to verify text fits correctly
5. **Use standard fonts**: Fonts like Arial, Calibri, or Times New Roman tend to have more predictable behavior

For more details, see the [python-pptx documentation on fit_text()](https://python-pptx.readthedocs.io/en/latest/api/text.html#pptx.text.text.TextFrame.fit_text).

### Multi-Language Limitations

The current implementation has the following constraints for multi-language support:

| Feature                             | Status             | Notes                                           |
| ----------------------------------- | ------------------ | ----------------------------------------------- |
| Multiple paragraphs                 | ✅ Supported       | Each list item becomes a paragraph              |
| Single language per placeholder     | ✅ Works well      | All items use the same font for measurement     |
| Mixed CJK + Latin in same paragraph | ⚠️ Partial         | Uses primary font (typically East Asian)        |
| Different languages per paragraph   | ❌ Not optimal     | Font measurement uses single font for all items |
| Automatic language detection        | ❌ Not implemented | Font is determined from shape/theme settings    |

**Best Practice:** Use separate placeholders for content in different languages, or ensure all text in a placeholder is covered by the template's configured font.

## License

See the repository's main LICENSE file.
