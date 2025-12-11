# PowerPoint Template Data Filler

A Python tool to fill JSON data into PowerPoint slide placeholders. Supports text, image, and list content types with automatic text fitting and image aspect ratio preservation.

## Features

- **Text Placeholders**: Fill text with auto-sizing to fit the placeholder shape
- **Image Placeholders**: Insert images while maintaining aspect ratio and centering within the frame
- **List Placeholders**: Create bulleted lists with bold labels (text before colon) and normal body text
- **Flexible Configuration**: Control font sizes and title formatting via JSON

## Requirements

- Python 3.x
- Dependencies listed in [requirements.txt](requirements.txt):
  - `python-pptx==0.6.23`
  - `Pillow>=10.0.0`
  - `lxml>=4.9.3`

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

| Field          | Type    | Required | Description                                   |
| -------------- | ------- | -------- | --------------------------------------------- |
| `templatePptx` | string  | Yes      | Path to the PowerPoint template file          |
| `outputPptx`   | string  | No       | Output file path (default: `output.pptx`)     |
| `slideIndex`   | integer | No       | Zero-based slide index to fill (default: `0`) |
| `placeholders` | object  | Yes      | Dictionary of placeholder configurations      |

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

See [samples/input.json](samples/input.json) for a complete example:

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

## How It Works

1. **Load Template**: Opens the specified PowerPoint template file
2. **Select Slide**: Navigates to the specified slide by index
3. **Find Shapes**: Locates shapes by their exact name in the slide
4. **Fill Content**:
   - **Text**: Clears existing content, sets new text, and auto-fits font size
   - **Image**: Calculates optimal size maintaining aspect ratio, centers within frame, and replaces original shape
   - **List**: Creates paragraphs for each item, with bold labels separated by colons
5. **Save Output**: Saves the modified presentation to the output path

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

**Workarounds:**

1. **Use conservative `maxFontSize` values**: Set smaller maximum font sizes than you might expect to need
2. **Test with your specific fonts**: Results vary significantly depending on which fonts are used in your template
3. **Keep text concise**: Shorter text is less likely to overflow
4. **Manual verification**: Always open the generated PPTX in PowerPoint to verify text fits correctly
5. **Use standard fonts**: Fonts like Arial, Calibri, or Times New Roman tend to have more predictable behavior

For more details, see the [python-pptx documentation on fit_text()](https://python-pptx.readthedocs.io/en/latest/api/text.html#pptx.text.text.TextFrame.fit_text).

## License

See the repository's main LICENSE file.
