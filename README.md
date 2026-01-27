# StimuPop

**Excel to PowerPoint Converter with Embedded Images**

[![Version](https://img.shields.io/badge/version-2.3.0-blue.svg)](VERSION_HISTORY.md)
[![Python](https://img.shields.io/badge/python-3.8+-green.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/streamlit-1.24+-red.svg)](https://streamlit.io)

StimuPop is a production-grade web application that converts Excel spreadsheet rows into professional PowerPoint presentations. Each row becomes a slide with embedded images and formatted text.

## Features

- **Embedded Image Extraction** - Automatically extracts images embedded in Excel cells
- **Uniform Image Sizing** - Multiple sizing modes ensure consistent image dimensions across all slides
- **Per-Column Text Formatting** - Customize font, size, color, bold/italic for each column
- **Template Support** - Use your own PowerPoint templates for branding
- **Portrait & Landscape** - Support for both slide orientations
- **Error Resilience** - Failed images don't stop generation; slides are created with available content
- **Portable Distribution** - Share with testers via a single ZIP file (no Python required)

## Quick Start

### For Users (Portable Version)

1. Download `StimuPop_Portable.zip`
2. Extract to any folder
3. Double-click `StimuPop.bat`
4. Browser opens automatically at `http://localhost:8501`

### For Developers

```bash
# Clone the repository
git clone <repository-url>
cd stimupop

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

## Usage

### 1. Prepare Your Excel File

Structure your Excel file with:
- One column containing images (embedded or file paths)
- One or more columns containing text content

| Column A | Column B (Image) | Column C (Title) | Column D (Description) |
|----------|------------------|------------------|------------------------|
| ID-001   | [Embedded Image] | Product One      | Description here...    |
| ID-002   | [Embedded Image] | Product Two      | Description here...    |

### 2. Configure Settings

**Basic Settings:**
- **Image Column**: Letter or name of column containing images (e.g., `B`)
- **Text Columns**: Comma-separated columns for text (e.g., `C,D,E`)
- **Font Size**: Default text size (10-32pt)

**Advanced Settings (Image Sizing):**

| Mode | Description |
|------|-------------|
| **Fit to Box** | Scale to fit within max width/height, preserve aspect ratio (recommended) |
| **Fit Width** | Fixed width, height adjusts automatically |
| **Fit Height** | Fixed height, width adjusts automatically |
| **Stretch** | Exact dimensions (may distort) |

### 3. Generate

Click "Generate Presentation" and download your `.pptx` file.

## Project Structure

```
stimupop/
├── app.py                    # Main Streamlit application
├── config.yaml               # Application configuration
├── requirements.txt          # Python dependencies
├── src/
│   ├── __init__.py          # Package exports
│   ├── config.py            # Configuration management
│   ├── exceptions.py        # Custom exceptions
│   ├── validators.py        # Input validation
│   ├── image_handler.py     # Image loading and caching
│   ├── excel_handler.py     # Excel file processing
│   ├── pptx_generator.py    # PowerPoint generation
│   └── logging_config.py    # Logging setup
├── tests/                    # Test suite
├── build_portable.bat        # Build portable distribution
├── create_user_guide.py      # Generate DOCX user guide
└── docs/
    └── StimuPop_User_Guide.docx
```

## Configuration

### config.yaml

```yaml
app:
  name: "StimuPop"
  version: "2.3.0"
  max_upload_size_mb: 200

presentation:
  default_orientation: "portrait"
  default_img_width: 5.5
  default_img_height: 4.0
  default_img_size_mode: "fit_box"
  default_font_size: 14

images:
  max_size_mb: 10
  allowed_formats: [".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"]
```

### Environment Variables

Override any config setting:
```bash
APP_IMAGES_MAX_SIZE_MB=20
APP_PRESENTATION_DEFAULT_FONT_SIZE=16
APP_LOGGING_LEVEL=DEBUG
```

## Building Portable Distribution

Create a self-contained distribution for testers:

```bash
# Run the build script
build_portable.bat

# Output: StimuPop_Portable.zip (~126MB)
```

The portable version includes:
- Embedded Python 3.11.9
- All dependencies pre-installed
- User guide (DOCX)
- Launch script

## API Reference

### SlideConfig

```python
from src import SlideConfig, IMG_SIZE_FIT_BOX

config = SlideConfig(
    img_column="B",
    text_columns=["C", "D", "E"],
    img_width=5.5,
    img_height=4.0,
    img_size_mode=IMG_SIZE_FIT_BOX,
    img_top=0.5,
    text_top=5.0,
    font_size=14,
    orientation="portrait"
)
```

### Image Sizing Modes

```python
from src import (
    IMG_SIZE_FIT_BOX,    # Fit within bounds, preserve ratio
    IMG_SIZE_FIT_WIDTH,  # Fixed width, auto height
    IMG_SIZE_FIT_HEIGHT, # Fixed height, auto width
    IMG_SIZE_STRETCH     # Exact size, may distort
)
```

### PPTXGenerator

```python
from src import PPTXGenerator, SlideConfig

generator = PPTXGenerator(config)
result = generator.generate(
    slide_data,
    embedded_images=embedded_dict,
    template_file=template_bytes,
    progress_callback=my_callback
)

if result.success:
    result.presentation.save("output.pptx")
```

## Testing

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src

# Run specific test file
pytest tests/test_pptx_generator.py -v
```

## Security Considerations

- **SSRF Prevention**: Private IP ranges are blocked for URL-based images
- **Input Validation**: All user inputs are sanitized
- **File Size Limits**: Configurable limits prevent resource exhaustion
- **No Secrets in Code**: All secrets use environment variables

## Troubleshooting

| Issue | Solution |
|-------|----------|
| App won't start | Ensure all files extracted; try Run as Administrator |
| Browser doesn't open | Navigate to `http://localhost:8501` manually |
| Images not appearing | Check images are embedded (not floating) in Excel |
| Different image sizes | Set Size Mode to "Fit to Box" in Advanced Settings |
| Slow generation | Reduce image sizes in Excel before upload |

## Version History

See [VERSION_HISTORY.md](VERSION_HISTORY.md) for detailed changelog.

### v2.3.0 (Current)
- Added uniform image sizing with 4 modes
- Added Max Height control
- Created portable distribution system
- Added comprehensive User Guide (DOCX)

### v2.2.0
- Per-column text formatting
- Enhanced error handling

### v2.0.0
- Modular architecture refactor
- SSRF protection
- Concurrent image downloads

## License

Proprietary - Internal Use Only

## Support

For issues or feature requests, contact the development team.

---

Built with [Streamlit](https://streamlit.io) | Powered by [python-pptx](https://python-pptx.readthedocs.io)
