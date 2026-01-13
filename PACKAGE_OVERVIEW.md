# StimuPop - Complete Package

## What's Included

This complete package contains everything needed to run a professional Excel to PowerPoint conversion tool with a web interface. StimuPop extracts embedded images from Excel files and generates formatted presentations.

### Core Application
- **app.py** - Main Streamlit application
- **src/** - Production-grade source modules
- **config.yaml** - External configuration file
- **requirements.txt** - Production dependencies

### Source Modules (`src/`)
- **config.py** - Configuration management with YAML + env overrides
- **exceptions.py** - Custom exception hierarchy
- **validators.py** - Input validation and text sanitization
- **image_handler.py** - Image loading from files and Excel
- **excel_handler.py** - Excel file processing and validation
- **pptx_generator.py** - PowerPoint generation engine
- **logging_config.py** - Structured logging with request IDs

### Testing (`tests/`)
- **72 unit tests** covering all core modules
- **conftest.py** - Shared test fixtures
- Run with: `pytest tests/ -v`

### Documentation
- **VERSION_HISTORY.md** - Release notes and changelog
- **MEMORY.md** - Architecture decisions and patterns

### Configuration
- **config.yaml** - All configurable settings
- **requirements-dev.txt** - Development dependencies

## Quick Start (3 Steps)

1. **Install Python** (if not already installed)
2. **Install dependencies**: `pip install -r requirements.txt`
3. **Run the app**: `streamlit run app.py`

## Features Implemented

### Core Functionality
- Excel file upload (.xlsx) with embedded images
- PowerPoint template upload (optional)
- Embedded image extraction from Excel
- Local file path image support
- Portrait and landscape orientation options
- Automatic image centering
- Automatic text centering
- One slide per Excel row

### User Interface
- Clean, professional web interface
- Real-time data preview
- Progress indicators during generation
- Download button for completed presentations
- Configuration sidebar
- Advanced settings panel
- Built-in instructions and help

### Customization Options
- Image column selection
- Text columns selection (comma-separated)
- Font size adjustment (10-32pt)
- Image width control (3-7 inches)
- Image position control
- Text position control
- Portrait/landscape orientation

### Error Handling
- Image load failure handling (skip and continue)
- Invalid Excel file handling
- Missing column detection
- User-friendly error messages
- Warning for problematic rows

### Image Sources
- Embedded images in Excel cells (extracted automatically)
- Local file paths (e.g., `C:\Images\photo.jpg`)
- Supports: JPG, JPEG, PNG, GIF, WEBP, BMP

## How It Works

1. **User uploads Excel file** with embedded images or file paths
2. **App extracts embedded images** from the Excel file
3. **App reads each row** and creates one slide per row
4. **Images are loaded** and centered on slides
5. **Text is formatted** and centered below images
6. **PowerPoint file is generated** and ready for download

## Default Settings

- **Slide Size**: 7.5" x 10" (portrait) or 10" x 7.5" (landscape)
- **Image Width**: 5.5 inches (centered)
- **Image Position**: 0.5 inches from top
- **Text Position**: 5 inches from top
- **Font Size**: 14pt
- **Text Alignment**: Centered
- **Default Columns**: B (images), C,D,E,F (text)

## Technical Details

### Built With
- **Streamlit** - Modern web framework for Python
- **python-pptx** - PowerPoint file generation
- **pandas** - Excel data processing
- **openpyxl** - Excel file handling and embedded image extraction
- **Pillow** - Image processing
- **PyYAML** - Configuration management

### System Requirements
- Python 3.8 or higher
- 2GB RAM minimum
- Modern web browser
- No internet connection required (images are local/embedded)

## File Structure

```
StimuPop/
├── app.py                      # Main Streamlit UI
├── config.yaml                 # External configuration
├── requirements.txt            # Production dependencies
├── requirements-dev.txt        # Dev/test dependencies
├── src/
│   ├── __init__.py             # Package exports
│   ├── config.py               # Configuration management
│   ├── exceptions.py           # Custom exceptions
│   ├── validators.py           # Input validation
│   ├── image_handler.py        # Image loading
│   ├── excel_handler.py        # Excel processing
│   ├── pptx_generator.py       # PPTX generation
│   └── logging_config.py       # Logging setup
├── tests/
│   ├── conftest.py             # Test fixtures
│   ├── test_validators.py
│   ├── test_image_handler.py
│   ├── test_excel_handler.py
│   └── test_pptx_generator.py
├── VERSION_HISTORY.md          # Changelog
├── MEMORY.md                   # Architecture docs
└── logs/                       # Log files (created at runtime)
```

## Ready to Use

This package is complete and production-ready:
1. Install dependencies
2. Run the app
3. Upload Excel file with embedded images
4. Generate presentations

No additional coding or setup required!

---

**Package Version**: 2.1.0
**Updated**: January 2026
**Python Version**: 3.8+
