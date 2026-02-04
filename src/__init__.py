"""
StimuPop - Source Package v8.0

Production-grade modules for Excel to PPTX conversion with embedded images.

New in v8.0:
- Multi-element support: multiple images and text boxes per slide (Template Mode)
- ImageElement and TextGroup dataclasses for multi-element configuration
- Dynamic shape matching in templates by placeholder name
- Backward compatible with all single-element workflows

New in v7.1:
- Fixed: Text overflow option now works in Template Mode
- Fixed: Intermittent Excel upload error with retry logic
- Browser warning workaround documented

New in v7.0:
- Dynamic Template Mode column mapping (fixes 5-column bug)
- No longer hardcoded to specific column letters

New in v6.2:
- Pictures Only mode (no text required)
- Text overflow handling option

New in v6.0:
- Configurable image alignment (top/center/bottom, left/center/right)
- Per-column fixed text positioning
- Simple/Advanced positioning modes
"""

from .exceptions import (
    AppError,
    ImageDownloadError,
    ExcelValidationError,
    PPTXGenerationError,
    ConfigurationError,
)
from .config import Config, get_config
from .validators import sanitize_text
from .image_handler import ImageLoader, ImageResult, load_image, extract_excel_images
from .excel_handler import ExcelProcessor
from .pptx_generator import (
    PPTXGenerator,
    SlideConfig,
    ColumnFormat,
    ImageAlignment,
    ColumnPosition,
    ImageElement,
    TextGroup,
    IMG_SIZE_FIT_BOX,
    IMG_SIZE_FIT_WIDTH,
    IMG_SIZE_FIT_HEIGHT,
    IMG_SIZE_STRETCH,
    IMG_ALIGN_TOP,
    IMG_ALIGN_CENTER,
    IMG_ALIGN_BOTTOM,
    IMG_ALIGN_LEFT,
    IMG_ALIGN_RIGHT,
    TEMPLATE_MODE_BLANK,
    TEMPLATE_MODE_PLACEHOLDER,
)

__all__ = [
    # Exceptions
    "AppError",
    "ImageDownloadError",
    "ExcelValidationError",
    "PPTXGenerationError",
    "ConfigurationError",
    # Config
    "Config",
    "get_config",
    # Validators
    "sanitize_text",
    # Image handling
    "ImageLoader",
    "ImageResult",
    "load_image",
    "extract_excel_images",
    # Excel handling
    "ExcelProcessor",
    # PPTX generation
    "PPTXGenerator",
    "SlideConfig",
    "ColumnFormat",
    "ImageAlignment",
    "ColumnPosition",
    "ImageElement",
    "TextGroup",
    # Image sizing modes
    "IMG_SIZE_FIT_BOX",
    "IMG_SIZE_FIT_WIDTH",
    "IMG_SIZE_FIT_HEIGHT",
    "IMG_SIZE_STRETCH",
    # Image alignment (NEW in v6.0)
    "IMG_ALIGN_TOP",
    "IMG_ALIGN_CENTER",
    "IMG_ALIGN_BOTTOM",
    "IMG_ALIGN_LEFT",
    "IMG_ALIGN_RIGHT",
    # Template modes (NEW in v5.1)
    "TEMPLATE_MODE_BLANK",
    "TEMPLATE_MODE_PLACEHOLDER",
]

__version__ = "8.0.0"
