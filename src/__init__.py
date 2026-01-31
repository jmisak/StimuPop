"""
StimuPop - Source Package v6.1

Production-grade modules for Excel to PPTX conversion with embedded images.

New in v6.1:
- UI improvements and documentation updates
- Layout Position controls hidden in Template mode

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

__version__ = "6.2.0"
