"""
StimuPop - Source Package v5.1

Production-grade modules for Excel to PPTX conversion with embedded images.

New in v5.1:
- Template-based placeholder population
- Configurable paragraph spacing
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
    IMG_SIZE_FIT_BOX,
    IMG_SIZE_FIT_WIDTH,
    IMG_SIZE_FIT_HEIGHT,
    IMG_SIZE_STRETCH,
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
    # Image sizing modes
    "IMG_SIZE_FIT_BOX",
    "IMG_SIZE_FIT_WIDTH",
    "IMG_SIZE_FIT_HEIGHT",
    "IMG_SIZE_STRETCH",
    # Template modes (NEW in v5.1)
    "TEMPLATE_MODE_BLANK",
    "TEMPLATE_MODE_PLACEHOLDER",
]

__version__ = "5.1.0"
