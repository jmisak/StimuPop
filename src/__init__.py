"""
StimuPop - Source Package

Production-grade modules for Excel to PPTX conversion with embedded images.
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
from .pptx_generator import PPTXGenerator, SlideConfig, ColumnFormat

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
]

__version__ = "2.2.0"
