"""
Input validation for StimuPop.

Provides validation including:
- Text sanitization
- Image format validation
"""

import re
from typing import List, Optional

from .config import get_config
from .exceptions import ValidationError
from .logging_config import get_logger

logger = get_logger(__name__)


def sanitize_text(text: str, max_length: int = 10000) -> str:
    """
    Sanitize text content from Excel for safe display.

    Removes or escapes potentially dangerous content:
    - Control characters (except newlines and tabs)
    - Null bytes
    - Excessive whitespace
    - Excessively long content

    Args:
        text: Text to sanitize
        max_length: Maximum allowed length

    Returns:
        Sanitized text

    Usage:
        safe_text = sanitize_text(excel_cell_value)
    """
    if not text:
        return ""

    # Convert to string if needed
    text = str(text)

    # Remove null bytes
    text = text.replace("\x00", "")

    # Remove control characters except \n, \r, \t
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", "", text)

    # Normalize excessive whitespace (preserve single newlines)
    text = re.sub(r"[ \t]+", " ", text)  # Multiple spaces/tabs to single space
    text = re.sub(r"\n{3,}", "\n\n", text)  # Max 2 consecutive newlines

    # Truncate if too long
    if len(text) > max_length:
        text = text[:max_length] + "..."
        logger.warning(f"Text truncated from {len(text)} to {max_length} characters")

    return text.strip()


def validate_image_format(filename: str, allowed_formats: Optional[List[str]] = None) -> bool:
    """
    Validate that a filename has an allowed image extension.

    Args:
        filename: Filename or path to check
        allowed_formats: List of allowed extensions (with dots, e.g., ['.jpg', '.png'])

    Returns:
        True if format is allowed

    Raises:
        ValidationError: If format is not allowed
    """
    if allowed_formats is None:
        config = get_config()
        allowed_formats = config.images.allowed_formats

    # Get extension from filename
    filename_lower = filename.lower()

    for fmt in allowed_formats:
        if filename_lower.endswith(fmt.lower()):
            return True

    # Extract just the extension for error message
    parts = filename.rsplit(".", 1)
    ext = f".{parts[-1]}" if len(parts) > 1 else "none"

    raise ValidationError(
        f"Image format '{ext}' not allowed. Allowed formats: {', '.join(allowed_formats)}",
        field=filename
    )
