"""
Tests for input validation.
"""

import pytest

from src.validators import (
    sanitize_text,
    validate_image_format,
)
from src.exceptions import ValidationError


class TestSanitizeText:
    """Tests for text sanitization."""

    def test_normal_text(self):
        assert sanitize_text("Hello World") == "Hello World"

    def test_removes_null_bytes(self):
        assert sanitize_text("Hello\x00World") == "HelloWorld"

    def test_removes_control_chars(self):
        assert sanitize_text("Hello\x01\x02World") == "HelloWorld"

    def test_preserves_newlines(self):
        result = sanitize_text("Line1\nLine2")
        assert "Line1" in result
        assert "Line2" in result

    def test_preserves_tabs(self):
        assert sanitize_text("Col1\tCol2") == "Col1 Col2"

    def test_normalizes_multiple_spaces(self):
        assert sanitize_text("Hello    World") == "Hello World"

    def test_normalizes_multiple_newlines(self):
        result = sanitize_text("Para1\n\n\n\nPara2")
        assert result.count("\n") <= 2

    def test_truncates_long_text(self):
        long_text = "a" * 20000
        result = sanitize_text(long_text, max_length=100)
        assert len(result) <= 103  # 100 + "..."

    def test_strips_whitespace(self):
        assert sanitize_text("  Hello World  ") == "Hello World"

    def test_empty_text(self):
        assert sanitize_text("") == ""

    def test_none_like_text(self):
        assert sanitize_text(None) == ""


class TestValidateImageFormat:
    """Tests for image format validation."""

    def test_valid_jpg(self):
        assert validate_image_format("image.jpg") is True
        assert validate_image_format("image.jpeg") is True

    def test_valid_png(self):
        assert validate_image_format("image.png") is True

    def test_valid_gif(self):
        assert validate_image_format("image.gif") is True

    def test_valid_webp(self):
        assert validate_image_format("image.webp") is True

    def test_valid_path_with_extension(self):
        assert validate_image_format("C:\\Images\\photo.jpg") is True
        assert validate_image_format("/home/user/images/photo.png") is True

    def test_case_insensitive(self):
        assert validate_image_format("image.JPG") is True
        assert validate_image_format("image.PNG") is True

    def test_invalid_format(self):
        with pytest.raises(ValidationError) as exc_info:
            validate_image_format("document.pdf")
        assert "not allowed" in str(exc_info.value).lower()

    def test_custom_allowed_formats(self):
        assert validate_image_format("image.tiff", allowed_formats=[".tiff"]) is True

    def test_no_extension(self):
        with pytest.raises(ValidationError):
            validate_image_format("imagefile")
