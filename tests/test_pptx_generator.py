"""
Tests for PowerPoint generation.
"""

import pytest
from io import BytesIO
from unittest.mock import MagicMock, patch

from PIL import Image

from src.pptx_generator import (
    PPTXGenerator,
    SlideConfig,
    SlideResult,
    GenerationResult,
    ColumnFormat,
    create_presentation,
)
from src.image_handler import ImageResult
from pptx.dml.color import RGBColor


def create_test_image_result():
    """Create a valid ImageResult for testing."""
    img = Image.new('RGB', (100, 100), color='blue')
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)

    return ImageResult(
        source="test.png",
        success=True,
        data=buffer,
        width=100,
        height=100,
        format="PNG",
        size_bytes=len(buffer.getvalue())
    )


def create_failed_image_result():
    """Create a failed ImageResult for testing."""
    return ImageResult(
        source="failed.png",
        success=False,
        error="Image not found"
    )


class TestColumnFormat:
    """Tests for ColumnFormat dataclass."""

    def test_default_values(self):
        """Test ColumnFormat with minimal required fields."""
        fmt = ColumnFormat(column="C")
        assert fmt.font_size == 14
        assert fmt.bold is False
        assert fmt.italic is False
        assert fmt.font_name == "Calibri"
        assert fmt.color == "000000"

    def test_custom_values(self):
        """Test ColumnFormat with custom values."""
        fmt = ColumnFormat(
            column="D",
            font_size=24,
            bold=True,
            italic=True,
            font_name="Arial",
            color="FF0000"
        )
        assert fmt.column == "D"
        assert fmt.font_size == 24
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.font_name == "Arial"
        assert fmt.color == "FF0000"

    def test_get_rgb_color_black(self):
        """Test RGB color conversion for black."""
        fmt = ColumnFormat(column="C", color="000000")
        rgb = fmt.get_rgb_color()
        assert isinstance(rgb, RGBColor)
        assert rgb == RGBColor(0, 0, 0)

    def test_get_rgb_color_red(self):
        """Test RGB color conversion for red."""
        fmt = ColumnFormat(column="C", color="FF0000")
        rgb = fmt.get_rgb_color()
        assert rgb == RGBColor(255, 0, 0)

    def test_get_rgb_color_custom(self):
        """Test RGB color conversion for custom color."""
        fmt = ColumnFormat(column="C", color="1A2B3C")
        rgb = fmt.get_rgb_color()
        assert rgb == RGBColor(26, 43, 60)


class TestSlideConfig:
    """Tests for SlideConfig dataclass."""

    def test_default_values(self):
        """Test SlideConfig with minimal required fields."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C", "D"]
        )
        assert config.img_width == 5.5
        assert config.font_size == 14
        assert config.orientation == "portrait"
        assert config.column_formats is None

    def test_custom_values(self):
        """Test SlideConfig with custom values."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C"],
            img_width=6.0,
            font_size=18,
            orientation="landscape"
        )
        assert config.img_width == 6.0
        assert config.font_size == 18
        assert config.orientation == "landscape"

    def test_with_column_formats(self):
        """Test SlideConfig with column_formats."""
        formats = {
            "C": ColumnFormat(column="C", font_size=20, bold=True),
            "D": ColumnFormat(column="D", font_size=12, italic=True),
        }
        config = SlideConfig(
            img_column="B",
            text_columns=["C", "D"],
            column_formats=formats
        )
        assert config.column_formats is not None
        assert len(config.column_formats) == 2

    def test_get_column_format_with_match(self):
        """Test get_column_format returns matching format."""
        fmt_c = ColumnFormat(column="C", font_size=20, bold=True)
        config = SlideConfig(
            img_column="B",
            text_columns=["C", "D"],
            column_formats={"C": fmt_c}
        )
        result = config.get_column_format("C")
        assert result.font_size == 20
        assert result.bold is True

    def test_get_column_format_fallback(self):
        """Test get_column_format returns default for missing column."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C", "D"],
            font_size=16,
            column_formats={"C": ColumnFormat(column="C", font_size=20)}
        )
        # D is not in column_formats, should fall back to default
        result = config.get_column_format("D")
        assert result.font_size == 16  # Falls back to config.font_size
        assert result.bold is False

    def test_get_column_format_no_formats(self):
        """Test get_column_format when column_formats is None."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C"],
            font_size=18
        )
        result = config.get_column_format("C")
        assert result.font_size == 18


class TestPPTXGenerator:
    """Tests for PPTXGenerator class."""

    def test_generate_empty_data(self):
        """Test generating with empty slide data."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        result = generator.generate([])

        assert result.success is False
        assert "no slide data" in result.error.lower()

    def test_generate_text_only(self):
        """Test generating slides with text only."""
        config = SlideConfig(img_column="B", text_columns=["C", "D"])
        generator = PPTXGenerator(config)

        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title 1", "Desc 1"]},
            {"row_index": 1, "image_source": None, "text_content": ["Title 2", "Desc 2"]},
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        assert result.slides_generated == 2
        assert result.slides_with_images == 0
        assert result.presentation is not None

    @patch.object(PPTXGenerator, '_create_slide')
    def test_generate_calls_create_slide(self, mock_create_slide):
        """Test that generate calls _create_slide for each row."""
        mock_create_slide.return_value = SlideResult(index=0, success=True)

        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title"]},
            {"row_index": 1, "image_source": None, "text_content": ["Title"]},
            {"row_index": 2, "image_source": None, "text_content": ["Title"]},
        ]

        result = generator.generate(slide_data)

        assert mock_create_slide.call_count == 3

    def test_generate_with_progress_callback(self):
        """Test progress callback is called."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        callbacks = []

        def callback(status, current, total):
            callbacks.append((status, current, total))

        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title"]},
        ]

        generator.generate(slide_data, progress_callback=callback)

        assert len(callbacks) > 0

    def test_generate_portrait_orientation(self):
        """Test portrait orientation dimensions."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C"],
            orientation="portrait"
        )
        generator = PPTXGenerator(config)

        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title"]},
        ]

        result = generator.generate(slide_data)

        # Portrait: width < height
        prs = result.presentation
        assert prs.slide_width < prs.slide_height

    def test_generate_landscape_orientation(self):
        """Test landscape orientation dimensions."""
        config = SlideConfig(
            img_column="B",
            text_columns=["C"],
            orientation="landscape"
        )
        generator = PPTXGenerator(config)

        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title"]},
        ]

        result = generator.generate(slide_data)

        # Landscape: width > height
        prs = result.presentation
        assert prs.slide_width > prs.slide_height


class TestSlideCreation:
    """Tests for individual slide creation."""

    def test_slide_with_text(self):
        """Test creating slide with text content."""
        config = SlideConfig(img_column="B", text_columns=["C", "D"])
        generator = PPTXGenerator(config)

        slide_data = [
            {
                "row_index": 0,
                "image_source": None,
                "text_content": ["Main Title", "Description text"]
            },
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        prs = result.presentation
        assert len(prs.slides) == 1

        # Check that text was added
        slide = prs.slides[0]
        shapes_with_text = [s for s in slide.shapes if hasattr(s, 'text_frame')]
        assert len(shapes_with_text) > 0

    def test_slide_handles_failed_image_gracefully(self):
        """Test that failed images don't crash generation."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        # Create failed image result
        failed_result = create_failed_image_result()
        embedded_images = {"B2": failed_result}

        slide_data = [
            {
                "row_index": 0,
                "image_cell": "B2",
                "image_source": None,
                "text_content": ["Title despite missing image"]
            },
        ]

        result = generator.generate(slide_data, embedded_images=embedded_images)

        # Should still succeed, just without image
        assert result.success is True
        assert result.slides_generated == 1
        assert result.slides_with_images == 0
        assert result.slides_with_errors == 1

    def test_slide_with_embedded_image(self):
        """Test creating slide with embedded image."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        # Create valid image result
        img_result = create_test_image_result()
        embedded_images = {"B2": img_result}

        slide_data = [
            {
                "row_index": 0,
                "image_cell": "B2",
                "image_source": None,
                "text_content": ["Title with image"]
            },
        ]

        result = generator.generate(slide_data, embedded_images=embedded_images)

        assert result.success is True
        assert result.slides_with_images == 1


class TestGenerationResult:
    """Tests for GenerationResult dataclass."""

    def test_default_slide_results(self):
        """Test that slide_results defaults to empty list."""
        result = GenerationResult(success=True)
        assert result.slide_results == []

    def test_with_slide_results(self):
        """Test with explicit slide results."""
        slides = [
            SlideResult(index=0, success=True, has_image=True),
            SlideResult(index=1, success=True, has_image=False),
        ]
        result = GenerationResult(
            success=True,
            slides_generated=2,
            slide_results=slides
        )
        assert len(result.slide_results) == 2


class TestConvenienceFunction:
    """Tests for create_presentation convenience function."""

    def test_create_presentation(self):
        """Test create_presentation function."""
        slide_data = [
            {"row_index": 0, "image_source": None, "text_content": ["Title"]},
        ]

        result = create_presentation(slide_data)

        assert isinstance(result, GenerationResult)
        assert result.success is True


class TestPerColumnFormatting:
    """Tests for per-column text formatting."""

    def test_generate_with_column_formats(self):
        """Test generating slides with per-column formatting."""
        formats = {
            "C": ColumnFormat(column="C", font_size=24, bold=True, color="FF0000"),
            "D": ColumnFormat(column="D", font_size=12, italic=True, color="0000FF"),
        }
        config = SlideConfig(
            img_column="B",
            text_columns=["C", "D"],
            column_formats=formats
        )
        generator = PPTXGenerator(config)

        # New format with column identity preserved
        slide_data = [
            {
                "row_index": 0,
                "image_source": None,
                "text_content": [
                    {"column": "C", "text": "Title"},
                    {"column": "D", "text": "Description"}
                ]
            },
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        assert result.slides_generated == 1

    def test_generate_with_dict_text_content(self):
        """Test generating slides with dict-style text content."""
        config = SlideConfig(img_column="B", text_columns=["C", "D"])
        generator = PPTXGenerator(config)

        slide_data = [
            {
                "row_index": 0,
                "image_source": None,
                "text_content": [
                    {"column": "C", "text": "Main Title"},
                    {"column": "D", "text": "Subtitle"}
                ]
            },
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        prs = result.presentation
        assert len(prs.slides) == 1

    def test_backward_compat_string_text_content(self):
        """Test backward compatibility with string text content."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        # Old format: list of strings
        slide_data = [
            {
                "row_index": 0,
                "image_source": None,
                "text_content": ["Title", "Subtitle"]
            },
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        assert result.slides_generated == 1

    def test_mixed_text_content_formats(self):
        """Test that generator handles mixed text formats."""
        config = SlideConfig(img_column="B", text_columns=["C"])
        generator = PPTXGenerator(config)

        # One slide with new format, one with old
        slide_data = [
            {
                "row_index": 0,
                "image_source": None,
                "text_content": [{"column": "C", "text": "New Format"}]
            },
            {
                "row_index": 1,
                "image_source": None,
                "text_content": ["Old Format"]
            },
        ]

        result = generator.generate(slide_data)

        assert result.success is True
        assert result.slides_generated == 2
