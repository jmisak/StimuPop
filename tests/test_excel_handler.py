"""
Tests for Excel file handling.
"""

import pytest
import pandas as pd
from io import BytesIO

from src.excel_handler import (
    ExcelProcessor,
    parse_column_input,
    read_excel_file,
)
from src.exceptions import ExcelValidationError


class TestExcelProcessor:
    """Tests for ExcelProcessor class."""

    def test_read_excel_valid(self, sample_excel_bytes):
        """Test reading a valid Excel file."""
        processor = ExcelProcessor()
        df = processor.read_excel(sample_excel_bytes, "test.xlsx")
        assert len(df) == 3
        assert 'A' in df.columns
        assert 'B' in df.columns

    def test_read_excel_empty(self):
        """Test handling empty Excel file."""
        processor = ExcelProcessor()

        # Create empty Excel
        buffer = BytesIO()
        pd.DataFrame().to_excel(buffer, index=False)
        buffer.seek(0)

        with pytest.raises(ExcelValidationError) as exc_info:
            processor.read_excel(buffer.getvalue(), "empty.xlsx")
        assert "empty" in str(exc_info.value).lower()

    def test_read_excel_too_large(self):
        """Test file size limit."""
        processor = ExcelProcessor(max_upload_size_mb=0.001)  # 1KB limit

        # Create larger file
        large_df = pd.DataFrame({'A': list(range(1000))})
        buffer = BytesIO()
        large_df.to_excel(buffer, index=False)
        buffer.seek(0)

        with pytest.raises(ExcelValidationError) as exc_info:
            processor.read_excel(buffer.getvalue(), "large.xlsx")
        assert "exceeds" in str(exc_info.value).lower()

    def test_read_excel_row_limit(self):
        """Test row count limiting."""
        processor = ExcelProcessor(max_rows=5)

        # Create file with many rows
        df = pd.DataFrame({'A': list(range(100))})
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        result = processor.read_excel(buffer.getvalue(), "many_rows.xlsx")
        assert len(result) == 5


class TestValidateColumns:
    """Tests for column validation."""

    def test_validate_letter_columns(self, sample_dataframe):
        """Test validating columns by letter."""
        processor = ExcelProcessor()
        img_col, text_cols = processor.validate_columns(
            sample_dataframe, "B", ["C", "D"]
        )
        assert img_col == "B"
        assert text_cols == ["C", "D"]

    def test_validate_name_columns(self, sample_dataframe_with_names):
        """Test validating columns by name."""
        processor = ExcelProcessor()
        img_col, text_cols = processor.validate_columns(
            sample_dataframe_with_names,
            "Image",
            ["Title", "Description"]
        )
        assert img_col == "Image"
        assert "Title" in text_cols
        assert "Description" in text_cols

    def test_validate_missing_image_column(self, sample_dataframe):
        """Test error for missing image column."""
        processor = ExcelProcessor()
        with pytest.raises(ExcelValidationError) as exc_info:
            processor.validate_columns(
                sample_dataframe, "Z", ["C", "D"]
            )
        assert "not found" in str(exc_info.value).lower()

    def test_validate_missing_text_columns(self, sample_dataframe):
        """Test handling missing text columns."""
        processor = ExcelProcessor()
        with pytest.raises(ExcelValidationError) as exc_info:
            processor.validate_columns(
                sample_dataframe, "B", ["X", "Y", "Z"]
            )
        assert "no valid text columns" in str(exc_info.value).lower()

    def test_validate_partial_text_columns(self, sample_dataframe):
        """Test partial match of text columns."""
        processor = ExcelProcessor()
        img_col, text_cols = processor.validate_columns(
            sample_dataframe, "B", ["C", "X", "D"]
        )
        # X should be skipped, C and D kept
        assert len(text_cols) == 2
        assert "C" in text_cols
        assert "D" in text_cols

    def test_validate_case_insensitive(self, sample_dataframe_with_names):
        """Test case-insensitive column matching."""
        processor = ExcelProcessor()
        img_col, text_cols = processor.validate_columns(
            sample_dataframe_with_names,
            "image",  # lowercase
            ["TITLE"]    # uppercase
        )
        assert img_col == "Image"
        assert "Title" in text_cols


class TestLetterToIndex:
    """Tests for Excel letter to index conversion."""

    def test_single_letters(self):
        assert ExcelProcessor._letter_to_index("A") == 0
        assert ExcelProcessor._letter_to_index("B") == 1
        assert ExcelProcessor._letter_to_index("Z") == 25

    def test_double_letters(self):
        assert ExcelProcessor._letter_to_index("AA") == 26
        assert ExcelProcessor._letter_to_index("AB") == 27
        assert ExcelProcessor._letter_to_index("AZ") == 51
        assert ExcelProcessor._letter_to_index("BA") == 52


class TestGetSlideData:
    """Tests for extracting slide data."""

    def test_get_slide_data(self, sample_dataframe):
        """Test basic slide data extraction with column identity."""
        processor = ExcelProcessor()
        slides = processor.get_slide_data(
            sample_dataframe, "B", ["C", "D"], preserve_column_identity=True
        )

        assert len(slides) == 3
        # Image column is empty in test data
        assert slides[0]["image_source"] is None or slides[0]["image_source"] == ""
        # Check cell reference is generated
        assert "image_cell" in slides[0]
        # With preserve_column_identity=True, text_content items are dicts
        assert isinstance(slides[0]["text_content"][0], dict)
        assert "column" in slides[0]["text_content"][0]
        assert "text" in slides[0]["text_content"][0]
        assert slides[0]["text_content"][0]["text"] == "Title 1"

    def test_get_slide_data_without_column_identity(self, sample_dataframe):
        """Test slide data extraction without column identity (backward compat)."""
        processor = ExcelProcessor()
        slides = processor.get_slide_data(
            sample_dataframe, "B", ["C", "D"], preserve_column_identity=False
        )

        assert len(slides) == 3
        # With preserve_column_identity=False, text_content items are strings
        assert isinstance(slides[0]["text_content"][0], str)
        assert slides[0]["text_content"][0] == "Title 1"
        assert slides[0]["text_content"][1] == "Description 1"

    def test_get_slide_data_column_letters(self, sample_dataframe):
        """Test that column letters are correctly assigned."""
        processor = ExcelProcessor()
        slides = processor.get_slide_data(
            sample_dataframe, "B", ["C", "D"], preserve_column_identity=True
        )

        # Column C should be the 3rd column (index 2), letter "C"
        # Column D should be the 4th column (index 3), letter "D"
        assert slides[0]["text_content"][0]["column"] == "C"
        assert slides[0]["text_content"][1]["column"] == "D"

    def test_get_slide_data_sanitizes(self, sample_dataframe):
        """Test that text is sanitized."""
        # Add control characters to data
        sample_dataframe.at[0, 'C'] = "Title\x00with\x01nulls"

        processor = ExcelProcessor()
        slides = processor.get_slide_data(
            sample_dataframe, "B", ["C"], sanitize=True, preserve_column_identity=False
        )

        assert "\x00" not in slides[0]["text_content"][0]
        assert "\x01" not in slides[0]["text_content"][0]

    def test_get_slide_data_handles_nan(self, sample_dataframe):
        """Test handling of NaN values."""
        sample_dataframe.at[0, 'C'] = None
        sample_dataframe.at[0, 'B'] = None

        processor = ExcelProcessor()
        slides = processor.get_slide_data(
            sample_dataframe, "B", ["C", "D"], preserve_column_identity=False
        )

        assert slides[0]["image_source"] is None
        # Title 1 should not be in text_content since it's None
        assert len(slides[0]["text_content"]) == 1  # Only Description


class TestGetSummary:
    """Tests for DataFrame summary."""

    def test_get_summary(self, sample_dataframe):
        """Test summary generation."""
        processor = ExcelProcessor()
        summary = processor.get_summary(sample_dataframe)

        assert summary["row_count"] == 3
        assert summary["column_count"] == 4
        assert "A" in summary["columns"]
        assert "A" in summary["column_letters"]
        assert "B" in summary["column_letters"]


class TestParseColumnInput:
    """Tests for column input parsing."""

    def test_simple_parse(self):
        assert parse_column_input("A,B,C") == ["A", "B", "C"]

    def test_with_spaces(self):
        assert parse_column_input("A, B, C") == ["A", "B", "C"]
        assert parse_column_input(" A , B , C ") == ["A", "B", "C"]

    def test_empty_string(self):
        assert parse_column_input("") == []

    def test_single_column(self):
        assert parse_column_input("A") == ["A"]

    def test_with_names(self):
        result = parse_column_input("Title, Description, Price")
        assert result == ["Title", "Description", "Price"]


# ---------------------------------------------------------------------------
# Lightweight stubs for ImageElement / TextGroup used by multi-element methods
# ---------------------------------------------------------------------------

class _ImageElement:
    """Stub mirroring the ImageElement dataclass for test isolation."""
    def __init__(self, column: str, placeholder_name: str):
        self.column = column
        self.placeholder_name = placeholder_name


class _TextGroup:
    """Stub mirroring the TextGroup dataclass for test isolation."""
    def __init__(self, columns: list, placeholder_name: str):
        self.columns = columns
        self.placeholder_name = placeholder_name


class TestValidateColumnsMulti:
    """Tests for multi-element column validation (v8.0)."""

    def test_single_image_single_text_group(self, sample_dataframe):
        """Basic happy path with one image element and one text group."""
        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C", "D"], "Text 1")]

        resolved_imgs, resolved_txts = processor.validate_columns_multi(
            sample_dataframe, images, texts
        )

        assert len(resolved_imgs) == 1
        assert resolved_imgs[0] == ("B", "Picture 1")

        assert len(resolved_txts) == 1
        cols, ph = resolved_txts[0]
        assert set(cols) == {"C", "D"}
        assert ph == "Text 1"

    def test_multiple_image_elements(self, sample_dataframe_with_names):
        """Validate two separate image elements pointing at different columns."""
        processor = ExcelProcessor()
        images = [
            _ImageElement("Image", "Picture 1"),
            _ImageElement("ID", "Picture 2"),
        ]
        texts = [_TextGroup(["Title"], "Text 1")]

        resolved_imgs, _ = processor.validate_columns_multi(
            sample_dataframe_with_names, images, texts
        )

        assert len(resolved_imgs) == 2
        assert resolved_imgs[0] == ("Image", "Picture 1")
        assert resolved_imgs[1] == ("ID", "Picture 2")

    def test_missing_image_column_raises(self, sample_dataframe):
        """Must raise if an image element references a nonexistent column."""
        processor = ExcelProcessor()
        images = [_ImageElement("Z", "Picture 1")]
        texts = [_TextGroup(["C"], "Text 1")]

        with pytest.raises(ExcelValidationError) as exc_info:
            processor.validate_columns_multi(sample_dataframe, images, texts)
        assert "not found" in str(exc_info.value).lower()

    def test_partial_text_columns_warns(self, sample_dataframe):
        """Individual missing text columns should be skipped, not fatal."""
        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C", "NONEXISTENT", "D"], "Text 1")]

        resolved_imgs, resolved_txts = processor.validate_columns_multi(
            sample_dataframe, images, texts
        )

        cols, _ = resolved_txts[0]
        assert len(cols) == 2
        assert "NONEXISTENT" not in cols

    def test_all_text_columns_missing_raises(self, sample_dataframe):
        """Must raise when ALL columns of a text group are missing."""
        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["X", "Y"], "Text 1")]

        with pytest.raises(ExcelValidationError) as exc_info:
            processor.validate_columns_multi(sample_dataframe, images, texts)
        assert "no valid text columns" in str(exc_info.value).lower()

    def test_case_insensitive_resolution(self, sample_dataframe_with_names):
        """Column resolution should be case-insensitive."""
        processor = ExcelProcessor()
        images = [_ImageElement("image", "Picture 1")]
        texts = [_TextGroup(["title", "DESCRIPTION"], "Text 1")]

        resolved_imgs, resolved_txts = processor.validate_columns_multi(
            sample_dataframe_with_names, images, texts
        )

        assert resolved_imgs[0][0] == "Image"
        cols, _ = resolved_txts[0]
        assert "Title" in cols
        assert "Description" in cols

    def test_empty_image_elements_list(self, sample_dataframe):
        """Empty image_elements should return empty resolved_images."""
        processor = ExcelProcessor()
        resolved_imgs, resolved_txts = processor.validate_columns_multi(
            sample_dataframe, [], [_TextGroup(["C"], "Text 1")]
        )
        assert resolved_imgs == []

    def test_empty_text_groups_list(self, sample_dataframe):
        """Empty text_groups should return empty resolved_texts."""
        processor = ExcelProcessor()
        resolved_imgs, resolved_txts = processor.validate_columns_multi(
            sample_dataframe, [_ImageElement("B", "Picture 1")], []
        )
        assert resolved_txts == []


class TestGetSlideDataMulti:
    """Tests for multi-element slide data extraction (v8.0)."""

    def test_basic_output_structure(self, sample_dataframe):
        """Verify the shape of every dict returned."""
        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C", "D"], "Text 1")]

        slides = processor.get_slide_data_multi(
            sample_dataframe, images, texts
        )

        assert len(slides) == 3
        for slide in slides:
            assert "row_index" in slide
            assert "image_sources" in slide
            assert "text_contents" in slide
            # Legacy fields
            assert "image_source" in slide
            assert "image_cell" in slide
            assert "text_content" in slide

    def test_image_sources_populated(self, sample_dataframe_with_names):
        """Image source values and cell refs should appear per element."""
        sample_dataframe_with_names.at[0, "Image"] = "photo.png"

        processor = ExcelProcessor()
        images = [_ImageElement("Image", "Picture 1")]
        texts = [_TextGroup(["Title"], "Text 1")]

        slides = processor.get_slide_data_multi(
            sample_dataframe_with_names, images, texts
        )

        first = slides[0]
        assert len(first["image_sources"]) == 1
        assert first["image_sources"][0]["image_source"] == "photo.png"
        assert first["image_sources"][0]["placeholder_name"] == "Picture 1"
        # Cell ref should be column B (Image is 2nd col) row 2
        assert first["image_sources"][0]["image_cell"] == "B2"

    def test_text_contents_populated(self, sample_dataframe):
        """Text content items should carry column letter and text."""
        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C", "D"], "Text 1")]

        slides = processor.get_slide_data_multi(
            sample_dataframe, images, texts
        )

        first_text = slides[0]["text_contents"][0]
        assert first_text["placeholder_name"] == "Text 1"
        assert len(first_text["text_content"]) == 2
        assert first_text["text_content"][0]["text"] == "Title 1"
        assert first_text["text_content"][1]["text"] == "Description 1"

    def test_multiple_image_elements_per_row(self, sample_dataframe_with_names):
        """Each image element should produce its own entry in image_sources."""
        sample_dataframe_with_names.at[0, "Image"] = "front.png"
        sample_dataframe_with_names.at[0, "ID"] = "42"

        processor = ExcelProcessor()
        images = [
            _ImageElement("Image", "Picture 1"),
            _ImageElement("ID", "Picture 2"),
        ]
        texts = [_TextGroup(["Title"], "Text 1")]

        slides = processor.get_slide_data_multi(
            sample_dataframe_with_names, images, texts
        )

        first = slides[0]
        assert len(first["image_sources"]) == 2
        assert first["image_sources"][0]["placeholder_name"] == "Picture 1"
        assert first["image_sources"][1]["placeholder_name"] == "Picture 2"
        assert first["image_sources"][1]["image_source"] == "42"

    def test_multiple_text_groups(self, sample_dataframe_with_names):
        """Multiple text groups should each have their own entry."""
        processor = ExcelProcessor()
        images = [_ImageElement("Image", "Picture 1")]
        texts = [
            _TextGroup(["Title"], "Text 1"),
            _TextGroup(["Description", "Price"], "Text 2"),
        ]

        slides = processor.get_slide_data_multi(
            sample_dataframe_with_names, images, texts
        )

        first = slides[0]
        assert len(first["text_contents"]) == 2
        assert first["text_contents"][0]["placeholder_name"] == "Text 1"
        assert first["text_contents"][1]["placeholder_name"] == "Text 2"
        # Text 2 should have two items (Description + Price)
        assert len(first["text_contents"][1]["text_content"]) == 2

    def test_legacy_fields_from_first_elements(self, sample_dataframe_with_names):
        """Legacy image_source/image_cell/text_content should mirror first elements."""
        sample_dataframe_with_names.at[0, "Image"] = "legacy.png"

        processor = ExcelProcessor()
        images = [
            _ImageElement("Image", "Picture 1"),
            _ImageElement("ID", "Picture 2"),
        ]
        texts = [
            _TextGroup(["Title"], "Text 1"),
            _TextGroup(["Description"], "Text 2"),
        ]

        slides = processor.get_slide_data_multi(
            sample_dataframe_with_names, images, texts
        )

        first = slides[0]
        assert first["image_source"] == first["image_sources"][0]["image_source"]
        assert first["image_cell"] == first["image_sources"][0]["image_cell"]
        assert first["text_content"] == first["text_contents"][0]["text_content"]

    def test_sanitize_flag(self, sample_dataframe):
        """Control chars should be stripped when sanitize=True."""
        sample_dataframe.at[0, "C"] = "Title\x00dirty"

        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C"], "Text 1")]

        slides_clean = processor.get_slide_data_multi(
            sample_dataframe, images, texts, sanitize=True
        )
        slides_raw = processor.get_slide_data_multi(
            sample_dataframe, images, texts, sanitize=False
        )

        clean_text = slides_clean[0]["text_contents"][0]["text_content"][0]["text"]
        raw_text = slides_raw[0]["text_contents"][0]["text_content"][0]["text"]

        assert "\x00" not in clean_text
        assert "\x00" in raw_text

    def test_nan_values_skipped(self, sample_dataframe):
        """NaN image sources and text values should be omitted."""
        sample_dataframe.at[0, "B"] = None
        sample_dataframe.at[0, "C"] = None

        processor = ExcelProcessor()
        images = [_ImageElement("B", "Picture 1")]
        texts = [_TextGroup(["C", "D"], "Text 1")]

        slides = processor.get_slide_data_multi(
            sample_dataframe, images, texts
        )

        first = slides[0]
        assert first["image_sources"][0]["image_source"] is None
        # Only D should remain
        assert len(first["text_contents"][0]["text_content"]) == 1
        assert first["text_contents"][0]["text_content"][0]["text"] == "Description 1"

    def test_empty_elements_produce_empty_legacy(self, sample_dataframe):
        """With no image/text elements, legacy fields use safe defaults."""
        processor = ExcelProcessor()

        slides = processor.get_slide_data_multi(
            sample_dataframe, [], []
        )

        for slide in slides:
            assert slide["image_source"] is None
            assert slide["text_content"] == []
