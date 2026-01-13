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
