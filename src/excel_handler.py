"""
Excel file handling for StimuPop.

Provides secure Excel processing with:
- File validation
- Column existence checking
- Data sanitization
- Row limits
- Caching for preview
"""

from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd

from .config import get_config
from .exceptions import ExcelValidationError
from .logging_config import get_logger
from .validators import sanitize_text

logger = get_logger(__name__)


class ExcelProcessor:
    """
    Processes Excel files for slide generation.

    Features:
    - File format validation
    - Column existence checking
    - Data sanitization
    - Row count limits
    - Efficient preview generation

    Usage:
        processor = ExcelProcessor()
        df = processor.read_excel(file_bytes, filename="data.xlsx")

        # Validate columns exist
        processor.validate_columns(df, img_column="B", text_columns=["C", "D"])

        # Get sanitized data
        rows = processor.get_slide_data(df, img_column="B", text_columns=["C", "D"])
    """

    def __init__(
        self,
        max_rows: int = 1000,
        max_upload_size_mb: Optional[int] = None
    ):
        """
        Initialize Excel processor.

        Args:
            max_rows: Maximum number of rows to process
            max_upload_size_mb: Maximum file size in MB
        """
        config = get_config()
        self.max_rows = max_rows
        self.max_upload_size_bytes = (
            max_upload_size_mb * 1024 * 1024 if max_upload_size_mb
            else config.app.max_upload_size_bytes
        )

    def read_excel(
        self,
        file_data: bytes,
        filename: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Read and validate an Excel file.

        Args:
            file_data: Excel file content as bytes
            filename: Optional filename for error messages

        Returns:
            DataFrame with Excel data

        Raises:
            ExcelValidationError: If file is invalid
        """
        # Check file size
        if len(file_data) > self.max_upload_size_bytes:
            raise ExcelValidationError(
                f"File size ({len(file_data) / 1024 / 1024:.1f}MB) exceeds "
                f"limit ({self.max_upload_size_bytes / 1024 / 1024:.0f}MB)",
                filename=filename
            )

        # Try to read the file
        try:
            df = pd.read_excel(BytesIO(file_data))
        except Exception as e:
            raise ExcelValidationError(
                f"Cannot read Excel file: {e}",
                filename=filename
            )

        # Check for empty file
        if df.empty:
            raise ExcelValidationError(
                "Excel file is empty",
                filename=filename
            )

        # Check row count
        if len(df) > self.max_rows:
            logger.warning(
                f"Excel file has {len(df)} rows, truncating to {self.max_rows}"
            )
            df = df.head(self.max_rows)

        logger.info(
            f"Read Excel file: {filename or 'unknown'}, "
            f"{len(df)} rows, {len(df.columns)} columns"
        )

        return df

    def validate_columns(
        self,
        df: pd.DataFrame,
        img_column: str,
        text_columns: List[str]
    ) -> Tuple[str, List[str]]:
        """
        Validate that required columns exist.

        Handles both letter-based (A, B, C) and name-based column references.

        Args:
            df: DataFrame to validate
            img_column: Image column reference
            text_columns: List of text column references

        Returns:
            Tuple of (resolved_img_column, resolved_text_columns)

        Raises:
            ExcelValidationError: If columns don't exist
        """
        available_columns = list(df.columns)

        # Resolve image column
        resolved_img = self._resolve_column(df, img_column, available_columns)
        if resolved_img is None:
            raise ExcelValidationError(
                f"Image column '{img_column}' not found",
                column=img_column,
                details=f"Available columns: {available_columns}"
            )

        # Resolve text columns
        resolved_text = []
        for col in text_columns:
            resolved = self._resolve_column(df, col, available_columns)
            if resolved is None:
                logger.warning(f"Text column '{col}' not found, skipping")
            else:
                resolved_text.append(resolved)

        if not resolved_text:
            raise ExcelValidationError(
                "No valid text columns found",
                details=f"Requested: {text_columns}, Available: {available_columns}"
            )

        return resolved_img, resolved_text

    def _resolve_column(
        self,
        df: pd.DataFrame,
        column_ref: str,
        available: List[str]
    ) -> Optional[str]:
        """
        Resolve a column reference to actual column name.

        Handles:
        - Direct name match
        - Letter-based reference (A, B, C, ...)
        - Index-based reference (0, 1, 2, ...)
        """
        column_ref = column_ref.strip()

        # Direct name match
        if column_ref in available:
            return column_ref

        # Case-insensitive name match
        for col in available:
            if str(col).lower() == column_ref.lower():
                return col

        # Letter-based reference (A=0, B=1, etc.)
        if column_ref.isalpha() and len(column_ref) <= 2:
            index = self._letter_to_index(column_ref.upper())
            if 0 <= index < len(available):
                return available[index]

        # Index-based reference
        if column_ref.isdigit():
            index = int(column_ref)
            if 0 <= index < len(available):
                return available[index]

        return None

    @staticmethod
    def _letter_to_index(letter: str) -> int:
        """Convert Excel-style column letter to index (A=0, B=1, AA=26, etc.)."""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1

    def get_slide_data(
        self,
        df: pd.DataFrame,
        img_column: str,
        text_columns: List[str],
        sanitize: bool = True,
        preserve_column_identity: bool = True
    ) -> List[dict]:
        """
        Extract and optionally sanitize slide data from DataFrame.

        Args:
            df: Source DataFrame
            img_column: Resolved image column name
            text_columns: Resolved text column names
            sanitize: Whether to sanitize text content
            preserve_column_identity: If True, text_content contains dicts with
                column info; if False, contains plain strings (backward compat)

        Returns:
            List of dicts with 'image_source', 'image_cell', and 'text_content' keys.
            text_content items are either strings or dicts with 'column' and 'text' keys.
        """
        slides = []

        # Get column index for cell reference
        col_index = list(df.columns).index(img_column) if img_column in df.columns else -1
        col_letter = self._get_column_letters(col_index + 1)[-1] if col_index >= 0 else "A"

        # Build column letter map for text columns
        all_columns = list(df.columns)
        text_col_letters = {}
        for col in text_columns:
            if col in all_columns:
                idx = all_columns.index(col)
                text_col_letters[col] = self._get_column_letters(idx + 1)[-1]

        for index, row in df.iterrows():
            # Create cell reference (e.g., "B2" for row 1 with 1-based indexing)
            row_num = index + 2  # +2 because Excel is 1-indexed and has header row
            cell_ref = f"{col_letter}{row_num}"

            slide_data = {
                "row_index": index,
                "image_source": None,
                "image_cell": cell_ref,
                "text_content": []
            }

            # Get image source (file path or other reference)
            if img_column in row and pd.notna(row[img_column]):
                source = str(row[img_column]).strip()
                if source:
                    slide_data["image_source"] = source

            # Get text content with column identity preserved
            for col in text_columns:
                if col in row and pd.notna(row[col]):
                    text = str(row[col])
                    if sanitize:
                        text = sanitize_text(text)
                    if text.strip():
                        if preserve_column_identity:
                            col_letter_ref = text_col_letters.get(col, col)
                            slide_data["text_content"].append({
                                "column": col_letter_ref,
                                "text": text
                            })
                        else:
                            slide_data["text_content"].append(text)

            slides.append(slide_data)

        logger.info(f"Extracted data for {len(slides)} slides")
        return slides

    def get_preview(
        self,
        df: pd.DataFrame,
        max_rows: int = 10
    ) -> pd.DataFrame:
        """
        Get a preview of the DataFrame for display.

        Args:
            df: Source DataFrame
            max_rows: Maximum rows to include in preview

        Returns:
            Preview DataFrame
        """
        return df.head(max_rows)

    def get_summary(self, df: pd.DataFrame) -> dict:
        """
        Get summary statistics about the DataFrame.

        Args:
            df: Source DataFrame

        Returns:
            Dict with summary info
        """
        return {
            "row_count": len(df),
            "column_count": len(df.columns),
            "columns": list(df.columns),
            "column_letters": self._get_column_letters(len(df.columns))
        }

    @staticmethod
    def _get_column_letters(count: int) -> List[str]:
        """Generate Excel-style column letters for N columns."""
        letters = []
        for i in range(count):
            letter = ""
            n = i
            while True:
                letter = chr(ord('A') + n % 26) + letter
                n = n // 26 - 1
                if n < 0:
                    break
            letters.append(letter)
        return letters


def read_excel_file(
    file_data: bytes,
    filename: Optional[str] = None
) -> pd.DataFrame:
    """
    Convenience function to read an Excel file.

    Args:
        file_data: Excel file content as bytes
        filename: Optional filename for error messages

    Returns:
        DataFrame with Excel data
    """
    processor = ExcelProcessor()
    return processor.read_excel(file_data, filename)


def parse_column_input(column_input: str) -> List[str]:
    """
    Parse comma-separated column input string.

    Args:
        column_input: String like "C,D,E,F" or "Title,Description"

    Returns:
        List of column references
    """
    return [col.strip() for col in column_input.split(",") if col.strip()]
