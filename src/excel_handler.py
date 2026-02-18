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
from typing import Dict, List, Optional, Tuple

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

        # Resolve text columns (can be empty for Pictures Only mode)
        resolved_text = []
        for col in text_columns:
            resolved = self._resolve_column(df, col, available_columns)
            if resolved is None:
                logger.warning(f"Text column '{col}' not found, skipping")
            else:
                resolved_text.append(resolved)

        # Allow empty text columns for Pictures Only mode (NEW in v6.2)
        if not resolved_text and text_columns:
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
        preserve_column_identity: bool = True,
        text_separator: str = "",
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
            text_separator: If non-empty, join all text columns into a single entry
                using this string as separator (e.g., " for ").

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

            # Join text columns with separator if specified
            if text_separator and len(slide_data["text_content"]) > 1:
                items = slide_data["text_content"]
                if preserve_column_identity and all(isinstance(it, dict) for it in items):
                    combined = text_separator.join([it["text"] for it in items])
                    slide_data["text_content"] = [{"column": items[0]["column"], "text": combined}]
                elif not preserve_column_identity and all(isinstance(it, str) for it in items):
                    slide_data["text_content"] = [text_separator.join(items)]

            slides.append(slide_data)

        logger.info(f"Extracted data for {len(slides)} slides")
        return slides

    def validate_columns_multi(
        self,
        df: pd.DataFrame,
        image_elements: list,
        text_groups: list
    ) -> Tuple[List[Tuple[str, str]], List[Tuple[List[str], str]]]:
        """
        Validate columns for multi-element mode (NEW in v8.0).

        Args:
            df: DataFrame to validate
            image_elements: List of ImageElement objects (each has .column and .placeholder_name)
            text_groups: List of TextGroup objects (each has .columns and .placeholder_name)

        Returns:
            Tuple of (resolved_images, resolved_texts) where:
            - resolved_images: List of (resolved_column_name, placeholder_name)
            - resolved_texts: List of (resolved_column_names_list, placeholder_name)

        Raises:
            ExcelValidationError: If any required image column doesn't exist,
                or if ALL columns for a text group are missing.
        """
        available_columns = list(df.columns)

        # --- Resolve image element columns ---
        resolved_images: List[Tuple[str, str]] = []
        for elem in image_elements:
            resolved = self._resolve_column(df, elem.column, available_columns)
            if resolved is None:
                raise ExcelValidationError(
                    f"Image column '{elem.column}' for placeholder "
                    f"'{elem.placeholder_name}' not found",
                    column=elem.column,
                    details=f"Available columns: {available_columns}"
                )
            resolved_images.append((resolved, elem.placeholder_name))

        # --- Resolve text group columns ---
        resolved_texts: List[Tuple[List[str], str]] = []
        for group in text_groups:
            resolved_cols: List[str] = []
            for col in group.columns:
                resolved = self._resolve_column(df, col, available_columns)
                if resolved is None:
                    logger.warning(
                        f"Text column '{col}' for placeholder "
                        f"'{group.placeholder_name}' not found, skipping"
                    )
                else:
                    resolved_cols.append(resolved)

            if not resolved_cols:
                raise ExcelValidationError(
                    f"No valid text columns found for placeholder "
                    f"'{group.placeholder_name}'",
                    details=(
                        f"Requested: {group.columns}, "
                        f"Available: {available_columns}"
                    )
                )
            resolved_texts.append((resolved_cols, group.placeholder_name))

        logger.info(
            f"Multi-element validation passed: "
            f"{len(resolved_images)} image(s), {len(resolved_texts)} text group(s)"
        )
        return resolved_images, resolved_texts

    def get_slide_data_multi(
        self,
        df: pd.DataFrame,
        image_elements: list,
        text_groups: list,
        sanitize: bool = True
    ) -> List[dict]:
        """
        Extract slide data for multi-element mode (NEW in v8.0).

        Args:
            df: Source DataFrame
            image_elements: List of ImageElement objects (each has .column and .placeholder_name)
            text_groups: List of TextGroup objects (each has .columns and .placeholder_name)
            sanitize: Whether to sanitize text content

        Returns:
            List of dicts with:
            - 'row_index': int
            - 'image_sources': list of {image_source, image_cell, placeholder_name}
            - 'text_contents': list of {text_content: [...], placeholder_name}
            - Legacy fields: 'image_source', 'image_cell', 'text_content' (from first element)
        """
        all_columns = list(df.columns)

        # Pre-compute column letters for image elements
        img_col_meta: List[Dict[str, str]] = []
        for elem in image_elements:
            col_name = elem.column
            # Resolve to actual column name (may already be resolved by validate)
            resolved = self._resolve_column(df, col_name, all_columns)
            if resolved and resolved in all_columns:
                idx = all_columns.index(resolved)
                letter = self._get_column_letters(idx + 1)[-1]
            else:
                resolved = col_name
                letter = "A"
            img_col_meta.append({
                "resolved": resolved,
                "letter": letter,
                "placeholder_name": elem.placeholder_name
            })

        # Pre-compute column letters for text groups
        txt_col_meta: List[Dict] = []
        for group in text_groups:
            group_cols: List[Dict[str, str]] = []
            for col in group.columns:
                resolved = self._resolve_column(df, col, all_columns)
                if resolved and resolved in all_columns:
                    idx = all_columns.index(resolved)
                    letter = self._get_column_letters(idx + 1)[-1]
                    group_cols.append({"resolved": resolved, "letter": letter})
            txt_col_meta.append({
                "columns": group_cols,
                "placeholder_name": group.placeholder_name,
                "separator": getattr(group, 'separator', ''),
            })

        slides: List[dict] = []

        for index, row in df.iterrows():
            row_num = index + 2  # Excel is 1-indexed + header row

            # --- Build image_sources ---
            image_sources: List[Dict] = []
            for meta in img_col_meta:
                cell_ref = f"{meta['letter']}{row_num}"
                source = None
                col_name = meta["resolved"]
                if col_name in row and pd.notna(row[col_name]):
                    val = str(row[col_name]).strip()
                    if val:
                        source = val
                image_sources.append({
                    "image_source": source,
                    "image_cell": cell_ref,
                    "placeholder_name": meta["placeholder_name"]
                })

            # --- Build text_contents ---
            text_contents: List[Dict] = []
            for meta in txt_col_meta:
                texts: List[Dict[str, str]] = []
                for col_info in meta["columns"]:
                    col_name = col_info["resolved"]
                    if col_name in row and pd.notna(row[col_name]):
                        text = str(row[col_name])
                        if sanitize:
                            text = sanitize_text(text)
                        if text.strip():
                            texts.append({
                                "column": col_info["letter"],
                                "text": text
                            })
                # If separator is set, join all column texts into a single entry
                separator = meta.get("separator", "")
                if separator and len(texts) > 1:
                    combined = separator.join([t["text"] for t in texts])
                    texts = [{"column": texts[0]["column"], "text": combined}]

                text_contents.append({
                    "text_content": texts,
                    "placeholder_name": meta["placeholder_name"]
                })

            # --- Legacy backward-compat fields from first elements ---
            first_img = image_sources[0] if image_sources else {
                "image_source": None, "image_cell": f"A{row_num}"
            }
            first_txt = text_contents[0] if text_contents else {
                "text_content": []
            }

            slide_data = {
                "row_index": index,
                "image_sources": image_sources,
                "text_contents": text_contents,
                # Legacy fields
                "image_source": first_img["image_source"],
                "image_cell": first_img["image_cell"],
                "text_content": first_txt["text_content"]
            }
            slides.append(slide_data)

        logger.info(f"Extracted multi-element data for {len(slides)} slides")
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
