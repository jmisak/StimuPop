"""
Custom exception hierarchy for StimuPop.

Provides specific exception types for different error categories,
enabling targeted error handling and informative error messages.
"""

from typing import Optional


class AppError(Exception):
    """
    Base exception for all application errors.

    All custom exceptions inherit from this class, allowing for
    broad exception catching when needed.

    Attributes:
        message: Human-readable error description
        details: Optional additional context for debugging
    """

    def __init__(self, message: str, details: Optional[str] = None):
        self.message = message
        self.details = details
        super().__init__(self.message)

    def __str__(self) -> str:
        if self.details:
            return f"{self.message} | Details: {self.details}"
        return self.message


class ValidationError(AppError):
    """
    Raised when input validation fails.

    This includes:
    - Invalid file format
    - Invalid data types
    - Missing required fields

    Attributes:
        field: The field that failed validation
        reason: Specific reason for validation failure
    """

    def __init__(self, reason: str, field: Optional[str] = None, details: Optional[str] = None):
        self.field = field
        self.reason = reason
        if field:
            message = f"Validation failed for '{field}': {reason}"
        else:
            message = f"Validation failed: {reason}"
        super().__init__(message, details)


class ImageDownloadError(AppError):
    """
    Raised when image download fails.

    This includes:
    - Network errors
    - Timeout errors
    - Invalid image format
    - File size exceeded
    - HTTP errors (4xx, 5xx)

    Attributes:
        url: The URL that failed to download
        status_code: HTTP status code if available
        is_retryable: Whether the error may be transient
    """

    def __init__(
        self,
        url: str,
        reason: str,
        status_code: Optional[int] = None,
        is_retryable: bool = False,
        details: Optional[str] = None
    ):
        self.url = url
        self.status_code = status_code
        self.is_retryable = is_retryable
        message = f"Image download failed for '{url}': {reason}"
        if status_code:
            message += f" (HTTP {status_code})"
        super().__init__(message, details)


class ExcelValidationError(AppError):
    """
    Raised when Excel file validation fails.

    This includes:
    - Missing required columns
    - Invalid data types
    - Empty file
    - Corrupted file
    - Row count exceeded

    Attributes:
        filename: Name of the Excel file
        row: Optional row number where error occurred
        column: Optional column name where error occurred
    """

    def __init__(
        self,
        reason: str,
        filename: Optional[str] = None,
        row: Optional[int] = None,
        column: Optional[str] = None,
        details: Optional[str] = None
    ):
        self.filename = filename
        self.row = row
        self.column = column

        location_parts = []
        if filename:
            location_parts.append(f"file '{filename}'")
        if row is not None:
            location_parts.append(f"row {row}")
        if column:
            location_parts.append(f"column '{column}'")

        if location_parts:
            location = " in " + ", ".join(location_parts)
        else:
            location = ""

        message = f"Excel validation error{location}: {reason}"
        super().__init__(message, details)


class PPTXGenerationError(AppError):
    """
    Raised when PowerPoint generation fails.

    This includes:
    - Template loading errors
    - Slide creation errors
    - Invalid layout
    - Save/write errors

    Attributes:
        slide_number: Optional slide number where error occurred
        operation: The operation that failed (e.g., 'add_picture', 'save')
    """

    def __init__(
        self,
        reason: str,
        slide_number: Optional[int] = None,
        operation: Optional[str] = None,
        details: Optional[str] = None
    ):
        self.slide_number = slide_number
        self.operation = operation

        context_parts = []
        if slide_number is not None:
            context_parts.append(f"slide {slide_number}")
        if operation:
            context_parts.append(f"operation '{operation}'")

        if context_parts:
            context = " during " + ", ".join(context_parts)
        else:
            context = ""

        message = f"PowerPoint generation error{context}: {reason}"
        super().__init__(message, details)


class ConfigurationError(AppError):
    """
    Raised when configuration is invalid or missing.

    This includes:
    - Missing config file
    - Invalid config values
    - Missing required settings
    - Type mismatches

    Attributes:
        setting: The configuration setting that caused the error
    """

    def __init__(
        self,
        reason: str,
        setting: Optional[str] = None,
        details: Optional[str] = None
    ):
        self.setting = setting

        if setting:
            message = f"Configuration error for '{setting}': {reason}"
        else:
            message = f"Configuration error: {reason}"

        super().__init__(message, details)
