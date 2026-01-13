"""
Logging configuration for the Excel to PowerPoint Converter.

Provides structured logging with:
- File and console handlers
- Log rotation
- Request ID tracking
- Configurable log levels
"""

import logging
import sys
import uuid
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler
from pathlib import Path
from threading import local
from typing import Generator, Optional

from .config import get_config


# Thread-local storage for request ID
_thread_local = local()


def get_request_id() -> str:
    """Get the current request ID for this thread."""
    return getattr(_thread_local, "request_id", "no-request-id")


def set_request_id(request_id: str) -> None:
    """Set the request ID for this thread."""
    _thread_local.request_id = request_id


def generate_request_id() -> str:
    """Generate a new unique request ID."""
    return str(uuid.uuid4())[:8]


@contextmanager
def request_context(request_id: Optional[str] = None) -> Generator[str, None, None]:
    """
    Context manager for request ID tracking.

    Usage:
        with request_context() as req_id:
            logger.info("Processing request")
            # All logs in this context will include the request ID
    """
    old_id = getattr(_thread_local, "request_id", None)
    new_id = request_id or generate_request_id()
    set_request_id(new_id)
    try:
        yield new_id
    finally:
        if old_id is None:
            delattr(_thread_local, "request_id")
        else:
            _thread_local.request_id = old_id


class RequestIdFilter(logging.Filter):
    """Logging filter that adds request ID to log records."""

    def filter(self, record: logging.LogRecord) -> bool:
        record.request_id = get_request_id()
        return True


def setup_logging(
    level: Optional[str] = None,
    log_file: Optional[str] = None,
    console: bool = True
) -> logging.Logger:
    """
    Set up application logging.

    Configures the root logger for the application with:
    - Rotating file handler (if log_file specified)
    - Console handler (if console=True)
    - Request ID tracking

    Args:
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL).
               If None, uses config value.
        log_file: Path to log file. If None, uses config value.
        console: Whether to log to console.

    Returns:
        Configured logger instance.
    """
    config = get_config()

    # Determine log level
    if level is None:
        level = config.logging.level
    numeric_level = getattr(logging, level.upper(), logging.INFO)

    # Get or create the application logger
    logger = logging.getLogger("excel_to_pptx")
    logger.setLevel(numeric_level)

    # Clear existing handlers
    logger.handlers.clear()

    # Create formatter with request ID
    log_format = config.logging.format
    if "%(request_id)s" not in log_format:
        # Add request ID to format
        log_format = log_format.replace(
            "%(message)s",
            "[%(request_id)s] %(message)s"
        )
    formatter = logging.Formatter(log_format)

    # Add request ID filter
    request_filter = RequestIdFilter()

    # File handler with rotation
    if log_file is None:
        log_file = config.logging.file

    if log_file:
        log_path = Path(log_file)
        # Create log directory if needed
        log_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=config.logging.max_bytes,
            backupCount=config.logging.backup_count,
            encoding="utf-8"
        )
        file_handler.setLevel(numeric_level)
        file_handler.setFormatter(formatter)
        file_handler.addFilter(request_filter)
        logger.addHandler(file_handler)

    # Console handler
    if console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(numeric_level)
        console_handler.setFormatter(formatter)
        console_handler.addFilter(request_filter)
        logger.addHandler(console_handler)

    # Prevent propagation to root logger
    logger.propagate = False

    return logger


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a specific module.

    Creates a child logger of the application logger, inheriting
    its configuration.

    Args:
        name: Logger name (typically __name__)

    Returns:
        Logger instance

    Usage:
        logger = get_logger(__name__)
        logger.info("Something happened")
    """
    # Ensure base logging is configured
    base_logger = logging.getLogger("excel_to_pptx")
    if not base_logger.handlers:
        setup_logging()

    # Return child logger
    if name.startswith("src."):
        name = name[4:]  # Remove 'src.' prefix
    return logging.getLogger(f"excel_to_pptx.{name}")


class LogContext:
    """
    Context manager for structured logging with additional context.

    Usage:
        with LogContext(logger, operation="download_image", url=url):
            # Operations logged here will include the context
            result = download(url)
    """

    def __init__(self, logger: logging.Logger, **context):
        self.logger = logger
        self.context = context
        self.old_factory = None

    def __enter__(self):
        # Store old factory
        self.old_factory = logging.getLogRecordFactory()

        # Create new factory that adds context
        context = self.context

        def record_factory(*args, **kwargs):
            record = self.old_factory(*args, **kwargs)
            for key, value in context.items():
                setattr(record, key, value)
            return record

        logging.setLogRecordFactory(record_factory)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Restore old factory
        logging.setLogRecordFactory(self.old_factory)
        return False
