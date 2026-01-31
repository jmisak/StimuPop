"""
Image handling for StimuPop.

Provides image loading from:
- Local file paths (absolute or relative)
- Embedded images from Excel files (via openpyxl)

Features:
- File validation and size limits
- Format validation
- In-memory caching
"""

import hashlib
import os
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional, Callable, Any

from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from .config import get_config
from .exceptions import ImageDownloadError
from .logging_config import get_logger

logger = get_logger(__name__)


# Allowed image extensions
ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp'}


@dataclass
class ImageResult:
    """Result of an image load operation."""
    source: str  # File path or cell reference
    success: bool
    data: Optional[BytesIO] = None
    error: Optional[str] = None
    width: Optional[int] = None
    height: Optional[int] = None
    format: Optional[str] = None
    size_bytes: Optional[int] = None
    from_cache: bool = False


@dataclass
class CacheEntry:
    """Cache entry for loaded images."""
    data: bytes
    timestamp: float
    width: int
    height: int
    format: str


class ImageCache:
    """
    Thread-safe in-memory cache for loaded images.

    Features:
    - TTL-based expiration
    - Path hash-based keys
    - Automatic cleanup of expired entries
    """

    def __init__(self, ttl_seconds: int = 3600, max_entries: int = 100):
        self._cache: Dict[str, CacheEntry] = {}
        self._lock = Lock()
        self._ttl = ttl_seconds
        self._max_entries = max_entries

    def _hash_key(self, key: str) -> str:
        """Create a hash key."""
        return hashlib.md5(key.encode()).hexdigest()

    def get(self, key: str) -> Optional[CacheEntry]:
        """Get cached image data. Returns None if not in cache or expired."""
        hashed = self._hash_key(key)
        with self._lock:
            entry = self._cache.get(hashed)
            if entry is None:
                return None

            if time.time() - entry.timestamp > self._ttl:
                del self._cache[hashed]
                return None

            return entry

    def put(self, key: str, data: bytes, width: int, height: int, fmt: str) -> None:
        """Cache image data."""
        hashed = self._hash_key(key)
        entry = CacheEntry(
            data=data,
            timestamp=time.time(),
            width=width,
            height=height,
            format=fmt
        )

        with self._lock:
            if len(self._cache) >= self._max_entries:
                oldest_key = min(
                    self._cache.keys(),
                    key=lambda k: self._cache[k].timestamp
                )
                del self._cache[oldest_key]

            self._cache[hashed] = entry

    def clear(self) -> None:
        """Clear all cached entries."""
        with self._lock:
            self._cache.clear()

    def cleanup_expired(self) -> int:
        """Remove expired entries. Returns count of removed entries."""
        now = time.time()
        removed = 0
        with self._lock:
            expired_keys = [
                k for k, v in self._cache.items()
                if now - v.timestamp > self._ttl
            ]
            for key in expired_keys:
                del self._cache[key]
                removed += 1
        return removed


# Global cache instance
_image_cache: Optional[ImageCache] = None


def get_image_cache() -> ImageCache:
    """Get the global image cache instance."""
    global _image_cache
    if _image_cache is None:
        config = get_config()
        _image_cache = ImageCache(ttl_seconds=config.images.cache_ttl_seconds)
    return _image_cache


class ImageLoader:
    """
    Loads images from local files or embedded Excel images.

    Features:
    - Local file path support (absolute and relative)
    - Size limits
    - Format validation
    - In-memory caching

    Usage:
        loader = ImageLoader()

        # Load from file path
        result = loader.load_from_path("/path/to/image.jpg")

        # Load embedded images from Excel
        embedded = loader.extract_embedded_images(excel_bytes)
    """

    def __init__(
        self,
        max_size_mb: Optional[int] = None,
        use_cache: bool = True,
        base_path: Optional[str] = None,
    ):
        """
        Initialize image loader.

        Args:
            max_size_mb: Maximum image size in MB
            use_cache: Whether to use image cache
            base_path: Base path for resolving relative file paths
        """
        config = get_config()
        img_config = config.images

        self.max_size_bytes = (
            max_size_mb * 1024 * 1024 if max_size_mb
            else img_config.max_size_bytes
        )
        self.use_cache = use_cache
        self.base_path = Path(base_path) if base_path else Path.cwd()
        self._cache = get_image_cache() if use_cache else None

    def load_from_path(self, file_path: str) -> ImageResult:
        """
        Load an image from a local file path.

        Args:
            file_path: Path to image file (absolute or relative)

        Returns:
            ImageResult with success status and data or error
        """
        # Check cache first
        if self._cache:
            cached = self._cache.get(file_path)
            if cached:
                logger.debug(f"Cache hit for {file_path}")
                return ImageResult(
                    source=file_path,
                    success=True,
                    data=BytesIO(cached.data),
                    width=cached.width,
                    height=cached.height,
                    format=cached.format,
                    size_bytes=len(cached.data),
                    from_cache=True
                )

        try:
            # Resolve path
            path = Path(file_path)
            if not path.is_absolute():
                path = self.base_path / path

            # Check file exists
            if not path.exists():
                return ImageResult(
                    source=file_path,
                    success=False,
                    error=f"File not found: {file_path}"
                )

            # Check extension
            ext = path.suffix.lower()
            if ext not in ALLOWED_EXTENSIONS:
                return ImageResult(
                    source=file_path,
                    success=False,
                    error=f"Invalid image format: {ext}. Allowed: {', '.join(ALLOWED_EXTENSIONS)}"
                )

            # Check file size
            file_size = path.stat().st_size
            if file_size > self.max_size_bytes:
                return ImageResult(
                    source=file_path,
                    success=False,
                    error=f"Image size ({file_size / 1024 / 1024:.1f}MB) exceeds limit ({self.max_size_bytes / 1024 / 1024:.0f}MB)"
                )

            # Read and validate image
            with open(path, 'rb') as f:
                image_data = f.read()

            return self._validate_and_create_result(file_path, image_data)

        except Exception as e:
            logger.error(f"Error loading image {file_path}: {e}")
            return ImageResult(
                source=file_path,
                success=False,
                error=str(e)
            )

    def load_from_bytes(self, image_bytes: bytes, source_name: str = "embedded") -> ImageResult:
        """
        Load an image from raw bytes.

        Args:
            image_bytes: Raw image data
            source_name: Name/identifier for the image source

        Returns:
            ImageResult with success status and data or error
        """
        # Check size
        if len(image_bytes) > self.max_size_bytes:
            return ImageResult(
                source=source_name,
                success=False,
                error=f"Image size ({len(image_bytes) / 1024 / 1024:.1f}MB) exceeds limit"
            )

        return self._validate_and_create_result(source_name, image_bytes)

    def _validate_and_create_result(self, source: str, image_data: bytes) -> ImageResult:
        """Validate image data and create result."""
        try:
            img_buffer = BytesIO(image_data)
            with Image.open(img_buffer) as img:
                width, height = img.size
                img_format = img.format or "unknown"
                # Verify it's a real image
                img.verify()

            # Cache the result
            if self._cache:
                self._cache.put(source, image_data, width, height, img_format)

            logger.info(f"Loaded image: {source} ({width}x{height}, {len(image_data) / 1024:.1f}KB)")

            return ImageResult(
                source=source,
                success=True,
                data=BytesIO(image_data),
                width=width,
                height=height,
                format=img_format,
                size_bytes=len(image_data)
            )

        except Exception as e:
            return ImageResult(
                source=source,
                success=False,
                error=f"Invalid image data: {e}"
            )

    def extract_embedded_images(self, excel_bytes: bytes) -> Dict[str, ImageResult]:
        """
        Extract embedded images from an Excel file.

        Images are keyed by their anchor cell (e.g., "B2", "C5").

        Supports:
        - Traditional embedded images (openpyxl _images)
        - Rich Data images (Excel 365 pasted images)

        Args:
            excel_bytes: Excel file content as bytes

        Returns:
            Dict mapping cell reference to ImageResult
        """
        results: Dict[str, ImageResult] = {}

        wb = None
        try:
            wb = load_workbook(BytesIO(excel_bytes))
            sheet = wb.active

            if sheet is None:
                return results

            # Method 1: Traditional embedded images via openpyxl
            for image in sheet._images:
                try:
                    anchor = image.anchor
                    if hasattr(anchor, '_from'):
                        col = anchor._from.col
                        row = anchor._from.row
                        col_letter = self._col_num_to_letter(col)
                        cell_ref = f"{col_letter}{row + 1}"
                    else:
                        cell_ref = f"image_{len(results)}"

                    img_data = image._data()
                    result = self.load_from_bytes(img_data, f"embedded:{cell_ref}")
                    results[cell_ref] = result

                except Exception as e:
                    logger.warning(f"Failed to extract traditional embedded image: {e}")
                    continue

            # Method 2: Rich Data images (Excel 365 pasted images)
            # Only try if no traditional images found
            if len(results) == 0:
                rich_data_results = self._extract_rich_data_images(excel_bytes)
                results.update(rich_data_results)

            logger.info(f"Extracted {len(results)} embedded images from Excel")

        except Exception as e:
            logger.error(f"Error extracting embedded images: {e}")
        finally:
            if wb is not None:
                wb.close()

        return results

    def _extract_rich_data_images(self, excel_bytes: bytes) -> Dict[str, ImageResult]:
        """
        Extract Rich Data images from Excel 365 files.

        These are images pasted directly into cells, stored in xl/richData/.

        Args:
            excel_bytes: Excel file content as bytes

        Returns:
            Dict mapping cell reference to ImageResult
        """
        import zipfile
        import re

        results: Dict[str, ImageResult] = {}

        try:
            with zipfile.ZipFile(BytesIO(excel_bytes), 'r') as z:
                # Check if richData exists
                file_list = z.namelist()
                has_rich_data = any('richData' in f for f in file_list)

                if not has_rich_data:
                    return results

                # Step 1: Get cell-to-vm mapping from sheet XML
                # Find cells with vm attribute (value metadata index)
                cell_to_vm = {}
                try:
                    sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
                    # Pattern: <c r="B2" ... vm="1">
                    pattern = r'<c r="([A-Z]+)(\d+)"[^>]*vm="(\d+)"'
                    for match in re.finditer(pattern, sheet_xml):
                        col_letter = match.group(1)
                        row_num = match.group(2)
                        vm_index = int(match.group(3))
                        cell_ref = f"{col_letter}{row_num}"
                        cell_to_vm[cell_ref] = vm_index
                except Exception as e:
                    logger.warning(f"Could not parse sheet XML for vm attributes: {e}")
                    return results

                if not cell_to_vm:
                    return results

                # Step 2: Get vm-to-image mapping from richValueRel relationships
                vm_to_image = {}
                try:
                    rels_xml = z.read('xl/richData/_rels/richValueRel.xml.rels').decode('utf-8')
                    # Pattern: <Relationship Id="rId1" ... Target="../media/image1.png"/>
                    pattern = r'Id="rId(\d+)"[^>]*Target="([^"]+)"'
                    for match in re.finditer(pattern, rels_xml):
                        rid_num = int(match.group(1))
                        target = match.group(2)
                        # vm index = rId number (vm=1 -> rId1 -> image)
                        vm_to_image[rid_num] = target.replace('../', 'xl/')
                except Exception as e:
                    logger.warning(f"Could not parse richValueRel relationships: {e}")
                    return results

                # Step 3: Extract images and map to cells
                for cell_ref, vm_index in cell_to_vm.items():
                    if vm_index in vm_to_image:
                        image_path = vm_to_image[vm_index]
                        try:
                            img_data = z.read(image_path)
                            result = self.load_from_bytes(img_data, f"richdata:{cell_ref}")
                            results[cell_ref] = result
                            logger.debug(f"Extracted Rich Data image for {cell_ref}: {image_path}")
                        except Exception as e:
                            logger.warning(f"Could not extract image {image_path} for {cell_ref}: {e}")

                logger.info(f"Extracted {len(results)} Rich Data images from Excel")

        except Exception as e:
            logger.error(f"Error extracting Rich Data images: {e}")

        return results

    @staticmethod
    def _col_num_to_letter(col_num: int) -> str:
        """Convert column number (0-based) to Excel letter."""
        result = ""
        col_num += 1  # Convert to 1-based
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result


def load_image(file_path: str, base_path: Optional[str] = None) -> ImageResult:
    """
    Convenience function to load an image from a file path.

    Args:
        file_path: Path to image file
        base_path: Optional base path for relative paths

    Returns:
        ImageResult
    """
    loader = ImageLoader(base_path=base_path)
    return loader.load_from_path(file_path)


def extract_excel_images(excel_bytes: bytes) -> Dict[str, ImageResult]:
    """
    Convenience function to extract embedded images from Excel.

    Args:
        excel_bytes: Excel file content as bytes

    Returns:
        Dict mapping cell reference to ImageResult
    """
    loader = ImageLoader()
    return loader.extract_embedded_images(excel_bytes)
