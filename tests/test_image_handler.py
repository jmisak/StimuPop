"""
Tests for image handling.
"""

import pytest
import time
import tempfile
from io import BytesIO
from pathlib import Path
from unittest.mock import patch, MagicMock

from PIL import Image

from src.image_handler import (
    ImageLoader,
    ImageCache,
    ImageResult,
    load_image,
    extract_excel_images,
)


class TestImageCache:
    """Tests for ImageCache."""

    def test_put_and_get(self):
        """Test basic put and get operations."""
        cache = ImageCache(ttl_seconds=3600)
        cache.put("test_path.jpg", b"data1", 100, 100, "JPEG")

        entry = cache.get("test_path.jpg")
        assert entry is not None
        assert entry.data == b"data1"
        assert entry.width == 100
        assert entry.height == 100

    def test_get_missing(self):
        """Test getting non-existent entry."""
        cache = ImageCache()
        assert cache.get("missing.jpg") is None

    def test_ttl_expiration(self):
        """Test TTL expiration."""
        cache = ImageCache(ttl_seconds=0.01)  # Very short expiration
        cache.put("test.jpg", b"data", 10, 10, "JPEG")

        # Wait for expiration
        time.sleep(0.02)

        # Should be expired
        assert cache.get("test.jpg") is None

    def test_max_entries(self):
        """Test max entries limit."""
        cache = ImageCache(ttl_seconds=3600, max_entries=2)

        cache.put("1.jpg", b"data1", 10, 10, "JPEG")
        cache.put("2.jpg", b"data2", 10, 10, "JPEG")
        cache.put("3.jpg", b"data3", 10, 10, "JPEG")

        # First entry should be evicted
        assert cache.get("1.jpg") is None
        assert cache.get("2.jpg") is not None
        assert cache.get("3.jpg") is not None

    def test_clear(self):
        """Test clearing cache."""
        cache = ImageCache()
        cache.put("test.jpg", b"data", 10, 10, "JPEG")
        cache.clear()
        assert cache.get("test.jpg") is None

    def test_cleanup_expired(self):
        """Test cleanup of expired entries."""
        cache = ImageCache(ttl_seconds=0.01)
        cache.put("test.jpg", b"data", 10, 10, "JPEG")

        # Wait for expiration
        time.sleep(0.02)

        removed = cache.cleanup_expired()
        assert removed == 1


class TestImageLoader:
    """Tests for ImageLoader."""

    def test_load_from_path_not_found(self):
        """Test loading non-existent file."""
        loader = ImageLoader()
        result = loader.load_from_path("nonexistent_file.jpg")
        assert result.success is False
        assert "not found" in result.error.lower()

    def test_load_from_path_invalid_format(self, tmp_path):
        """Test loading file with invalid extension."""
        # Create a file with invalid extension
        invalid_file = tmp_path / "test.pdf"
        invalid_file.write_text("not an image")

        loader = ImageLoader()
        result = loader.load_from_path(str(invalid_file))
        assert result.success is False
        assert "format" in result.error.lower()

    def test_load_from_path_success(self, tmp_path):
        """Test successful image loading from path."""
        # Create a valid image file
        img = Image.new('RGB', (100, 100), color='red')
        img_path = tmp_path / "test.png"
        img.save(str(img_path), format='PNG')

        loader = ImageLoader()
        result = loader.load_from_path(str(img_path))

        assert result.success is True
        assert result.width == 100
        assert result.height == 100
        assert result.format == "PNG"
        assert result.data is not None

    def test_load_from_path_uses_cache(self, tmp_path):
        """Test that cache is used for repeated loads."""
        # Create a valid image file
        img = Image.new('RGB', (50, 50), color='blue')
        img_path = tmp_path / "cached.png"
        img.save(str(img_path), format='PNG')

        loader = ImageLoader(use_cache=True)

        # First load
        result1 = loader.load_from_path(str(img_path))
        assert result1.success is True
        assert result1.from_cache is False

        # Second load should be from cache
        result2 = loader.load_from_path(str(img_path))
        assert result2.success is True
        assert result2.from_cache is True

    def test_load_from_path_relative(self, tmp_path):
        """Test loading with relative path."""
        # Create a valid image file
        img = Image.new('RGB', (30, 30), color='green')
        img_path = tmp_path / "relative.png"
        img.save(str(img_path), format='PNG')

        loader = ImageLoader(base_path=str(tmp_path))
        result = loader.load_from_path("relative.png")

        assert result.success is True
        assert result.width == 30
        assert result.height == 30

    def test_load_from_path_size_limit(self, tmp_path):
        """Test size limit enforcement."""
        # Create a large image
        img = Image.new('RGB', (1000, 1000), color='red')
        img_path = tmp_path / "large.png"
        img.save(str(img_path), format='PNG')

        # Very small size limit
        loader = ImageLoader(max_size_mb=0.001)  # ~1KB
        result = loader.load_from_path(str(img_path))

        assert result.success is False
        assert "size" in result.error.lower()

    def test_load_from_bytes_success(self):
        """Test loading image from bytes."""
        # Create image data
        img = Image.new('RGB', (20, 20), color='yellow')
        buffer = BytesIO()
        img.save(buffer, format='PNG')
        img_bytes = buffer.getvalue()

        loader = ImageLoader()
        result = loader.load_from_bytes(img_bytes, "test_image")

        assert result.success is True
        assert result.width == 20
        assert result.height == 20

    def test_load_from_bytes_invalid(self):
        """Test loading invalid bytes."""
        loader = ImageLoader()
        result = loader.load_from_bytes(b"not an image", "invalid")

        assert result.success is False
        assert "invalid" in result.error.lower()


class TestConvenienceFunctions:
    """Tests for convenience functions."""

    def test_load_image(self, tmp_path):
        """Test load_image function."""
        # Create a valid image file
        img = Image.new('RGB', (50, 50), color='purple')
        img_path = tmp_path / "func_test.png"
        img.save(str(img_path), format='PNG')

        result = load_image(str(img_path))
        assert isinstance(result, ImageResult)
        assert result.success is True

    def test_load_image_not_found(self):
        """Test load_image with non-existent file."""
        result = load_image("nonexistent.jpg")
        assert isinstance(result, ImageResult)
        assert result.success is False


class TestEmbeddedImageExtraction:
    """Tests for embedded Excel image extraction."""

    def test_extract_from_invalid_excel(self):
        """Test extracting from invalid Excel data."""
        loader = ImageLoader()
        results = loader.extract_embedded_images(b"not excel data")
        assert results == {}

    # Note: Testing actual embedded image extraction would require
    # creating a test Excel file with embedded images, which is complex.
    # In production, this would be tested with integration tests.
