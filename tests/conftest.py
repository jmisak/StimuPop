"""
Pytest fixtures for StimuPop tests.
"""

import os
import sys
from io import BytesIO
from unittest.mock import MagicMock

import pandas as pd
import pytest
from PIL import Image

# Add src to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def sample_dataframe():
    """Create a sample DataFrame for testing."""
    return pd.DataFrame({
        'A': [1, 2, 3],
        'B': ['', '', ''],  # Image column (would contain embedded images or file paths)
        'C': ['Title 1', 'Title 2', 'Title 3'],
        'D': ['Description 1', 'Description 2', 'Description 3']
    })


@pytest.fixture
def sample_dataframe_with_names():
    """Create a DataFrame with named columns."""
    return pd.DataFrame({
        'ID': [1, 2, 3],
        'Image': ['', '', ''],  # Image column
        'Title': ['Product 1', 'Product 2', 'Product 3'],
        'Description': ['Desc 1', 'Desc 2', 'Desc 3'],
        'Price': ['$10', '$20', '$30']
    })


@pytest.fixture
def sample_excel_bytes(sample_dataframe):
    """Create sample Excel file as bytes."""
    buffer = BytesIO()
    sample_dataframe.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


@pytest.fixture
def slide_config():
    """Create a sample SlideConfig for testing."""
    from src.pptx_generator import SlideConfig
    return SlideConfig(
        img_column='B',
        text_columns=['C', 'D'],
        img_width=5.5,
        img_top=0.5,
        text_top=5.0,
        font_size=14
    )


@pytest.fixture
def mock_image_result():
    """Create a mock successful ImageResult."""
    from src.image_handler import ImageResult

    # Create a valid PNG image
    img = Image.new('RGB', (100, 100), color='blue')
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)

    return ImageResult(
        source='test.png',
        success=True,
        data=buffer,
        width=100,
        height=100,
        format='PNG',
        size_bytes=len(buffer.getvalue())
    )


@pytest.fixture
def mock_failed_image_result():
    """Create a mock failed ImageResult."""
    from src.image_handler import ImageResult

    return ImageResult(
        source='notfound.jpg',
        success=False,
        error='File not found'
    )


@pytest.fixture
def temp_config(tmp_path):
    """Create a temporary config file."""
    config_content = """
app:
  name: "Test App"
  version: "1.0.0"
  max_upload_size_mb: 50

images:
  max_size_mb: 5
  cache_ttl_seconds: 3600
"""
    config_file = tmp_path / "config.yaml"
    config_file.write_text(config_content)
    return str(config_file)
