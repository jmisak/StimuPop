"""
StimuPop - Excel to PowerPoint Converter

A production-grade web application that converts Excel data to
PowerPoint presentations with embedded images and formatted text.

Features:
- Embedded Excel image extraction
- Local file path image support
- Configurable slide layout
- Progress tracking
- Comprehensive error handling

Version: 2.2.0
"""

import tempfile
import os
from io import BytesIO

import streamlit as st
import pandas as pd

from src import (
    Config,
    get_config,
    ExcelProcessor,
    PPTXGenerator,
    SlideConfig,
    ColumnFormat,
    ImageLoader,
    ExcelValidationError,
    PPTXGenerationError,
)
from src.logging_config import setup_logging, request_context, get_logger
from src.excel_handler import parse_column_input

# Initialize logging
setup_logging()
logger = get_logger(__name__)


# Page configuration
st.set_page_config(
    page_title="StimuPop",
    page_icon="🎯",
    layout="wide"
)


def main():
    """Main application entry point."""
    with request_context() as req_id:
        logger.info(f"Session started: {req_id}")
        render_app()


def render_app():
    """Render the main application UI."""
    st.title("🎯 StimuPop")
    st.markdown("*Excel to PowerPoint with embedded images*")
    st.markdown("---")

    # Create two columns for layout
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📁 Upload Files")
        excel_file, template_file = render_file_uploaders()

    with col2:
        st.subheader("⚙️ Configuration")
        img_column, text_columns, font_size = render_basic_config()

    # Advanced settings
    img_width, img_top, text_top, orientation, column_formats = render_advanced_settings(text_columns, font_size)

    st.markdown("---")

    # Preview Excel data
    df = render_data_preview(excel_file)

    st.markdown("---")

    # Generate button
    render_generate_section(
        excel_file=excel_file,
        template_file=template_file,
        df=df,
        img_column=img_column,
        text_columns=text_columns,
        font_size=font_size,
        img_width=img_width,
        img_top=img_top,
        text_top=text_top,
        orientation=orientation,
        column_formats=column_formats
    )

    # Instructions
    render_instructions()

    # Footer
    render_footer()


def render_file_uploaders():
    """Render file upload widgets."""
    config = get_config()

    excel_file = st.file_uploader(
        "Upload Excel File (.xlsx)",
        type=['xlsx'],
        help=f"Upload your Excel file with embedded images (max {config.app.max_upload_size_mb}MB)"
    )

    template_file = st.file_uploader(
        "Upload PowerPoint Template (optional)",
        type=['pptx'],
        help="Upload a .pptx template for custom styling"
    )

    return excel_file, template_file


def render_basic_config():
    """Render basic configuration inputs."""
    img_column = st.text_input(
        "Image Column",
        "B",
        help="Column letter or name containing embedded images or file paths (e.g., 'B' or 'Image')"
    )

    text_columns = st.text_input(
        "Text Columns (comma-separated)",
        "C,D,E,F",
        help="Column letters or names for text content (e.g., 'C,D,E,F')"
    )

    font_size = st.slider(
        "Font Size (pt)",
        min_value=10,
        max_value=32,
        value=14,
        help="Font size for text content"
    )

    return img_column, text_columns, font_size


def render_advanced_settings(text_columns_str: str, default_font_size: int):
    """Render advanced settings in an expander."""
    column_formats = None

    with st.expander("🔧 Advanced Settings"):
        st.markdown("#### Layout")
        adv_col1, adv_col2 = st.columns(2)

        with adv_col1:
            img_width = st.slider(
                "Image Width (inches)",
                min_value=3.0,
                max_value=7.0,
                value=5.5,
                step=0.5,
                help="Width of the image on each slide"
            )

            img_top = st.slider(
                "Image Top Position (inches)",
                min_value=0.0,
                max_value=3.0,
                value=0.5,
                step=0.25,
                help="Distance from top of slide to image"
            )

        with adv_col2:
            text_top = st.slider(
                "Text Top Position (inches)",
                min_value=3.0,
                max_value=7.0,
                value=5.0,
                step=0.5,
                help="Distance from top of slide to text"
            )

            orientation = st.selectbox(
                "Slide Orientation",
                options=["portrait", "landscape"],
                index=0,
                help="Portrait (tall) or Landscape (wide) slides"
            )

        # Per-column formatting section
        st.markdown("---")
        st.markdown("#### Column Formatting")
        column_formats = render_column_format_config(text_columns_str, default_font_size)

    return img_width, img_top, text_top, orientation, column_formats


def render_column_format_config(text_columns_str: str, default_font_size: int) -> dict:
    """Render per-column font formatting controls."""
    columns = parse_column_input(text_columns_str)

    if not columns:
        st.caption("Enter text columns above to configure formatting")
        return None

    column_formats = {}

    # Use tabs for each column
    tabs = st.tabs([f"Column {col}" for col in columns])

    font_options = ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Tahoma"]

    for tab, col in zip(tabs, columns):
        with tab:
            fmt_col1, fmt_col2 = st.columns(2)

            with fmt_col1:
                font_size = st.slider(
                    "Font Size",
                    min_value=8,
                    max_value=48,
                    value=default_font_size,
                    key=f"size_{col}"
                )

                font_name = st.selectbox(
                    "Font",
                    options=font_options,
                    index=0,
                    key=f"font_{col}"
                )

            with fmt_col2:
                color = st.color_picker(
                    "Color",
                    value="#000000",
                    key=f"color_{col}"
                )

                bold = st.checkbox("Bold", key=f"bold_{col}")
                italic = st.checkbox("Italic", key=f"italic_{col}")

            column_formats[col] = ColumnFormat(
                column=col,
                font_size=font_size,
                bold=bold,
                italic=italic,
                font_name=font_name,
                color=color.lstrip("#")
            )

    return column_formats


@st.cache_data(show_spinner=False)
def load_excel_preview(file_bytes: bytes, filename: str):
    """Load Excel file with caching for preview."""
    processor = ExcelProcessor()
    return processor.read_excel(file_bytes, filename)


def render_data_preview(excel_file):
    """Render Excel data preview."""
    if not excel_file:
        return None

    st.subheader("📋 Data Preview")

    try:
        df = load_excel_preview(excel_file.getvalue(), excel_file.name)
        processor = ExcelProcessor()
        summary = processor.get_summary(df)

        # Display summary
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Rows:** {summary['row_count']}")
        with col2:
            st.write(f"**Columns:** {summary['column_count']}")

        # Column mapping info
        letter_mapping = ", ".join(
            f"{letter}={name}"
            for letter, name in zip(summary['column_letters'], summary['columns'])
        )
        st.caption(f"Column mapping: {letter_mapping}")

        # Data preview
        st.dataframe(processor.get_preview(df), use_container_width=True)

        return df

    except ExcelValidationError as e:
        st.error(f"❌ {e.message}")
        logger.warning(f"Excel validation error: {e}")
        return None
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        logger.error(f"Excel read error: {e}", exc_info=True)
        return None


def render_generate_section(
    excel_file,
    template_file,
    df,
    img_column,
    text_columns,
    font_size,
    img_width,
    img_top,
    text_top,
    orientation,
    column_formats
):
    """Render the generate button and handle generation."""
    if st.button("🎨 Generate Presentation", type="primary", use_container_width=True):
        if not excel_file:
            st.error("❌ Please upload an Excel file first!")
            return

        if df is None or len(df) == 0:
            st.error("❌ Excel file appears to be empty or invalid!")
            return

        generate_presentation(
            df=df,
            excel_file=excel_file,
            template_file=template_file,
            img_column=img_column,
            text_columns=text_columns,
            font_size=font_size,
            img_width=img_width,
            img_top=img_top,
            text_top=text_top,
            orientation=orientation,
            column_formats=column_formats
        )


def generate_presentation(
    df,
    excel_file,
    template_file,
    img_column,
    text_columns,
    font_size,
    img_width,
    img_top,
    text_top,
    orientation,
    column_formats
):
    """Generate the PowerPoint presentation."""
    logger.info(f"Starting presentation generation for {excel_file.name}")

    try:
        # Parse and validate columns
        processor = ExcelProcessor()
        text_cols = parse_column_input(text_columns)

        try:
            resolved_img, resolved_text = processor.validate_columns(
                df, img_column, text_cols
            )
        except ExcelValidationError as e:
            st.error(f"❌ {e.message}")
            if e.details:
                st.info(f"ℹ️ {e.details}")
            return

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Extract embedded images from Excel
        status_text.text("Extracting embedded images...")
        loader = ImageLoader()
        embedded_images = loader.extract_embedded_images(excel_file.getvalue())

        if embedded_images:
            st.info(f"📷 Found {len(embedded_images)} embedded images in Excel")

        # Extract slide data with column identity preserved
        slide_data = processor.get_slide_data(df, resolved_img, resolved_text, preserve_column_identity=True)

        # Create slide configuration with per-column formats
        config = SlideConfig(
            img_column=resolved_img,
            text_columns=resolved_text,
            img_width=img_width,
            img_top=img_top,
            text_top=text_top,
            font_size=font_size,
            orientation=orientation,
            column_formats=column_formats
        )

        # Get template bytes if provided
        template_bytes = template_file.getvalue() if template_file else None

        def progress_callback(status: str, current: int, total: int):
            if total > 0:
                progress_bar.progress(current / total)
            status_text.text(status)

        # Generate presentation
        generator = PPTXGenerator(config)
        result = generator.generate(
            slide_data,
            embedded_images=embedded_images,
            template_file=template_bytes,
            progress_callback=progress_callback
        )

        progress_bar.empty()
        status_text.empty()

        if not result.success:
            st.error(f"❌ Generation failed: {result.error}")
            logger.error(f"Generation failed: {result.error}")
            return

        # Show results summary
        if result.slides_with_errors > 0:
            st.warning(
                f"⚠️ Generated {result.slides_generated} slides, "
                f"but {result.slides_with_errors} had issues (images may be missing)"
            )

            # Show detailed errors in expander
            with st.expander("View slide issues"):
                for sr in result.slide_results:
                    if sr.image_error:
                        st.caption(f"Slide {sr.index + 1}: {sr.image_error}")
        else:
            st.success(
                f"✅ Generated {result.slides_generated} slides "
                f"with {result.slides_with_images} images"
            )

        # Save to temporary file and provide download
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
            result.presentation.save(tmp.name)
            tmp_path = tmp.name

        with open(tmp_path, 'rb') as f:
            pptx_data = f.read()

        os.unlink(tmp_path)

        # Download button
        output_filename = excel_file.name.rsplit('.', 1)[0] + '_presentation.pptx'
        st.download_button(
            label="📥 Download Presentation",
            data=pptx_data,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

        logger.info(f"Presentation generated successfully: {output_filename}")

    except Exception as e:
        st.error(f"❌ Error generating presentation: {str(e)}")
        logger.error(f"Generation error: {e}", exc_info=True)
        st.exception(e)


def render_instructions():
    """Render usage instructions."""
    st.markdown("---")
    with st.expander("📖 How to Use"):
        st.markdown("""
### Instructions

1. **Prepare your Excel file:**
   - Each row becomes one slide
   - Include images embedded in a column (e.g., column B) or file paths to local images
   - Include columns with text content (e.g., columns C, D, E, F)

2. **Upload your files:**
   - Upload the Excel file with embedded images (required)
   - Optionally upload a PowerPoint template

3. **Configure settings:**
   - Specify which column contains images
   - Specify which columns contain text
   - Adjust font size and positioning as needed

4. **Generate:**
   - Click "Generate Presentation"
   - Download your completed PowerPoint file

### Column Reference

You can reference columns by:
- **Letter**: A, B, C, ... (Excel-style)
- **Name**: The actual column header name

### Example Excel Structure

| A (Skip) | B (Image) | C (Title) | D (Description) |
|----------|-----------|-----------|-----------------|
| Row 1    | [embedded]| Title 1   | Description 1   |
| Row 2    | [embedded]| Title 2   | Description 2   |

### Tips

- Embed images directly in Excel cells for best results
- Or use local file paths (e.g., `C:\\Images\\photo.jpg`)
- Images are automatically centered and sized
- Missing images are skipped (slide still created)
- Text is centered below images
        """)


def render_footer():
    """Render footer."""
    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; color: gray;'>"
        "🎯 StimuPop v2.2.0 | "
        "Built with Streamlit"
        "</p>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
