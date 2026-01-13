"""
PowerPoint generation for StimuPop.

Provides presentation creation with:
- Template support
- Embedded image support
- Local file path support
- Configurable layout
- Progress callbacks
- Error handling per slide
"""

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Callable, Dict, List, Optional, Union

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from .config import get_config
from .exceptions import PPTXGenerationError
from .image_handler import ImageLoader, ImageResult
from .logging_config import get_logger

logger = get_logger(__name__)


@dataclass
class ColumnFormat:
    """
    Font formatting for a single text column.

    Attributes:
        column: Column identifier (letter or name)
        font_size: Font size in points
        bold: Whether text is bold
        italic: Whether text is italic
        font_name: Font family name
        color: Hex color without # (e.g., "000000" for black)
    """
    column: str
    font_size: int = 14
    bold: bool = False
    italic: bool = False
    font_name: str = "Calibri"
    color: str = "000000"

    def get_rgb_color(self) -> RGBColor:
        """Convert hex color to RGBColor."""
        r = int(self.color[0:2], 16)
        g = int(self.color[2:4], 16)
        b = int(self.color[4:6], 16)
        return RGBColor(r, g, b)


@dataclass
class SlideConfig:
    """
    Configuration for slide generation.

    Attributes:
        img_column: Column containing images (embedded or file paths)
        text_columns: Columns containing text content
        img_width: Image width in inches
        img_top: Image top position in inches
        text_top: Text top position in inches
        font_size: Font size in points (fallback default)
        orientation: 'portrait' or 'landscape'
        column_formats: Per-column font formatting
    """
    img_column: str
    text_columns: List[str]
    img_width: float = 5.5
    img_top: float = 0.5
    text_top: float = 5.0
    font_size: int = 14
    orientation: str = "portrait"
    column_formats: Optional[Dict[str, ColumnFormat]] = None

    def get_column_format(self, column: str) -> ColumnFormat:
        """Get format for column, falling back to defaults."""
        if self.column_formats and column in self.column_formats:
            return self.column_formats[column]
        return ColumnFormat(column=column, font_size=self.font_size)


@dataclass
class SlideResult:
    """Result of generating a single slide."""
    index: int
    success: bool
    has_image: bool = False
    image_error: Optional[str] = None
    text_added: bool = False
    error: Optional[str] = None


@dataclass
class GenerationResult:
    """Result of generating a presentation."""
    success: bool
    presentation: Optional[Presentation] = None
    slides_generated: int = 0
    slides_with_images: int = 0
    slides_with_errors: int = 0
    slide_results: List[SlideResult] = None
    error: Optional[str] = None

    def __post_init__(self):
        if self.slide_results is None:
            self.slide_results = []


class PPTXGenerator:
    """
    Generates PowerPoint presentations from slide data.

    Features:
    - Template support
    - Embedded Excel images
    - Local file path images
    - Configurable slide layout
    - Progress callbacks
    - Per-slide error handling (continues on error)

    Usage:
        generator = PPTXGenerator(config)
        result = generator.generate(
            slide_data,
            embedded_images=embedded_dict,
            progress_callback=my_callback
        )
        if result.success:
            result.presentation.save("output.pptx")
    """

    def __init__(
        self,
        config: Optional[SlideConfig] = None,
        image_loader: Optional[ImageLoader] = None
    ):
        """
        Initialize generator.

        Args:
            config: Slide configuration
            image_loader: Image loader instance
        """
        app_config = get_config()
        pres_config = app_config.presentation

        if config is None:
            config = SlideConfig(
                img_column="B",
                text_columns=["C", "D", "E", "F"],
                img_width=pres_config.default_img_width,
                img_top=pres_config.default_img_top,
                text_top=pres_config.default_text_top,
                font_size=pres_config.default_font_size,
                orientation=pres_config.default_orientation
            )

        self.config = config
        self.loader = image_loader or ImageLoader()
        self.app_config = app_config

    def generate(
        self,
        slide_data: List[dict],
        embedded_images: Optional[Dict[str, ImageResult]] = None,
        template_file: Optional[Union[bytes, BytesIO, str, Path]] = None,
        progress_callback: Optional[Callable[[str, int, int], None]] = None
    ) -> GenerationResult:
        """
        Generate a PowerPoint presentation.

        Args:
            slide_data: List of dicts with 'image_source' and 'text_content' keys
            embedded_images: Dict of embedded images keyed by cell reference
            template_file: Optional template file (bytes, BytesIO, or path)
            progress_callback: Optional callback(status, current, total)

        Returns:
            GenerationResult with presentation and statistics
        """
        if not slide_data:
            return GenerationResult(
                success=False,
                error="No slide data provided"
            )

        embedded_images = embedded_images or {}

        try:
            # Create or load presentation
            prs = self._create_presentation(template_file)

            total_slides = len(slide_data)

            if progress_callback:
                progress_callback("Creating slides...", 0, total_slides)

            # Generate slides
            slide_results = []
            slides_with_images = 0
            slides_with_errors = 0

            for i, data in enumerate(slide_data):
                if progress_callback:
                    progress_callback(
                        f"Creating slide {i + 1}/{total_slides}",
                        i + 1,
                        total_slides
                    )

                result = self._create_slide(prs, data, embedded_images)
                slide_results.append(result)

                if result.has_image:
                    slides_with_images += 1
                if result.image_error or result.error:
                    slides_with_errors += 1

            logger.info(
                f"Generated presentation: {total_slides} slides, "
                f"{slides_with_images} with images, "
                f"{slides_with_errors} with errors"
            )

            return GenerationResult(
                success=True,
                presentation=prs,
                slides_generated=total_slides,
                slides_with_images=slides_with_images,
                slides_with_errors=slides_with_errors,
                slide_results=slide_results
            )

        except Exception as e:
            logger.error(f"Presentation generation failed: {e}", exc_info=True)
            return GenerationResult(
                success=False,
                error=str(e)
            )

    def _create_presentation(
        self,
        template_file: Optional[Union[bytes, BytesIO, str, Path]]
    ) -> Presentation:
        """Create or load a presentation."""
        pres_config = self.app_config.presentation

        if template_file:
            try:
                if isinstance(template_file, bytes):
                    template_file = BytesIO(template_file)
                prs = Presentation(template_file)
                logger.info("Loaded presentation template")
            except Exception as e:
                raise PPTXGenerationError(
                    f"Cannot load template: {e}",
                    operation="load_template"
                )
        else:
            prs = Presentation()

            # Set dimensions based on orientation
            if self.config.orientation == "landscape":
                prs.slide_width = Inches(pres_config.landscape_width_inches)
                prs.slide_height = Inches(pres_config.landscape_height_inches)
            else:
                prs.slide_width = Inches(pres_config.portrait_width_inches)
                prs.slide_height = Inches(pres_config.portrait_height_inches)

            logger.info(
                f"Created new presentation: "
                f"{prs.slide_width.inches}x{prs.slide_height.inches} inches"
            )

        return prs

    def _create_slide(
        self,
        prs: Presentation,
        data: dict,
        embedded_images: Dict[str, ImageResult]
    ) -> SlideResult:
        """Create a single slide."""
        index = data.get("row_index", len(prs.slides))
        result = SlideResult(index=index, success=True)

        try:
            # Add blank slide (layout 6 is typically blank)
            slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
            slide = prs.slides.add_slide(slide_layout)

            # Handle image
            image_source = data.get("image_source")
            image_cell = data.get("image_cell")  # Cell reference for embedded image

            img_result = None

            # Check for embedded image first (by cell reference)
            if image_cell and image_cell in embedded_images:
                img_result = embedded_images[image_cell]

            # Check for file path
            elif image_source and not image_source.startswith("http"):
                img_result = self.loader.load_from_path(image_source)

            # Add image if we have one
            if img_result:
                if img_result.success:
                    try:
                        self._add_image(slide, prs, img_result)
                        result.has_image = True
                    except Exception as e:
                        result.image_error = str(e)
                        logger.warning(f"Slide {index}: Image add failed: {e}")
                else:
                    result.image_error = img_result.error
                    logger.warning(f"Slide {index}: Image load failed: {img_result.error}")

            # Add text
            text_content = data.get("text_content", [])
            if text_content:
                try:
                    self._add_text(slide, prs, text_content)
                    result.text_added = True
                except Exception as e:
                    result.error = f"Text add failed: {e}"
                    logger.warning(f"Slide {index}: Text add failed: {e}")

        except Exception as e:
            result.success = False
            result.error = str(e)
            logger.error(f"Slide {index}: Creation failed: {e}", exc_info=True)

        return result

    def _add_image(
        self,
        slide,
        prs: Presentation,
        img_result: ImageResult
    ) -> None:
        """Add image to slide."""
        # Reset BytesIO position
        img_result.data.seek(0)

        # Add picture with specified width, let height auto-scale
        pic = slide.shapes.add_picture(
            img_result.data,
            Inches(1),  # Temporary left position
            Inches(self.config.img_top),
            width=Inches(self.config.img_width)
        )

        # Center horizontally
        pic.left = int((prs.slide_width - pic.width) / 2)

    def _add_text(
        self,
        slide,
        prs: Presentation,
        text_content: List[Union[str, dict]]
    ) -> None:
        """Add text content to slide with per-column formatting."""
        # Calculate text box dimensions
        text_height = prs.slide_height.inches - self.config.text_top - 0.5

        textbox = slide.shapes.add_textbox(
            Inches(0.5),
            Inches(self.config.text_top),
            Inches(prs.slide_width.inches - 1.0),
            Inches(text_height)
        )

        text_frame = textbox.text_frame
        text_frame.word_wrap = True

        for i, item in enumerate(text_content):
            # Handle both dict (new format) and string (backward compat)
            if isinstance(item, dict):
                text = item.get("text", "")
                col_format = self.config.get_column_format(item.get("column", ""))
            else:
                text = item
                col_format = ColumnFormat(column="", font_size=self.config.font_size)

            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            # Use run for more control over formatting
            run = p.add_run()
            run.text = text

            # Apply per-column formatting
            font = run.font
            font.size = Pt(col_format.font_size)
            font.bold = col_format.bold
            font.italic = col_format.italic
            font.name = col_format.font_name
            font.color.rgb = col_format.get_rgb_color()

            p.alignment = PP_ALIGN.CENTER
            p.space_after = Pt(12)


def create_presentation(
    slide_data: List[dict],
    config: Optional[SlideConfig] = None,
    embedded_images: Optional[Dict[str, ImageResult]] = None,
    template_file: Optional[Union[bytes, BytesIO, str, Path]] = None,
    progress_callback: Optional[Callable[[str, int, int], None]] = None
) -> GenerationResult:
    """
    Convenience function to create a presentation.

    Args:
        slide_data: List of dicts with 'image_source' and 'text_content' keys
        config: Optional slide configuration
        embedded_images: Dict of embedded images
        template_file: Optional template file
        progress_callback: Optional progress callback

    Returns:
        GenerationResult
    """
    generator = PPTXGenerator(config)
    return generator.generate(slide_data, embedded_images, template_file, progress_callback)
