# Project Memory

## Project Overview
Excel to PowerPoint Converter (StimuPop) - A Streamlit web application that converts Excel spreadsheet rows into PowerPoint presentation slides with images and formatted text. Features template-based generation, Rich Data image extraction, uniform image sizing, per-column text formatting, and portable distribution for easy sharing.

**Current Version:** 8.1.0

## Architecture Decisions

### 1. Modular Structure (v2.0.0)
**Decision**: Refactor monolithic 300-line `app.py` into 8 focused modules.

**Rationale**:
- Single Responsibility Principle
- Easier testing and maintenance
- Clear separation of concerns
- Better IDE support with type hints

**Structure**:
```
src/
├── config.py         # Configuration management
├── exceptions.py     # Custom exception hierarchy
├── validators.py     # URL/input validation
├── image_handler.py  # Image downloading
├── excel_handler.py  # Excel processing
├── pptx_generator.py # PowerPoint generation
└── logging_config.py # Logging setup
```

### 2. SSRF Prevention
**Decision**: Block all private IP ranges and validate URLs before download.

**Rationale**: Image URLs come from user-uploaded Excel files, creating SSRF risk.

**Implementation**:
- Block: 127.0.0.0/8, 10.0.0.0/8, 172.16.0.0/12, 192.168.0.0/16, 169.254.0.0/16
- Resolve hostname to check actual IP before connecting
- Configurable domain whitelist/blacklist

### 3. Concurrent Image Downloads
**Decision**: Use ThreadPoolExecutor with 5 workers.

**Rationale**: Sequential downloads were slow for presentations with many images.

**Trade-offs**:
- Pro: 5x faster for 5+ images
- Pro: Configurable worker count
- Con: More complex error handling

### 4. In-Memory Caching
**Decision**: Cache downloaded images in memory with TTL.

**Rationale**:
- Faster than file-based caching
- No disk cleanup needed
- Session-scoped (appropriate for Streamlit)

**Configuration**: 1-hour TTL, max 100 entries

### 5. Skip-and-Continue Error Handling
**Decision**: Failed image downloads don't abort presentation generation.

**Rationale**:
- User preference from design phase
- Better UX for partially valid data
- Detailed error reporting in UI

**Alternative Considered**: Abort on first error (rejected as too strict)

### 6. YAML Configuration
**Decision**: External config.yaml with environment variable overrides.

**Rationale**:
- Easy to modify without code changes
- Environment-specific deployments
- Secure secret handling via env vars

**Pattern**: `APP_SECTION_KEY` format for overrides

### 7. Uniform Image Sizing (v2.3.0)
**Decision**: Implement multiple image sizing modes with "Fit to Box" as default.

**Rationale**:
- Users need consistent image sizes across slides for professional presentations
- Different use cases require different sizing strategies
- Original behavior (fit width only) preserved as an option

**Implementation**:
- Four sizing modes: `fit_box`, `fit_width`, `fit_height`, `stretch`
- `fit_box` (default): Scales images to fit within max width AND height while preserving aspect ratio
- `fit_width`: Fixed width, auto height (original v2.2.0 behavior)
- `fit_height`: Fixed height, auto width
- `stretch`: Exact dimensions (may distort)

**Code Location**: `src/pptx_generator.py:_calculate_scaled_size()`

### 8. Portable Distribution (v2.3.0)
**Decision**: Create self-contained portable bundle with embedded Python.

**Rationale**:
- Testers need simple installation (extract and run)
- No Python installation required on target machine
- All dependencies bundled

**Implementation**:
- `build_portable.bat`: Downloads Python embeddable, installs dependencies
- `StimuPop.bat`: Launcher script for end users
- `StimuPop_Portable.zip`: ~126MB distribution package

### 9. Template-Based Generation (v5.1.0)
**Decision**: Support template placeholder mode alongside blank slide generation.

**Rationale**:
- Users have existing "Variety Card" templates with precise formatting
- Manual recreation of font sizes/styles is error-prone
- Template approach preserves exact design intent

**Implementation**:
- `TEMPLATE_MODE_PLACEHOLDER`: Clone template slide for each row
- Extract template shape info (position, size, font properties)
- Map Excel columns to template paragraphs by index
- Preserve empty paragraphs as spacers

**Template Mapping (Variety Card)**:
| Paragraph | Excel Column | Format |
|-----------|--------------|--------|
| P0 | C (Brand) | Arial 24pt Bold |
| P1 | D (Product Name) | Arial 24pt Bold |
| P2 | *(spacer)* | Empty |
| P3 | E (Size) | Arial 19pt |
| P4 | *(spacer)* | Empty |
| P5 | F (Summary) | Arial 16pt |

### 10. Rich Data Image Extraction (v5.1.0)
**Decision**: Extract images from Excel 365 Rich Data structure.

**Rationale**:
- Users paste images directly into cells (Copy → Paste)
- These images are stored in `xl/richData/` not `xl/drawings/`
- Traditional `openpyxl._images` returns empty for these
- Cell shows `#VALUE!` but image exists in archive

**Implementation**:
```
xl/worksheets/sheet1.xml
  └── <c r="B2" vm="1">  (vm = value metadata index)
        ↓
xl/richData/richValueRel.xml.rels
  └── <Relationship Id="rId1" Target="../media/image1.png"/>
        ↓
xl/media/image1.png (extracted via zipfile)
```

**Code Location**: `src/image_handler.py:_extract_rich_data_images()`

### 11. Configurable Paragraph Spacing (v5.1.0)
**Decision**: Make paragraph spacing configurable with 0pt default.

**Rationale**:
- Previous hardcoded 12pt spacing created unwanted gaps between text lines
- Users wanted text lines closer together (like their templates)
- 0pt default means no extra spacing

**Implementation**:
- `SlideConfig.paragraph_spacing`: Float, points (default 0.0)
- UI slider: 0-24pt range
- Applied via `p.space_after = Pt(spacing)`

### 12. Configurable Image Alignment (v6.0.0)
**Decision**: Add configurable image alignment within bounding box.

**Rationale**:
- Test user DR requested bottom-alignment for variety cards
- Images should align to bottom of designated area, not center
- Different layouts may need different alignment strategies

**Implementation**:
- `ImageAlignment` dataclass with `vertical` and `horizontal` fields
- Vertical: top, center (default), bottom
- Horizontal: left, center (default), right
- `_calculate_image_position()` computes actual position based on alignment

**Code Location**: `src/pptx_generator.py:_calculate_image_position()`

### 13. Per-Column Fixed Positioning (v6.0.0)
**Decision**: Allow columns to have fixed positions independent of preceding content.

**Rationale**:
- Test user DR feedback: columns E and F should stay in same position
- Currently, text flows sequentially - E position depends on C/D content length
- Variety cards need consistent layout regardless of content

**Implementation**:
- `ColumnPosition` dataclass with `mode` (auto/fixed), `top`, `left`, `width`
- Auto mode: text flows from previous content (existing behavior)
- Fixed mode: text placed at exact `top` position in separate textbox
- `_add_text_fixed()`: Creates independent textbox for fixed columns
- `_add_text_auto_flow()`: Groups auto columns into single flowing textbox

**Code Location**:
- `src/pptx_generator.py:_add_text_fixed()`
- `src/pptx_generator.py:_add_text_auto_flow()`

### 14. Simple/Advanced Mode Toggle (v6.0.0)
**Decision**: Provide two UI complexity levels for positioning controls.

**Rationale**:
- Most users just need basic alignment (Simple mode)
- Power users need per-column control (Advanced mode)
- Avoid overwhelming casual users with options

**Implementation**:
- Simple mode (default): Just vertical/horizontal alignment dropdowns
- Advanced mode (checkbox): Adds per-column position expanders
- Column defaults: E at 5.0", F at 6.5" (based on DR's variety card layout)

**UI Location**: `app.py:render_advanced_settings()`

### 15. Pictures Only Mode (v6.2.0)
**Decision**: Add checkbox to skip all text columns and create image-only slides.

**Rationale**:
- Users need to create photo albums and image galleries without text
- Existing workflow required manually clearing text columns or ignoring output
- Simple toggle provides clear intent

**Implementation**:
- `SlideConfig.pictures_only`: Boolean (default False)
- When enabled, `text_columns` parameter is ignored
- Only image column is processed for each slide
- Text box creation is skipped entirely

**Code Location**: `src/pptx_generator.py:_create_slide()`

### 16. Text Overflow Control (v6.2.0)
**Decision**: Add dropdown to control text box overflow behavior.

**Rationale**:
- Long product descriptions or titles can exceed text box boundaries
- Users need choice between preserving text size or preserving layout
- PowerPoint supports multiple autofit modes

**Implementation**:
- `SlideConfig.text_overflow`: String ("resize" or "shrink")
- "resize" (default): `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` - box expands
- "shrink": `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` - text shrinks

**Code Location**: `src/pptx_generator.py:_add_text()`

### 17. Expanded Slider Ranges (v6.2.0)
**Decision**: Expand minimum/maximum values for dimension sliders.

**Rationale**:
- Users requested thumbnail-sized images (smaller than 2 inches)
- Text at absolute top of slide requires 0.0 starting position
- Larger font sizes needed for title slides and headlines

**Changes**:
| Slider | Old Min | New Min | Old Max | New Max |
|--------|---------|---------|---------|---------|
| Image Width | 2.0 | 0.0 | 9.0 | 9.0 |
| Image Height | 2.0 | 0.0 | 7.0 | 7.0 |
| Text Top | 1.0 | 0.0 | 9.0 | 9.0 |
| Font Size | 10 | 8 | 32 | 48 |

**UI Location**: `app.py:render_advanced_settings()`

### 18. Dynamic Template Column Mapping (v7.0.0)
**Decision**: Replace hardcoded column sequence with dynamic spacer detection and column mapping.

**Rationale**:
- Bug #22: Template Mode assumed columns were always C, D, E, F with spacers at positions 2 and 4
- Users with different column letters (e.g., B, C) or different column counts experienced broken output
- Hardcoded sequence `["C", "D", "", "E", "", "F"]` was not flexible

**Problem Analysis**:
```
Hardcoded Assumption:
  Paragraph 0 → Column C
  Paragraph 1 → Column D
  Paragraph 2 → (spacer)
  Paragraph 3 → Column E
  Paragraph 4 → (spacer)
  Paragraph 5 → Column F

Reality: Users may have columns B, C, D or A, B, C, D, E, F, G...
```

**Solution**:
1. **Spacer Detection**: Analyze template paragraphs; empty text indicates spacer position
2. **Non-Spacer Enumeration**: Build list of paragraph indices that are NOT spacers
3. **Dynamic Mapping**: Map user columns (in order) to non-spacer positions (in order)

**Implementation**:
```python
# Detect spacers from template
spacer_positions = {i for i, p in enumerate(template_paragraphs)
                    if not p.text.strip()}

# Get usable positions
non_spacer_positions = [i for i in range(len(paragraphs))
                        if i not in spacer_positions]

# Map columns dynamically
for idx, col in enumerate(user_columns):
    if idx < len(non_spacer_positions):
        paragraphs[non_spacer_positions[idx]].text = content_map[col]
```

**Breaking Change Justification**:
- This is a behavioral change that warrants major version bump
- Users who created workarounds for the bug may see different results
- More correct behavior overall, but edge cases may need review

**Code Location**: `src/pptx_generator.py` lines 584-615

### 19. Excel Upload Retry with Cache Clearing (v7.1.0)
**Decision**: Add a single-retry pattern with Streamlit cache clearing for Excel file uploads.

**Rationale**:
- Users reported intermittent upload failures where the Excel file could not be read
- Root cause: Streamlit's internal file cache occasionally serves a stale or incomplete buffer
- A simple retry after clearing the cache resolves the issue without user intervention

**Implementation**:
- On initial `ExcelProcessor` read failure, call `st.cache_data.clear()` to invalidate cached file data
- Retry the read operation exactly once
- If the retry also fails, surface the original error to the user
- Pattern is intentionally limited to one retry to avoid masking persistent file errors

**Code Location**: `app.py` (Excel upload handler)

### 20. Multi-Element Template Support (v8.0.0)
**Decision**: Support multiple images and text boxes per slide in Template Mode.

**Rationale**:
- Users need slides with multiple product images side-by-side
- Templates often have multiple text boxes for different data sections
- Single-image/single-text assumption limited template design flexibility

**Implementation**:
- `ImageElement` dataclass: column, placeholder_name, sizing_mode, alignment
- `TextGroup` dataclass: columns list, placeholder_name
- `SlideConfig.image_elements` / `SlideConfig.text_groups`: Optional lists
- `get_image_elements()` / `get_text_groups()`: Backward-compat helpers that build single-item lists from legacy fields when multi-element is not configured
- `_extract_template_info()`: Builds `image_shapes` and `text_shapes` dicts keyed by shape name
- `_populate_slide_multi()`: Loops over each element, matching by placeholder name
- `_populate_slide_legacy()`: Preserves exact pre-v8.0 behavior
- `validate_columns_multi()` / `get_slide_data_multi()`: New Excel handler methods
- UI: Session state counter pattern for dynamic add/remove of elements

**Scope**: Template Mode only (v8.0). Blank Mode multi-element deferred to v8.1.

**Code Locations**:
- `src/pptx_generator.py`: ImageElement, TextGroup, SlideConfig extensions, _populate_slide_multi
- `src/excel_handler.py`: validate_columns_multi, get_slide_data_multi
- `app.py`: Multi-element UI section, generate_presentation branching

## Coding Conventions

### Type Hints
All functions use type hints:
```python
def download_image(url: str, timeout: int = 30) -> Optional[BytesIO]:
```

### Docstrings
Google-style docstrings with Args, Returns, Raises:
```python
def example(param: str) -> bool:
    """
    Short description.

    Args:
        param: Description

    Returns:
        Description

    Raises:
        ValueError: When invalid
    """
```

### Exception Handling
- Use custom exceptions from `src/exceptions.py`
- Include context (URL, filename, row number)
- Log errors before raising

### Logging
- Use `get_logger(__name__)` for module loggers
- Include request ID for traceability
- Log at appropriate levels (INFO for operations, WARNING for recoverable, ERROR for failures)

## Known Pitfalls

### 1. Streamlit Caching
**Issue**: `@st.cache_data` doesn't work with file objects.
**Solution**: Convert to bytes before caching.

### 2. BytesIO Position
**Issue**: After reading BytesIO, position is at end.
**Solution**: Call `.seek(0)` before reuse.

### 3. PIL Image Verification
**Issue**: `Image.open()` doesn't fully validate.
**Solution**: Call `img.verify()` after opening.

### 4. Excel Column Letters
**Issue**: Column letters can be single (A-Z) or double (AA-AZ...).
**Solution**: Use formula: `result = result * 26 + (ord(char) - ord('A') + 1)`

## Code Map

### Entry Points
- `app.py:main()` - Streamlit application entry
- `app.py:generate_presentation()` - Core generation logic
- `build_portable.bat` - Creates portable distribution
- `create_user_guide.py` - Generates DOCX user guide

### Key Classes
- `Config` - Configuration management
- `URLValidator` - Security-focused URL validation
- `ImageDownloader` - Concurrent image downloads
- `ImageCache` - TTL-based image caching
- `ExcelProcessor` - Excel file handling
- `PPTXGenerator` - PowerPoint generation
- `SlideConfig` - Slide layout configuration (includes `img_height`, `img_size_mode`)
- `ColumnFormat` - Per-column text formatting
- `ImageElement` - Multi-element image config (v8.0)
- `TextGroup` - Multi-element text group config (v8.0)

### Image Sizing Constants
- `IMG_SIZE_FIT_BOX` - Fit within box, preserve aspect ratio
- `IMG_SIZE_FIT_WIDTH` - Fixed width, auto height
- `IMG_SIZE_FIT_HEIGHT` - Fixed height, auto width
- `IMG_SIZE_STRETCH` - Exact size, may distort

### Template Mode Constants (v5.1.0)
- `TEMPLATE_MODE_BLANK` - Generate blank slides (original behavior)
- `TEMPLATE_MODE_PLACEHOLDER` - Clone template and populate placeholders

### Data Flow
```
Excel File → ExcelProcessor → slide_data → PPTXGenerator
                │                              ↓
                │                ImageLoader → images
                │                (including Rich Data extraction)
                │                              ↓
                │                Template Mode?
                │                ├── BLANK: _create_slide()
                │                └── PLACEHOLDER: _create_slide_from_template()
                │                    ├── Multi: _populate_slide_multi()
                │                    └── Legacy: _populate_slide_legacy()
                │                              ↓
                │                         Presentation
                │
                ├── Single-element: get_slide_data()
                └── Multi-element: get_slide_data_multi()
```

### Rich Data Image Flow (v5.1.0)
```
Excel (xlsx archive)
├── xl/worksheets/sheet1.xml  → Cell vm="N" attributes
├── xl/richData/richValueRel.xml.rels → rIdN → image path
└── xl/media/imageN.png → Actual image data
```

## Test Coverage

### Critical Paths (Must Test)
1. URL validation (security)
2. Private IP blocking (security)
3. Image download with errors
4. Excel column resolution
5. Slide generation with/without images
6. Image sizing modes (fit_box, fit_width, fit_height, stretch)
7. Aspect ratio preservation in fit modes

### Test Fixtures
Located in `tests/conftest.py`:
- `sample_dataframe` - Basic test DataFrame
- `sample_excel_bytes` - Excel file as bytes
- `mock_image_result` - Successful image download
- `mock_failed_image_result` - Failed image download

### Image Sizing Test Cases
- Wide image (1920x1080) with fit_box → should fit within bounds
- Tall image (1080x1920) with fit_box → should fit within bounds
- Square image with fit_box → should fit within bounds
- Any image with stretch → should match exact dimensions

## Deployment Notes

### Environment Variables
```bash
APP_IMAGES_MAX_SIZE_MB=20
APP_SECURITY_BLOCK_PRIVATE_IPS=true
APP_LOGGING_LEVEL=INFO
```

### Azure VM Considerations
- Bind to 0.0.0.0 for external access
- Use production logging level (INFO or WARNING)
- Consider reverse proxy for HTTPS

### Log Files
Default location: `logs/app.log`
Rotation: 10MB max, 5 backups

### Portable Distribution
For sharing with testers:
1. Run `build_portable.bat` to create distribution
2. Distribution created in `dist/` folder
3. ZIP file: `StimuPop_Portable.zip` (~126MB)
4. Tester extracts ZIP and runs `StimuPop.bat`

### Distribution Contents
```
dist/
├── StimuPop.bat          # Launch script
├── StimuPop_User_Guide.docx  # User documentation
├── app.py                # Main application
├── config.yaml           # Configuration
├── requirements.txt      # Dependencies list
├── src/                  # Application modules
├── python/               # Embedded Python 3.11.9
└── logs/                 # Log directory
```
