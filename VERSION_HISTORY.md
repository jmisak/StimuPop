# Version History

## [2026-01-31] - v7.0.0: Template Mode Dynamic Column Mapping (Breaking Change)

### Summary
Major version release fixing a critical bug in Template Mode where hardcoded column sequence assumptions broke presentations with non-standard column configurations. This release introduces dynamic column mapping that correctly handles any user-defined column letters and properly detects spacer paragraphs in templates.

### Critical Bug Fix (Issue #22)

**Problem**: Template Mode had a hardcoded column sequence `["C", "D", "", "E", "", "F"]` that assumed:
- Paragraph 0 mapped to Column C
- Paragraph 1 mapped to Column D
- Paragraph 2 was an empty spacer
- Paragraph 3 mapped to Column E
- Paragraph 4 was an empty spacer
- Paragraph 5 mapped to Column F

This broke when users had different column configurations (e.g., columns B and C, or a different number of text columns). Users with non-standard setups would see incorrect text placement or missing content.

**Solution**: Implemented dynamic column mapping that:
1. Analyzes template paragraphs to detect which are spacers (empty text in template)
2. Identifies non-spacer paragraph positions dynamically
3. Maps user's actual columns to non-spacer paragraph positions in order
4. Works with any column letters and any number of columns

### Technical Details

#### Affected Code
- **File**: `src/pptx_generator.py`
- **Lines**: 584-615 (Template Mode column mapping logic)

#### Before (Hardcoded)
```python
column_sequence = ["C", "D", "", "E", "", "F"]
for i, col in enumerate(column_sequence):
    if col and col in content_map:
        paragraphs[i].text = content_map[col]
```

#### After (Dynamic)
```python
# Step 1: Detect spacer positions from template
spacer_positions = set()
for i, para in enumerate(template_paragraphs):
    if not para.text.strip():
        spacer_positions.add(i)

# Step 2: Build list of non-spacer positions
non_spacer_positions = [i for i in range(len(paragraphs))
                        if i not in spacer_positions]

# Step 3: Map user columns to non-spacer positions in order
user_columns = list(content_map.keys())
for idx, col in enumerate(user_columns):
    if idx < len(non_spacer_positions):
        para_index = non_spacer_positions[idx]
        paragraphs[para_index].text = content_map[col]
```

#### Algorithm Flow
```
Template Analysis:
  P0: "Brand Name"     → non-spacer → position 0
  P1: "Product Name"   → non-spacer → position 1
  P2: ""               → SPACER     → skip
  P3: "Size Info"      → non-spacer → position 2
  P4: ""               → SPACER     → skip
  P5: "Description"    → non-spacer → position 3

User Content Map: {"B": "Acme", "C": "Widget", "D": "Large", "E": "Details"}

Dynamic Mapping:
  Column B → non_spacer_positions[0] → P0
  Column C → non_spacer_positions[1] → P1
  Column D → non_spacer_positions[2] → P3 (skips P2 spacer)
  Column E → non_spacer_positions[3] → P5 (skips P4 spacer)
```

### Why Major Version (v7.0.0)

This release is marked as a **major version** because:

1. **Behavioral Change**: Template Mode now works fundamentally differently in how it maps columns to paragraphs

2. **Potential Breaking Change**: Users who had worked around the bug by:
   - Specifically using columns C, D, E, F to match the hardcoded sequence
   - Adding dummy columns to align with expected positions
   - Restructuring their templates to match the assumed layout

   These workarounds may now produce different (more correct) results

3. **Edge Case Considerations**: Templates with unusual spacer patterns may render differently than before

### Migration Notes

- **Standard Users (C/D/E/F columns)**: No action required; behavior unchanged for the original assumed configuration
- **Non-Standard Column Users**: Presentations should now work correctly without workarounds
- **Workaround Users**: Review generated output; remove any dummy columns or template modifications made to accommodate the bug

### Files Modified
- `src/pptx_generator.py` - Dynamic column mapping implementation
- `VERSION_HISTORY.md` - This changelog
- `MEMORY.md` - Architecture decision #18

---

## [2026-01-30] - v6.2.0: Pictures Only Mode & Text Overflow Control

### Summary
Feature release adding Pictures Only mode for image-only slideshows, Text Overflow control options, expanded UI slider ranges, and improved label clarity throughout the interface.

### Features Added
- **Pictures Only Mode**: New checkbox "Pictures Only (no text)" in Basic Configuration
  - When enabled, text columns are skipped entirely
  - Use case: Creating image-only slideshows, photo albums, image catalogs
  - Text Columns setting is ignored when enabled
- **Text Overflow Option**: New dropdown in Advanced Settings under Text Spacing
  - "Resize shape to fit text" (default): Text box expands to fit content
  - "Shrink text on overflow": Text shrinks to fit within boundaries
  - Gives users control over how long text is handled

### UI Improvements
- **Image Width/Height Sliders**: Minimum changed from 2.0 to 0.0 inches
  - Allows thumbnail-sized images for compact layouts
- **Text Top Position Slider**: Minimum changed from 1.0 to 0.0 inches
  - Allows text to be positioned at the very top of the slide
- **Font Size Range**: Expanded from 10-32pt to 8-48pt
  - Supports both smaller captions and larger headlines
- **Label Clarifications**:
  - Image Sizing: Removed "(Blank mode only)" - works for both Blank and Template modes
  - Layout Position: Added "(Blank mode only)" label for clarity
  - Image Alignment: Added "(Blank mode only)" label for clarity

### Documentation Updates
- Added Template Mode summary section in Advanced Settings
- Moved Image Alignment explanation after Layout Position for better flow
- Added Text Spacing section with Paragraph Spacing and Text Overflow documentation
- Added Column Flexibility section clarifying Excel column requirements
- Updated all version references to 6.2.0

### Technical Details

#### New Configuration Options
| Option | Location | Values | Default |
|--------|----------|--------|---------|
| `pictures_only` | Basic Config | Boolean | False |
| `text_overflow` | Advanced > Text Spacing | "resize" / "shrink" | "resize" |

#### Slider Range Changes
| Slider | Old Range | New Range |
|--------|-----------|-----------|
| Image Width | 2.0-9.0 | 0.0-9.0 |
| Image Height | 2.0-7.0 | 0.0-7.0 |
| Text Top Position | 1.0-9.0 | 0.0-9.0 |
| Font Size | 10-32 | 8-48 |

### Files Modified
- `app.py` - UI changes for new features and slider ranges
- `src/pptx_generator.py` - Pictures Only and Text Overflow implementation
- `create_user_guide.py` - Documentation generator updates
- `StimuPop_User_Guide.docx` - Regenerated with v6.2.0 content
- `StimuPop_User_Guide.html` - New HTML version of user guide
- `VERSION_HISTORY.md` - This changelog
- `MEMORY.md` - Architecture decisions for new features

---

## [2026-01-30] - v6.1.0: User Guide Documentation Update

### Summary
Documentation update release addressing user feedback to improve clarity and usability of the User Guide. Updated for v6.1.0 with enhanced first-launch instructions and documentation of v6.0 features.

### Documentation Updates
- **Version**: Updated throughout to 6.1.0
- **First Launch Instructions**: Clarified that users should wait for "Server ready!" message before clicking the localhost link
- **Browser Access**: Added instructions to Ctrl+click the localhost link or copy/paste the URL into browser
- **First-Launch Timing Note**: Clarified that server may take 30-60 seconds on first launch
- **Layout Position Settings**: Added note that these settings only appear in Blank mode (not Template mode)
- **Per-Column Formatting**: Added note that these settings only appear in Blank mode
- **Image Alignment (v6.0)**: Documented vertical (Top/Center/Bottom) and horizontal (Left/Center/Right) alignment options
- **Advanced Positioning (v6.0)**: Documented the Advanced Positioning checkbox and per-column position controls
  - Auto vs Fixed positioning modes
  - Default positions for Column E (5.0") and Column F (6.5")
- **Step-by-Step Guide**: Updated from 9 steps to 10 steps to include waiting for server and Ctrl+click instructions

### Files Modified
- `create_user_guide.py` - Generator script updated with all documentation changes
- `StimuPop_User_Guide.docx` - Regenerated with updated content
- `MEMORY.md` - Version updated to 6.1.0
- `VERSION_HISTORY.md` - Added v6.1.0 release notes

---

## [2026-01-29] - v6.0.0: Configurable Positioning System

### Summary
User-requested feature release adding configurable image alignment and per-column text positioning with Simple + Advanced mode toggle. Addresses feedback from test user DR for variety card layout customization.

### Features Added
- **Image Alignment**: Vertical (top/center/bottom) and horizontal (left/center/right) alignment
  - Default: center (backward compatible)
  - DR feedback: bottom alignment preferred for variety cards
- **Per-Column Fixed Positioning**: Columns E and F can have fixed positions regardless of C/D content length
  - Auto mode: text flows sequentially (existing behavior)
  - Fixed mode: text placed at exact top position
- **Simple/Advanced Mode Toggle**: Hide complexity from casual users
  - Simple mode: Just image alignment dropdowns
  - Advanced mode: Full per-column position controls

### Technical Details

#### New Components
| Component | Description |
|-----------|-------------|
| `ImageAlignment` | Dataclass with vertical/horizontal alignment |
| `ColumnPosition` | Dataclass with mode (auto/fixed), top, left, width |
| `SlideConfig.image_alignment` | Image positioning configuration |
| `SlideConfig.column_positions` | Dict mapping columns to ColumnPosition |
| `SlideConfig.positioning_mode` | "simple" or "advanced" |
| `_calculate_image_position()` | Calculates image position based on alignment |
| `_add_text_fixed()` | Adds text at fixed position (separate textbox) |
| `_add_text_auto_flow()` | Adds text in sequential flow |

#### New Constants
| Constant | Value | Description |
|----------|-------|-------------|
| `IMG_ALIGN_TOP` | "top" | Vertical alignment |
| `IMG_ALIGN_CENTER` | "center" | Vertical alignment (default) |
| `IMG_ALIGN_BOTTOM` | "bottom" | Vertical alignment |
| `IMG_ALIGN_LEFT` | "left" | Horizontal alignment |
| `IMG_ALIGN_RIGHT` | "right" | Horizontal alignment (default: center) |

#### Image Alignment Logic
```
Vertical:
  - top:    img_top = box_top
  - center: img_top = box_top + (box_height - img_height) / 2
  - bottom: img_top = box_top + box_height - img_height

Horizontal:
  - left:   img_left = box_left
  - center: img_left = box_left + (box_width - img_width) / 2
  - right:  img_left = box_left + box_width - img_width
```

#### UI Changes
```
Advanced Settings
├── Image Alignment (NEW)
│   ├── Vertical Alignment: [Center | Top | Bottom]
│   └── Horizontal Alignment: [Center | Left | Right]
├── Advanced Positioning (NEW, optional)
│   └── Per-Column Position
│       ├── Column C: [Auto | Fixed]
│       ├── Column D: [Auto | Fixed]
│       ├── Column E: [Auto | Fixed] ← default 5.0"
│       └── Column F: [Auto | Fixed] ← default 6.5"
├── Template Mode
├── Text Spacing
├── Image Sizing
└── Column Formatting
```

### API Changes
- `SlideConfig` accepts `image_alignment`, `column_positions`, `positioning_mode`
- New exports: `ImageAlignment`, `ColumnPosition`, `IMG_ALIGN_TOP`, `IMG_ALIGN_CENTER`, `IMG_ALIGN_BOTTOM`, `IMG_ALIGN_LEFT`, `IMG_ALIGN_RIGHT`

### Backward Compatibility
- `image_alignment = None` → uses legacy center behavior
- `column_positions = None` → uses legacy sequential flow
- All existing configurations continue to work unchanged

---

## [2026-01-27] - v5.1.0: Template Mode & Rich Data Image Support

### Summary
Major feature release adding template-based slide generation and support for Excel 365 Rich Data images (pasted images). Slides now preserve exact template formatting and layout.

### Features Added
- **Template Placeholder Mode**: Clone template slides and populate with data
  - Preserves exact font sizes, styles, and positioning from template
  - Maps Excel columns to template paragraphs (C→P0, D→P1, E→P3, F→P5)
  - Replaces image placeholder with actual product images
- **Rich Data Image Extraction**: Support for Excel 365 pasted images
  - Extracts images from `xl/richData/` structure
  - Maps cell `vm` attributes to image files via relationship chain
  - Works with images pasted directly into cells (shows as `#VALUE!`)
- **Configurable Paragraph Spacing**: Control gap between text lines
  - Default 0pt (no extra spacing)
  - Slider range 0-24pt in UI

### Technical Details

#### New Components
| Component | Description |
|-----------|-------------|
| `TEMPLATE_MODE_BLANK` | Constant for original blank slide generation |
| `TEMPLATE_MODE_PLACEHOLDER` | Constant for template-based generation |
| `SlideConfig.template_mode` | Selected generation mode |
| `SlideConfig.paragraph_spacing` | Space after each paragraph (points) |
| `SlideConfig.image_placeholder_name` | Name of image shape in template |
| `SlideConfig.text_placeholder_name` | Name of text shape in template |
| `_extract_template_info()` | Extracts shape data from template slide |
| `_create_slide_from_template()` | Creates slide using template layout |
| `_extract_rich_data_images()` | Extracts Excel 365 pasted images |

#### Rich Data Image Extraction Flow
```
Excel File (xl/worksheets/sheet1.xml)
  └── Cell B2: vm="1" (value metadata index)
        │
        ▼
xl/richData/richValueRel.xml.rels
  └── rId1 → ../media/image1.png
        │
        ▼
xl/media/image1.png (extracted)
```

#### Template Mapping (Variety Card Format)
| Template Paragraph | Excel Column | Font |
|--------------------|--------------|------|
| P0 | C (Brand) | Arial 24pt Bold |
| P1 | D (Product Name) | Arial 24pt Bold |
| P2 | *(spacer)* | - |
| P3 | E (Size) | Arial 19pt |
| P4 | *(spacer)* | - |
| P5 | F (Summary) | Arial 16pt |

#### UI Changes
```
Advanced Settings
├── Template Mode (NEW)
│   ├── Generation Mode dropdown (Blank/Placeholder)
│   ├── Image Placeholder Name input
│   └── Text Placeholder Name input
├── Text Spacing (NEW)
│   └── Paragraph Spacing slider (0-24pt)
├── Image Sizing
└── Column Formatting
```

### API Changes
- `SlideConfig` accepts `template_mode`, `paragraph_spacing`, `image_placeholder_name`, `text_placeholder_name`
- `ImageLoader.extract_embedded_images()` now falls back to Rich Data extraction
- New exports: `TEMPLATE_MODE_BLANK`, `TEMPLATE_MODE_PLACEHOLDER`

### Portable Distribution
- `Stimupopv5.1_FINAL.zip` (507 MB) - Fully standalone executable

---

## [2026-01-13] - v2.3.0: Uniform Image Sizing & Portable Distribution

### Summary
Added uniform image sizing with multiple sizing modes, ensuring all images appear consistent across slides. Created portable distribution system for easy sharing with testers.

### Features Added
- **Image Size Modes**: Four modes for controlling image dimensions
  - `Fit to Box` (default): Scale to fit within max width/height, preserve aspect ratio
  - `Fit Width`: Fixed width, auto-calculated height
  - `Fit Height`: Fixed height, auto-calculated width
  - `Stretch`: Exact dimensions (may distort images)
- **Max Height Control**: New slider for maximum image height
- **Portable Distribution**: Self-contained ZIP with embedded Python
- **User Guide**: Professional 12-page DOCX documentation
- **README**: Comprehensive project documentation

### Technical Details

#### New Components
| Component | Description |
|-----------|-------------|
| `IMG_SIZE_FIT_BOX` | Constant for fit-to-box sizing mode |
| `IMG_SIZE_FIT_WIDTH` | Constant for fit-width sizing mode |
| `IMG_SIZE_FIT_HEIGHT` | Constant for fit-height sizing mode |
| `IMG_SIZE_STRETCH` | Constant for stretch sizing mode |
| `SlideConfig.img_height` | Maximum image height in inches |
| `SlideConfig.img_size_mode` | Selected sizing mode |
| `_calculate_scaled_size()` | Calculates dimensions based on mode |
| `_get_image_dimensions()` | Gets original image size via PIL |

#### UI Changes
```
Advanced Settings
├── Image Sizing (NEW)
│   ├── Size Mode dropdown
│   ├── Max Width slider
│   ├── Max Height slider (NEW)
│   └── Mode description info box
├── Layout Position
│   ├── Image Top Position
│   ├── Text Top Position
│   └── Slide Orientation
└── Column Formatting
```

#### Distribution Files
| File | Purpose |
|------|---------|
| `build_portable.bat` | Creates portable distribution |
| `StimuPop.bat` | Launcher for end users |
| `StimuPop_Portable.zip` | ~126MB distribution package |
| `create_user_guide.py` | Generates DOCX documentation |

### API Changes
- `SlideConfig` now accepts `img_height` and `img_size_mode` parameters
- New exports: `IMG_SIZE_FIT_BOX`, `IMG_SIZE_FIT_WIDTH`, `IMG_SIZE_FIT_HEIGHT`, `IMG_SIZE_STRETCH`

### Configuration Changes
```yaml
presentation:
  default_img_height: 4.0      # NEW
  default_img_size_mode: "fit_box"  # NEW
```

---

## [2026-01-12] - v2.2.0: Per-Column Font Formatting

### Summary
Added per-column font formatting, allowing users to configure distinct formatting (font size, bold, italic, font name, color) for each text column via a tabbed UI.

### Features Added
- **ColumnFormat Dataclass**: New configuration object for per-column styling
- **Tabbed UI Controls**: Configure each text column's formatting separately
- **Per-Column Properties**: Font size, bold, italic, font family, and color
- **Backward Compatibility**: Existing configurations still work without column_formats

### Technical Details

#### New Components
| Component | Description |
|-----------|-------------|
| `ColumnFormat` | Dataclass with font_size, bold, italic, font_name, color |
| `SlideConfig.column_formats` | Dict mapping column letters to ColumnFormat |
| `SlideConfig.get_column_format()` | Helper to get format with fallback |
| `render_column_format_config()` | Streamlit UI for per-column settings |

#### Data Flow Changes
| Before (v2.1) | After (v2.2) |
|---------------|--------------|
| `text_content: ["Title", "Desc"]` | `text_content: [{"column": "C", "text": "Title"}, ...]` |
| Single font_size for all text | Per-column formatting via ColumnFormat |
| Uniform paragraph styling | Distinct formatting per column |

#### API Changes
- `get_slide_data()` now has `preserve_column_identity` parameter (default True)
- `SlideConfig` accepts optional `column_formats: Dict[str, ColumnFormat]`
- `_add_text()` handles both string and dict text content formats

### UI Layout
```
Advanced Settings
├── Layout (Image Width, Position, Orientation)
└── Column Formatting
    ├── [Column C] Font Size | Font | Color | Bold | Italic
    ├── [Column D] Font Size | Font | Color | Bold | Italic
    └── ...
```

---

## [2026-01-12] - v2.1.0: StimuPop Rebrand + Embedded Image Support

### Summary
Major update rebranding to StimuPop and adding support for embedded Excel images and local file paths, removing URL-based image downloading.

### Features Changed
- **Rebranded to StimuPop**: New name, updated UI, and branding throughout
- **Embedded Image Support**: Extract images embedded directly in Excel cells
- **Local File Path Support**: Load images from local file paths
- **Removed URL Downloads**: Images now come from Excel or local files (no network access needed)
- **Slide Orientation Selection**: Added UI option for portrait/landscape slides

### Technical Details

#### Image Handling Changes
| Before (v2.0) | After (v2.1) |
|---------------|--------------|
| `ImageDownloader` class | `ImageLoader` class |
| `download()` / `download_many()` | `load_from_path()` / `load_from_bytes()` |
| `extract_embedded_images()` | New method for Excel embedded images |
| URL validation, SSRF prevention | No longer needed (local files only) |

#### Removed Components
- URL validation (validators.py simplified)
- Private IP blocking
- Domain whitelist/blacklist
- Network timeout/retry logic
- `SecurityConfig` class

#### API Changes
- `PPTXGenerator.generate()` now accepts `embedded_images` parameter
- Excel handler now returns `image_source` and `image_cell` instead of `image_url`

### Migration Notes
- Excel files should now contain embedded images or file paths
- URL-based images are no longer supported
- Configuration simplified (security section removed)

---

## [2026-01-12] - v2.0.0: Production Enhancement Release

### Summary
Major refactoring to transform the prototype into production-grade code with security hardening, performance optimizations, and comprehensive testing.

### Features Added
- **Security Module**: URL validation with SSRF prevention (private IP blocking)
- **Concurrent Downloads**: 5 parallel image downloads using ThreadPoolExecutor
- **Image Caching**: In-memory cache with TTL to avoid redundant downloads
- **Retry Logic**: Exponential backoff for transient network failures
- **Configuration System**: YAML-based external configuration with env overrides
- **Logging Infrastructure**: Structured logging with rotation and request ID tracking
- **Custom Exceptions**: Typed exception hierarchy for better error handling

### Technical Details

#### New Project Structure
```
├── src/
│   ├── __init__.py         # Package exports
│   ├── config.py           # Configuration management
│   ├── exceptions.py       # Custom exception classes
│   ├── validators.py       # URL and input validation
│   ├── image_handler.py    # Concurrent image downloads
│   ├── excel_handler.py    # Excel file processing
│   ├── pptx_generator.py   # PowerPoint generation
│   └── logging_config.py   # Logging setup
├── tests/
│   ├── conftest.py         # Test fixtures
│   ├── test_validators.py
│   ├── test_excel_handler.py
│   ├── test_image_handler.py
│   └── test_pptx_generator.py
├── config.yaml             # External configuration
├── requirements.txt        # Production dependencies
└── requirements-dev.txt    # Development dependencies
```

#### Security Enhancements
| Feature | Description |
|---------|-------------|
| Private IP Blocking | Blocks 127.0.0.0/8, 10.0.0.0/8, 172.16.0.0/12, 192.168.0.0/16, 169.254.0.0/16 |
| Protocol Validation | Only allows http/https |
| Domain Filtering | Configurable whitelist/blacklist |
| Size Limits | 10MB per image, 200MB total upload |
| MIME Validation | Verifies image content type |
| Text Sanitization | Removes control characters and null bytes |

#### Performance Improvements
| Feature | Impact |
|---------|--------|
| Concurrent Downloads | 5x faster for multiple images |
| Image Caching | Zero latency for repeated images |
| Streamlit Caching | Faster Excel preview |
| Streaming Downloads | Reduced memory usage |

### Breaking Changes
- Configuration now loaded from `config.yaml` (falls back to defaults if missing)
- Log files written to `logs/app.log` by default

### Dependencies Added
- PyYAML>=6.0 (configuration)

### Migration Notes
- Existing deployments should work without changes
- To customize settings, create `config.yaml` in project root
- Set `APP_*` environment variables to override config values

---

## [2026-01-05] - v1.0.0: Initial Release

### Summary
Initial working prototype of Excel to PowerPoint converter.

### Features
- Excel file upload (.xlsx, .xls)
- Optional PowerPoint template upload
- Image download from URLs
- Configurable slide layout
- Progress indicators
- Download button for generated presentations
