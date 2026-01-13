# Version History

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
