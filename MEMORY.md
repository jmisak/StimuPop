# Project Memory

## Project Overview
Excel to PowerPoint Converter - A Streamlit web application that converts Excel spreadsheet rows into PowerPoint presentation slides with images and formatted text.

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

### Key Classes
- `Config` - Configuration management
- `URLValidator` - Security-focused URL validation
- `ImageDownloader` - Concurrent image downloads
- `ImageCache` - TTL-based image caching
- `ExcelProcessor` - Excel file handling
- `PPTXGenerator` - PowerPoint generation
- `SlideConfig` - Slide layout configuration

### Data Flow
```
Excel File → ExcelProcessor → slide_data → PPTXGenerator
                                              ↓
                              ImageDownloader → images
                                              ↓
                                         Presentation
```

## Test Coverage

### Critical Paths (Must Test)
1. URL validation (security)
2. Private IP blocking (security)
3. Image download with errors
4. Excel column resolution
5. Slide generation with/without images

### Test Fixtures
Located in `tests/conftest.py`:
- `sample_dataframe` - Basic test DataFrame
- `sample_excel_bytes` - Excel file as bytes
- `mock_image_result` - Successful image download
- `mock_failed_image_result` - Failed image download

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
