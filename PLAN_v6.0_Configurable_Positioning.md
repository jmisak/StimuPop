# StimuPop v6.0 Implementation Plan
## Configurable Positioning System

**Status:** IMPLEMENTED
**Requested by:** User (DR feedback)
**Date:** 2026-01-28

---

## Overview

Implement configurable image alignment and per-column text positioning with Simple + Advanced mode toggle.

**Key Features:**
1. Image vertical alignment (top/center/bottom) - default: center
2. Image horizontal alignment (left/center/right) - default: center
3. Per-column fixed text positions (E/F stay fixed regardless of C/D content)
4. Simple/Advanced mode toggle in UI

---

## Phase 1: Schema & Dataclasses

### 1.1 New Dataclass: ImageAlignment

**File:** `src/pptx_generator.py`

```python
@dataclass
class ImageAlignment:
    vertical: str = "center"    # top, center, bottom
    horizontal: str = "center"  # left, center, right
```

### 1.2 New Dataclass: ColumnPosition

**File:** `src/pptx_generator.py`

```python
@dataclass
class ColumnPosition:
    mode: str = "auto"           # auto, fixed
    top: Optional[float] = None  # inches from top (if fixed)
    left: float = 0.5            # inches from left
    width: Optional[float] = None # None = auto (slide width - margins)
```

### 1.3 Extended SlideConfig

**File:** `src/pptx_generator.py`

Add to existing SlideConfig:
```python
# NEW fields
image_alignment: Optional[ImageAlignment] = None
column_positions: Optional[Dict[str, ColumnPosition]] = None
positioning_mode: str = "simple"  # simple, advanced
```

---

## Phase 2: Generator Logic

### 2.1 Image Position Calculation

**File:** `src/pptx_generator.py`

New method `_calculate_image_position()`:

```
INPUTS: scaled_width, scaled_height, alignment, bounding_box
OUTPUT: (left_inches, top_inches)

Vertical:
  - top:    img_top = box_top
  - center: img_top = box_top + (box_height - img_height) / 2
  - bottom: img_top = box_top + box_height - img_height

Horizontal:
  - left:   img_left = box_left
  - center: img_left = box_left + (box_width - img_width) / 2
  - right:  img_left = box_left + box_width - img_width
```

### 2.2 Update `_add_image()` method

- Check if `image_alignment` is set
- If None, use legacy center behavior (backward compatible)
- If set, call `_calculate_image_position()`

### 2.3 New method `_add_text_with_positions()`

For columns with `mode=fixed`:
- Create separate textbox at specified position
- Each fixed column = independent shape

For columns with `mode=auto`:
- Group into single textbox (current behavior)
- Flow sequentially from `text_top`

**Logic flow:**
```
1. Separate columns into auto_columns and fixed_columns
2. If auto_columns exist:
   - Create single textbox at text_top
   - Add paragraphs for each auto column
3. For each fixed_column:
   - Create separate textbox at column.top position
   - Add single paragraph with column content
```

---

## Phase 3: UI Implementation

### 3.1 Simple Mode UI

**File:** `app.py`

Add to Advanced Settings expander:

```
Image Alignment:
├── Vertical: [Dropdown: Top | Center* | Bottom]
└── Horizontal: [Dropdown: Left | Center* | Right]

(* = default)
```

### 3.2 Advanced Mode Toggle

```
[Checkbox] Enable Advanced Positioning

If checked, show:
├── Image Bounding Box (optional - for power users)
│   ├── Box Top: [slider]
│   ├── Box Left: [slider]
│   ├── Box Width: [slider]
│   └── Box Height: [slider]
│
└── Per-Column Positioning
    └── Tabs: [C] [D] [E] [F]
        └── Position Mode: (Auto) (Fixed)
            └── If Fixed: Top Position [input] inches
```

### 3.3 SlideConfig Construction

Update config creation in `app.py`:

```python
# Simple mode - just alignment dropdowns
image_alignment = ImageAlignment(
    vertical=img_v_align,      # from dropdown
    horizontal=img_h_align     # from dropdown
)

# Advanced mode - add column positions
if advanced_mode:
    column_positions = {}
    for col in text_columns:
        if col_mode[col] == "fixed":
            column_positions[col] = ColumnPosition(
                mode="fixed",
                top=col_top[col]
            )
        # auto columns don't need entry
```

---

## Phase 4: Template Mode Enhancement

### 4.1 Extract alignment from template

When `template_mode = TEMPLATE_MODE_PLACEHOLDER`:
- Detect image position relative to placeholder bounds
- Infer alignment (if image bottom matches placeholder bottom → "bottom")

### 4.2 Extract column positions from template

- Parse paragraph positions in text placeholder
- Map to column_positions automatically
- Allow override via Advanced UI

---

## File Changes Summary

| File | Changes |
|------|---------|
| `src/pptx_generator.py` | Add ImageAlignment, ColumnPosition dataclasses; extend SlideConfig; add _calculate_image_position(); modify _add_image(); add _add_text_with_positions() |
| `src/__init__.py` | Export new classes |
| `app.py` | Add alignment dropdowns; add Advanced toggle; add per-column position UI |
| `config.yaml` | Add default_image_alignment section (optional) |
| `src/config.py` | Add defaults for new fields |

---

## Testing Checklist

- [ ] Image vertical alignment: top works
- [ ] Image vertical alignment: center works (default)
- [ ] Image vertical alignment: bottom works
- [ ] Image horizontal alignment: left/center/right work
- [ ] Column E fixed position stays fixed when C/D content varies
- [ ] Column F fixed position stays fixed
- [ ] Auto columns still flow correctly
- [ ] Mixed auto + fixed columns work together
- [ ] Backward compatibility: old configs still work
- [ ] Template mode still works
- [ ] Simple mode hides advanced options
- [ ] Advanced mode shows all options

---

## Backward Compatibility

- `image_alignment = None` → use legacy center behavior
- `column_positions = None` → use legacy sequential flow
- `positioning_mode = "simple"` → default, matches current behavior
- Existing config.yaml files continue to work unchanged

---

## Deliverables

1. Updated `src/pptx_generator.py` with new positioning logic
2. Updated `app.py` with Simple + Advanced UI
3. Updated exports in `src/__init__.py`
4. Updated `MEMORY.md` with architecture decision
5. Updated `VERSION_HISTORY.md` for v6.0.0
6. Rebuilt portable EXE for testing

---

## Estimated Scope

| Component | Effort |
|-----------|--------|
| Dataclasses | Small |
| Image alignment logic | Medium |
| Text positioning logic | Medium |
| UI - Simple mode | Small |
| UI - Advanced mode | Medium |
| Testing | Medium |
| Documentation | Small |

---

## Approval Required

Please confirm:
1. Schema design is acceptable
2. UI layout is acceptable
3. Ready to begin implementation

**Awaiting your approval to proceed.**
