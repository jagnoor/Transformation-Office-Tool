# CLAUDE.md — AI Assistant Guide

## Project Overview

**Transformation Office — Block & Gantt Creator Tool** (v0.9 Beta)

A Streamlit web app that transforms Excel spreadsheets into presentation-ready project visualizations: Gantt charts and space-filling block diagrams. Users upload `.xlsx` files and export results as PNG, PDF, or editable PowerPoint.

## Tech Stack

- **Language**: Python 3.11+
- **Framework**: Streamlit (port 8501)
- **Data**: pandas, openpyxl
- **Visualization**: matplotlib
- **Export**: python-pptx (PowerPoint), Pillow (images)
- **Dev Environment**: VS Code DevContainer (Debian Bookworm)

## Repository Structure

```
├── app.py                  # Main Streamlit UI (entry point)
├── models.py               # Data models: WorkItem, ChartConfig, color palettes
├── gantt_renderer.py       # Gantt chart rendering (matplotlib)
├── block_renderer.py       # Block diagram rendering (matplotlib)
├── pptx_export.py          # PowerPoint export (native editable shapes)
├── excel_io.py             # Excel reading, validation, template generation
├── requirements.txt        # Python dependencies
├── run.sh                  # Cross-platform launcher script
├── README.md               # User documentation
├── sample_data/
│   └── sample_roadmap.xlsx # 25-item sample dataset
├── .devcontainer/
│   └── devcontainer.json   # GitHub Codespaces / VS Code config
└── .github/workflows/
    ├── claude.yml           # Claude interactive workflow (@claude mentions)
    └── claude-code-review.yml  # Automated PR review
```

## Architecture

The app follows a layered architecture with clear module boundaries:

```
User → app.py (UI) → excel_io.py (parsing) → models.py (data)
                    → gantt_renderer.py (visualization)
                    → block_renderer.py (visualization)
                    → pptx_export.py (export)
```

**Data flow**: Upload Excel → parse to `WorkItem` list → configure via `ChartConfig` → render chart → export

### Key Data Models (`models.py`)

- **`WorkItem`** (dataclass): `title`, `start_date`, `end_date`, `category`, `description`, `status`, `owner`, `label`, `is_milestone`, `color_override`. Milestones auto-detected when start == end date.
- **`ChartConfig`** (dataclass): Chart title, subtitle, palette, toggles (today line, status, legend), date range, slide size (WIDE/A4/A3).
- **Color palettes**: Ocean, Sunset, Forest, Corporate (default), Vibrant, Monochrome — each with 10 category colors.
- **Status colors**: planned (gray), in_progress (blue), done (green), at_risk (red).

### Module Details

- **`app.py`**: Streamlit session state management, sidebar controls, homepage with onboarding cards, visualization tabs (Gantt/Block/Data), export buttons (PNG 300 DPI, PDF, PPTX).
- **`excel_io.py`**: Flexible column name matching (e.g., "Name" → "Title"), multi-format date parsing (YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY, etc.), auto-swap reversed dates, template generation with formatting and validation.
- **`gantt_renderer.py`**: Swim lane layout, auto-stacking to prevent overlaps, category color coding, today marker, smart date axis granularity. Font fallback: Arial → Helvetica → DejaVu Sans.
- **`block_renderer.py`**: Greedy row-packing algorithm, blocks sized proportionally to duration, text wrapping, contrast-aware text colors.
- **`pptx_export.py`**: Generates native PowerPoint shapes (not images), 16:9 slide size (13.333" x 7.5"), hex-to-RGB color conversion.

## Running the App

```bash
# Quick start
bash run.sh

# Manual
pip install -r requirements.txt
streamlit run app.py --server.headless true
```

The app runs on `http://localhost:8501`.

## Development Conventions

### Code Style
- Python with type hints on dataclasses and function signatures
- Modules are self-contained with clear docstrings at the top
- Constants defined at module level (e.g., `ROW_HEIGHT`, `BAR_HEIGHT` in renderers)
- Streamlit session state keys: `items`, `config`, `warnings`, `load_error`
- Error handling via try-except with user-facing messages and expandable details

### Naming
- Snake_case for all Python identifiers
- Descriptive function names: `render_gantt()`, `read_excel()`, `export_gantt_to_pptx()`
- File names match their primary purpose: `gantt_renderer.py`, `block_renderer.py`

### Rendering
- Preview DPI: 150, Export DPI: 300
- All renderers return matplotlib `Figure` objects
- Color utilities (lighten, darken, contrast detection) live in `block_renderer.py`

### Input Validation
- 3 required Excel columns: Title, Start Date, End Date
- 5 optional columns: Category, Description, Status, Owner, Label
- Warnings collected in a list and displayed to users (non-blocking)
- Invalid rows are skipped with warnings rather than raising errors

## Testing

There is no automated test suite currently. When adding tests:
- Use `pytest` as the test framework
- Place tests in a `tests/` directory
- Test data parsing (excel_io), rendering output (gantt/block), and model validation

## CI/CD

- **`claude.yml`**: Responds to `@claude` mentions in issues/PRs using `anthropics/claude-code-action@v1`
- **`claude-code-review.yml`**: Auto-reviews PRs on create/update using the code-review plugin

## Common Tasks

### Adding a new color palette
1. Add palette name and 10 hex colors to `COLOR_PALETTES` in `models.py`
2. Add the name to `PALETTE_NAMES` list in `models.py`
3. Both renderers pick it up automatically via `ChartConfig.palette`

### Adding a new export format
1. Create a new module (e.g., `svg_export.py`)
2. Accept `items: list[WorkItem]` and `config: ChartConfig` as inputs
3. Wire up in `app.py` with a new download button in the export section

### Modifying chart appearance
- Gantt layout constants: `gantt_renderer.py` top-level (ROW_HEIGHT, BAR_HEIGHT, etc.)
- Block layout: `block_renderer.py` greedy packing algorithm in `_calculate_layout()`
- Fonts: fallback chain defined in each renderer's `_get_font()` or similar

### Adding new Excel columns
1. Add field to `WorkItem` dataclass in `models.py`
2. Add column name mapping in `excel_io.py` `COLUMN_ALIASES`
3. Update renderers if the field should be displayed
