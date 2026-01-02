# Transformation Office Block & Gantt Creator Tool (Excel → PDF + High‑Res PNG + Editable PPTX)

A local Streamlit app for generating executive-ready blocked-diagram / swimlane Gantt roadmap charts from an Excel table.

Outputs:
- Print-ready PDF (vector)
- High-resolution PNG (selectable DPI)
- Editable PowerPoint (PPTX) with shapes

## What you need

- Windows 10/11 or macOS
- Python 3.13.9 or newer
- Microsoft Excel (or compatible .xlsx editor)

## Install (Windows)

1) Install Python 3.13.9+
- During installation, enable “Add Python to PATH”.

2) Open PowerShell and go to the project folder:
```powershell
cd path\to\roadmap_swimlane_app
```

3) Create and activate a virtual environment:
```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

4) Install dependencies:
```powershell
pip install -r requirements.txt
```

5) Run the app:
```powershell
streamlit run app.py
```

Your browser will open to the app UI.

## Install (macOS)

1) Install Python 3.13.9+

2) Open Terminal and go to the project folder:
```bash
cd /path/to/roadmap_swimlane_app
```

3) Create and activate a virtual environment:
```bash
python3 -m venv .venv
source .venv/bin/activate
```

4) Install dependencies:
```bash
pip install -r requirements.txt
```

5) Run the app:
```bash
streamlit run app.py
```

## Smoke test (2–3 minutes)

1) Launch the app (`streamlit run app.py`)
2) In the app, download:
   - `Roadmap_Input_TEMPLATE.xlsx` (blank template)
   - `Roadmap_Sample.xlsx` (pre-filled example)
3) Upload `Roadmap_Sample.xlsx`
   - Or click **Try sample now** in the app (no upload needed)
4) Confirm:
   - Live preview renders (blocks are stacked within a workstream when dates overlap)
   - “Today” line shows if it falls in range
5) Export:
   - PDF (open it; verify it’s crisp when zoomed)
   - PNG at 300 DPI (open it; verify text and lines remain sharp)
   - PPTX (open in PowerPoint; verify you can click and edit individual blocks)

Tip: if you make edits in the in-app grids, use the sidebar button **Download Excel** to save your edited workbook back to disk.

Note: when you upload a workbook, the app performs a **hard replace**: it clears all prior grid state and loads only the uploaded workbook data (no merging).

If you edited the *same file on disk* and re-selected it, some browsers won’t detect the change. Use **Upload → Having trouble? → Clear uploaded file selection** and upload again.

If you ever see old rows hanging around, use **Upload → Having trouble? → Force reload workbook (replace everything)**.

## Excel format (must match)

Workbook: `Roadmap_Input.xlsx`

Sheet 1: `Settings` (key/value)
- chart_title (string)
- chart_subtitle (string, optional)
- confidentiality_label (string, optional)
- overall_start_date (date)
- overall_end_date (date)
- timezone (default: America/Chicago)
- week_start_day (Mon or Sun; default Mon)
- time_granularity (timeline labels are auto-scaled by overall range: <4 months → weeks, 4–12 months → months, 13–24 months → quarters, >24 months → years + quarters)
- output_dpi (default 300; allowed 150/300/600)
- show_today_line (true/false)
- today_line_date (optional override; else uses “today” in timezone)
- page_size (A3 landscape default; allow A4 landscape)
- font_family (default Calibri or Arial; app auto-falls back if missing)

Sheet 2: `Workstreams`
Columns:
- workstream (required, unique)
- order (optional; numeric)
- color (optional; dropdown values: Auto, Blue, Orange, Green, Red, Purple, Brown, Pink, Gray, Olive, Cyan, Sky Blue, Peach. A hex like #1F77B4 is also accepted for legacy/power users.)

Sheet 3: `Tasks`
Columns:
- id (required; if blank, the app will auto-generate a stable ID)
- workstream (required; must match Workstreams.workstream)
- title (required)
- description (optional)
- start_date (required)
- end_date (required; can be same-day)
Optional:
- status (planned | in_progress | done | risk)
- owner (string)
- color_override (optional; dropdown values: Auto, Blue, Orange, Green, Red, Purple, Brown, Pink, Gray, Olive, Cyan, Sky Blue, Peach. A hex like #1F77B4 is also accepted for legacy/power users.)
- type (block | milestone)
- hyperlink (url)

## Overlap handling rule (important)

Tasks are treated as inclusive intervals for stacking:
- A task ending on 2025‑10‑01 overlaps a task starting on 2025‑10‑01.
- A task starting on 2025‑10‑02 can reuse the same sublane.

Algorithm: greedy interval partitioning (“first available sublane” by index), sorted by:
(start_date, end_date, id). This keeps stacking deterministic.

## Out-of-range tasks

If a task is partially outside the overall range:
- It will be clamped visually (warning shown).

If a task is completely outside the overall range:
- It is hidden by default.
- Toggle “Include out-of-range tasks” to show it anyway.

## Editing tips (in-app)

- Dates in the **Edit** grid are entered as text: use `YYYY-MM-DD` (example: `2026-01-15`).
- To delete rows in the grids, tick the **delete** checkbox on the row and click **Delete checked**.

## Troubleshooting

### “ModuleNotFoundError …”
- Confirm the virtual environment is activated.
- Re-run:
```bash
pip install -r requirements.txt
```

### Excel date parsing issues
Common causes:
- Dates entered as text (e.g., “10/1”) instead of an actual date cell
- Mixed formats in the same column

Fix:
- Re-enter the cell using Excel date formatting (yyyy-mm-dd recommended).
- Use the provided template; it sets date formats for you.

### Empty exports (blank PDF/PNG)
- Confirm overall_start_date / overall_end_date are present
- Confirm at least one task is within range (or toggle “Include out-of-range tasks”)
- Confirm every task has a valid workstream value that matches the Workstreams sheet

### Fonts not found
Calibri is not installed on many macOS machines by default.
- Set `font_family` to Arial in Settings, or leave it as Calibri and the app will fall back to a safe default (DejaVu Sans).

### Streamlit grid editing errors (column type mismatch)
If you previously saw an error like:
"configured column type text ... not compatible ... FLOAT"
this build fixes it by forcing optional text columns (like `hyperlink`) to a real string dtype before editing.

## Running tests (optional)

```bash
pytest -q
```

If you want a more "realistic" export stress test (generates multiple random roadmaps and exports each one), run:

```bash
python scripts/smoke_test.py
```

This script runs multiple randomized iterations and will print byte-size checks for preview PNG, high-res PNG, PDF, and PPTX.
