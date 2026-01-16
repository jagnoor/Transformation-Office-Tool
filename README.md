# Transformation Office — Block & Gantt Creator Tool

A lightweight, executive-ready roadmap generator that converts a single Excel workbook into a clean blocked-diagram / swimlane Gantt view of organizational work. Exports: PDF (print-ready, vector), PNG (high-resolution), and PPTX (editable PowerPoint shapes).

## Quick summary

* Purpose: let transformation offices and workstream leaders treat Excel as the single source of truth and produce polished, editable executive visuals without manual redraws.
* Inputs: one Excel workbook (template included).
* Outputs: PDF (vector), PNG (high-res), PPTX (editable shapes).
* Primary audience: transformation offices, program leads, PMOs.


## Features

* Left-to-right timeline spanning user-defined overall start/end dates.
* Swimlane workstreams; tasks render as rounded blocks, milestones as diamonds.
* Overlap stacking: overlapping tasks auto-assign sublanes so blocks never visually overlap.
* Status styling across exports: planned / in_progress / done / risk.
* Smart timeline header: auto-adjusts (weeks, months, quarters, years) by date range.
* Exports: high-quality PDF (vector), PNG (selectable DPI), and editable PPTX.

## What problem this solves

* Eliminates manual redrawing in PowerPoint.
* Keeps Excel as the canonical source of truth.
* Produces leadership-friendly visuals and an editable PPTX, not screenshots.
* Reduces multiple versions of truth and manual layout errors.

## Outputs (what you get)

* PDF (vector): crisp for print and large displays.
* PNG (high-res): quick sharing, embeds in docs and tickets.
* PPTX (editable): move/resize/edit blocks in PowerPoint; hyperlinks preserved.

## How the visual works

* Timeline: left-to-right from overall_start_date to overall_end_date.
* Swimlanes: workstreams grouped in rows; optional ordering.
* Tasks: rounded blocks from start_date to end_date.
* Milestones: diamond placed on a single date.
* Stacking: overlapping tasks in a workstream stack into sublanes.

## Smart timeline header (automatic)

* < 4 months: months + weeks (two-row header)
* 4–12 months: months
* 13–24 months: quarters
* > 24 months: years + quarters (two-row header)

## Status styling (applied across all exports)

* planned: neutral
* in_progress: accent (e.g., dashed edge + status stripe)
* done: de-emphasized (lighter fill)
* risk: strong callout (red stripe/edge)

## Recommended workflow (best practice)

1. Download the Excel template.
2. Fill settings → workstreams → tasks in Excel.
3. Upload workbook to the app.
4. Preview, then export (PDF/PNG/PPTX).
5. For updates: update Excel and re-upload; use in-app grid edits only for quick fixes.

## Quick start — hosted app

Open the hosted app (if available):
`https://transformation-office-tool.streamlit.app/`
Upload the sample Excel in `sample_inputs/` to test end-to-end. For confidential roadmaps, run locally.

## Install and run locally — Windows

1. Install Python 3.13.9+ and enable “Add Python to PATH”.
2. Clone or download the repo and `cd` into it:

```powershell
cd C:\Users\<you>\Downloads\transformation-office-tool
```

3. Create venv:

```powershell
python -m venv .venv
```

4. Activate:

```powershell
.\.venv\Scripts\Activate.ps1
```

5. Install deps:

```powershell
pip install -r requirements.txt
```

6. Run the app:

```powershell
streamlit run app.py
```

## Install and run locally — macOS

1. Install Python 3.13.9+.
2. Clone or download and `cd` into repo:

```bash
cd ~/Downloads/transformation-office-tool
```

3. Create venv:

```bash
python3 -m venv .venv
```

4. Activate:

```bash
source .venv/bin/activate
```

5. Install deps:

```bash
pip install -r requirements.txt
```

6. Run:

```bash
streamlit run app.py
```

## How to use — step-by-step (non-technical)

### Step 1 — Download template

From the app Upload tab: “Download Excel template”.

### Step 2 — Fill the template (3 sheets)

Sheet 1: Settings (key/value) — required: `chart_title`, `overall_start_date`, `overall_end_date`. Common: `chart_subtitle`, `timezone` (default America/Chicago), `page_size` (A3 landscape recommended), `output_dpi` (300 recommended), `show_today_line`.

Sheet 2: Workstreams — columns:

* `workstream` (required, unique)
* `order` (optional integer; lower appears higher)
* `color` (dropdown; Auto recommended)

Sheet 3: Tasks — required:

* `workstream` (must match Workstreams)
* `title`
* `start_date`
* `end_date`
  Optional: `description`, `status` (planned|in_progress|done|risk), `owner`, `color_override`, `type` (block|milestone), `hyperlink`.

Tips:

* Use `YYYY-MM-DD` for typed dates.
* Keep titles short (15–40 chars).
* Milestones = key single-day events.

### Step 3 — Upload workbook

Upload the completed workbook via Upload tab. If browser caching shows old content, use “clear file selection” and re-upload.

### Step 4 — Review & export

Use Preview. Export: PDF / PNG / PPTX.

## Template cheat sheet (copy/paste)

Settings: `chart_title`, `overall_start_date`, `overall_end_date`, `output_dpi=300`, `page_size=A3 landscape`.
Workstreams: short unique names; color: Auto.
Tasks: `title` short, dates valid, `status` from allowed values, `type` block|milestone, `hyperlink` optional.

## Common mistakes

* Workstream names mismatch between sheets.
* end_date earlier than start_date.
* Dates entered as free text (use date cells or `YYYY-MM-DD`).

## PowerPoint editing tips (PPTX export)

* Use PPT guides and gridlines.
* Align and distribute shapes to keep spacing consistent.
* Group blocks with their status stripe.
* Make final title tweaks in PPTX; content changes should come from Excel + re-export.

## Lightweight governance

* Weekly: workstream owners update Excel.
* Biweekly: transformation office reviews risks and overlaps.
* Monthly: export leadership-ready deck + narrative.
  Ownership model:
* Workstream owner: task data and status.
* Transformation office: structure, ordering, date window.
* One designated publisher: final export for leadership.

## Versioning and filenames

Save exports with date in filename, e.g. `Roadmap_2026-02-01.pptx`. Keep Excel as canonical record.

## Troubleshooting

* Module not found: ensure venv activated; reinstall with `pip install -r requirements.txt`.
* Dates won’t parse: use Excel date cells or `YYYY-MM-DD`; ensure `end_date >= start_date`.
* Upload cached: use “clear file selection” and re-upload.
* Fonts differ: exports will still work; pick a cross-platform font in Settings (Arial recommended).

## Repo structure

```text
.
├─ app.py
├─ roadmap_models.py
├─ excel_io.py
├─ scheduler.py
├─ date_utils.py
├─ renderer.py
├─ pptx_export.py
├─ export.py
├─ requirements.txt
├─ sample_inputs/
│  ├─ Roadmap_Input_TEMPLATE.xlsx
│  └─ Roadmap_Sample.xlsx
├─ scripts/
│  └─ smoke_test.py
└─ tests/
   ├─ conftest.py
   ├─ test_color_mapping.py
   ├─ test_date_mapping.py
   ├─ test_excel_roundtrip.py
   ├─ test_export_smoke.py
   ├─ test_stacking.py
   └─ test_timeline_mode.py
```

File summaries

* `app.py`: Streamlit UI (upload, edit, preview, export).
* `roadmap_models.py`: Pydantic models and validation.
* `excel_io.py`: Excel read/write and template generation.
* `scheduler.py`: overlap stacking / sublane assignment.
* `date_utils.py`: date → X-position mapping.
* `renderer.py`: matplotlib rendering for PDF/PNG/preview.
* `pptx_export.py`: python-pptx editable export.
* `export.py`: export orchestration.
* `scripts/smoke_test.py`: end-to-end smoke tests.
* `tests/`: unit and integration tests.

## Tests (maintainers)

Run unit tests:

```bash
pytest -q
```

Run full smoke test:

```bash
python scripts/smoke_test.py
```



