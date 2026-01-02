# Transformation Office Block & Gantt Creator Tool

A local, non-technical-friendly Streamlit app that turns a simple Excel workbook into an executive-ready roadmap chart (blocked diagram / swimlane Gantt) with clean overlap stacking, a smart timeline header, and one-click exports to:

- Print-ready PDF (vector)
- High-resolution PNG (DPI selectable)
- Editable PowerPoint (PPTX) with shapes you can move, resize, and edit

The visual style is designed for transformation office and executive roadmap communication, inspired by common “blocked diagram” decks used in leadership updates. :contentReference[oaicite:0]{index=0}

---

## Why this exists

Executives want to understand:
- What work is happening
- When it happens
- What’s running in parallel
- Where risk exists

This tool makes that obvious—fast—using a consistent, readable visual that is resilient to messy input.

---

## Key features

- Excel-first workflow (download template → fill → upload)
- Optional in-app editing (Workstreams + Tasks in editable grids)
- Hard replace on upload: when you upload a workbook, the app uses that workbook’s data as the source of truth
- Automatic overlap handling:
  - Tasks never visually overlap inside a workstream
  - Overlapping tasks are stacked into sublanes deterministically (greedy interval partitioning)
- Smart timeline header (auto selects the most readable time scale):
  - Under 4 months: months + weeks
  - 4–12 months: months
  - 13–24 months: quarters
  - Over 24 months: years + quarters
- Clear status visualization:
  - planned / in_progress / done / risk are visually distinct at a glance
- Accessible color workflow:
  - Color inputs are English names via dropdown (no hex codes required)
  - “Auto” assigns a clean palette automatically

---

## What you export

- PDF (vector): best for printing and high-quality sharing
- PNG (high-res): best for email, Slack, docs, and screenshots
- PPTX (editable): best for leadership decks—every block is an editable shape

---

## Repository layout

- `app.py`  
  Streamlit UI (Upload → Edit → Preview & Export + Instructions)

- `roadmap_models.py`  
  Pydantic models + validation rules

- `excel_io.py`  
  Excel template generation + workbook parsing/writing

- `scheduler.py`  
  Overlap stacking (interval partitioning) and layout preparation

- `renderer.py`  
  Matplotlib rendering for PDF/PNG + timeline header logic

- `pptx_exporter.py`  
  Editable PPTX export (PowerPoint shapes)

- `export.py`  
  Export orchestration + filenames + in-memory bytes

- `sample_inputs/`  
  - `Roadmap_Input_TEMPLATE.xlsx`
  - `Roadmap_Sample.xlsx`

- `tests/`  
  - `test_stacking.py`
  - `test_date_mapping.py`
  - other export and validation tests

- `scripts/`  
  - `smoke_test.py` (randomized multi-run export test)

---

## Requirements

- Python 3.13.9 or newer
- Windows or macOS
- Excel (.xlsx) editor (Microsoft Excel recommended)

---

## Install and run (Windows)

1) Open PowerShell and go to the repo folder:
```powershell
cd path\to\your\repo

