# Transformation Office — Block & Gantt Creator Tool

> **Version 1.0** — General Release

**Turn a simple spreadsheet into presentation-ready project visualizations in under 60 seconds.**

No coding required. No complex project management software. Just your data in Excel and a web browser.

---

## What This Tool Does

The Block & Gantt Creator Tool takes your project data from a basic Excel file and instantly generates two types of professional visualizations:

### 1. Swim Lane Gantt Chart

A timeline view where each team or workstream gets its own horizontal "lane." Tasks appear as color-coded bars stretched across their start and end dates. When tasks overlap within a team, the chart automatically stacks them so nothing gets hidden.

**Best for:** Showing the sequencing and timing of work across teams. Ideal for steering committee decks, status reviews, and planning workshops.

### 2. Space-Filling Block Diagram

Every work item is packed into a single slide, arranged by time and stacked to show how much work runs in parallel. The result is a dense, at-a-glance view of the entire program — similar to a heatmap of workload over time.

**Best for:** Executive summaries, one-pager overviews, and all-hands presentations where you need to show the full scope of work on a single slide.

### Export Formats

Both chart types can be exported in three formats:

| Format | Description | Best Used For |
|---|---|---|
| **PowerPoint (.pptx)** | Fully editable native shapes — every bar, label, and header can be moved, resized, and recolored | Incorporating into existing decks, making adjustments, adding annotations |
| **PDF** | Vector-quality document that scales to any size without loss of clarity | Print-ready documents, sharing via email, consistent rendering across devices |
| **PNG (300 DPI)** | High-resolution image | Pasting into emails, Word documents, Slack messages, or any application that accepts images |

---

## Prerequisites

Before running the tool, ensure you have the following:

1. **A computer** — Mac or Windows
2. **A web browser** — Chrome (recommended), Safari, Edge, or Firefox
3. **Python 3** — the programming language that powers the tool behind the scenes

   **Check if Python is already installed:**

   | Platform | How to Check | What to Look For |
   |---|---|---|
   | **Mac** | Open Terminal (Cmd + Space, type "Terminal") and run `python3 --version` | A version number like `Python 3.11.5` means you're set |
   | **Windows** | Open Command Prompt (search for "cmd") and run `python --version` | A version number means you're set |

   **If Python is not installed:** Download it from [python.org/downloads](https://www.python.org/downloads/). On Windows, **check the box labeled "Add Python to PATH"** during installation — this is critical.

4. **Your project data** in Excel format (.xlsx) — or use the built-in sample data to explore the tool first

---

## Setup & Installation

### Step 1 — Download and Extract

Download the tool package and unzip it. You will see a folder with the following structure:

```
block_gantt_creator/
    app.py                  ← Main application (user interface)
    models.py               ← Data definitions and color palettes
    gantt_renderer.py       ← Gantt chart rendering engine
    block_renderer.py       ← Block diagram rendering engine
    pptx_export.py          ← PowerPoint export engine
    excel_io.py             ← Excel file reader and template generator
    requirements.txt        ← Software dependencies
    run.sh                  ← Quick-start script (Mac/Linux)
    README.md               ← This document
    sample_data/
        sample_roadmap.xlsx ← Sample dataset (25 items, 5 workstreams)
```

### Step 2 — Launch the Tool

#### Mac / Linux

1. Open **Terminal** (press Cmd + Space, type "Terminal", press Enter).
2. Navigate to the tool folder. You can do this by typing `cd ` (with a space after it) and then dragging the folder from Finder into the Terminal window. Press Enter.
   ```
   cd /Users/yourname/Downloads/block_gantt_creator
   ```
3. Run the start script:
   ```
   bash run.sh
   ```
4. The tool will install its dependencies (first time only) and open in your default browser.

#### Windows

1. Open **Command Prompt** (press the Windows key, type "cmd", press Enter).
2. Navigate to the tool folder:
   ```
   cd %USERPROFILE%\Downloads\block_gantt_creator
   ```
3. Install dependencies (first time only):
   ```
   pip install -r requirements.txt
   ```
4. Start the application:
   ```
   streamlit run app.py
   ```

#### What to Expect

- A Terminal or Command Prompt window will remain open with log output — **do not close this window** while using the tool.
- Your browser will open automatically to **http://localhost:8501**. If it doesn't, open your browser and navigate to that address manually.
- To stop the tool, return to Terminal/Command Prompt and press **Ctrl + C**.
- To restart later, simply repeat the launch command. Dependencies only install once.

---

## User Guide

### Getting Started — Sample Data

When the tool opens, you will see the home screen. Click **"Load Sample Data"** to instantly explore both chart types with a pre-built product launch roadmap (25 items across 5 workstreams). This is the fastest way to understand the tool's capabilities before using your own data.

### Using Your Own Data

#### Step 1 — Download the Template

Click **"Download Excel Template"** in the left sidebar. This provides a pre-formatted `.xlsx` file with sample data and an instructions sheet. Open it in Microsoft Excel or Google Sheets.

#### Step 2 — Enter Your Data

Replace the sample rows with your own project data. The following columns are available:

| Column | Required | Description | Example |
|---|---|---|---|
| **Title** | Yes | Name of the initiative, task, or deliverable | Website Redesign |
| **Start Date** | Yes | When work begins | 2025-01-15 |
| **End Date** | Yes | Target completion date | 2025-03-31 |
| **Category** | No | Workstream, team, or department — determines color coding and swim lane grouping | Marketing |
| **Description** | No | Brief description displayed inside chart blocks | Complete site overhaul with new branding |
| **Status** | No | Current status: `planned`, `in_progress`, `done`, or `at_risk` | in_progress |
| **Owner** | No | Responsible individual or lead | Sarah |
| **Label** | No | Short identifier displayed on the block (e.g., "D1", "MVP", "Phase 2") | M2 |

**Key points:**
- Only three columns are required — Title, Start Date, and End Date. All others are optional.
- Column names are flexible — "Name" works for "Title", "Team" works for "Category", "Due" works for "End Date", etc.
- Dates can be entered in most standard formats: `2025-01-15`, `01/15/2025`, `1/15/25`.
- To represent a milestone (a single point in time), set Start Date and End Date to the same day.

#### Step 3 — Upload

Drag your `.xlsx` file into the upload area in the left sidebar, or click **"Browse files"** to select it. The tool reads and validates your file instantly.

#### Step 4 — Customize

Once data is loaded, the sidebar displays customization options:

| Setting | What It Does |
|---|---|
| **Title & Subtitle** | Set the heading displayed on your chart |
| **Color Palette** | Choose from 6 schemes: Ocean, Sunset, Forest, Corporate, Vibrant, Monochrome |
| **Show today line** | Toggle a vertical marker for today's date |
| **Show status indicators** | Display colored status markers (Planned, In Progress, Done, At Risk) |
| **Show legend** | Display the color key for categories |
| **Date Range** | Adjust the visible start and end dates to focus on a specific period |

#### Step 5 — Export

Select the **Gantt Chart** or **Block Diagram** tab, then click one of the three export buttons below the chart:

- **Download PNG (300 DPI)** — high-resolution image
- **Download PDF** — vector-quality document
- **Download PowerPoint** — fully editable native shapes

---

## Troubleshooting

### Setup Issues

| Problem | Solution |
|---|---|
| "I can't find Terminal" (Mac) | Press **Cmd + Space** to open Spotlight, type **Terminal**, press Enter |
| "Python is not recognized" (Windows) | Install Python from [python.org/downloads](https://www.python.org/downloads/) — **check "Add Python to PATH"** during installation. Close and reopen Command Prompt afterward. |
| "pip is not recognized" | Try `pip3` instead of `pip`, or run `python3 -m pip install -r requirements.txt` |
| "streamlit is not recognized" | Run `python3 -m streamlit run app.py` instead |
| Tool opens but shows a blank page | Try Chrome. If the port is in use, run `streamlit run app.py --server.port 8502` |

### File Upload Issues

| Problem | Solution |
|---|---|
| File won't upload | Ensure the file is `.xlsx` format — not `.csv`, `.xls`, or `.numbers`. If using Google Sheets, export via *File > Download > Microsoft Excel (.xlsx)* |
| "Missing required column" error | Verify your file has **Title**, **Start Date**, and **End Date** columns in Row 1. No merged cells in the header row. |
| Date parsing errors | Ensure date cells are formatted as **Date** in Excel, not as text. Remove blank rows between data rows. |

### Chart Display Issues

| Problem | Solution |
|---|---|
| Chart looks squished or has too much empty space | Adjust the **Date Range** in the sidebar to focus on the relevant time period |
| Category colors are hard to distinguish | Switch to the **Vibrant** or **Corporate** palette — these have the highest contrast |
| Block Diagram blocks appear too narrow | Try the **4:3** aspect ratio instead of 16:9 |
| Categories appear duplicated | Ensure consistent spelling and capitalization (e.g., "Marketing" and "marketing" are treated as separate categories) |

---

## Frequently Asked Questions

**Can I use Google Sheets?**
Yes. Build your spreadsheet in Google Sheets, then download it as Excel: *File > Download > Microsoft Excel (.xlsx)*. Upload the downloaded file.

**What if I only have three columns?**
That works perfectly. All items will be assigned to a default "General" category with standard settings applied.

**Can I show milestones?**
Yes. Set Start Date and End Date to the same day. The tool displays these as diamond-shaped milestone markers on the Gantt chart.

**Can I edit the PowerPoint output?**
Yes. Every element — bars, labels, headers, shapes — is a native PowerPoint object. You can move, resize, recolor, delete, or add annotations freely.

**How many items can it handle?**
The tool performs well with up to several hundred items. For very large programs (500+), consider filtering by workstream or splitting into phases to keep the visualization readable.

**Is my data sent anywhere?**
No. The tool runs entirely on your local computer. Your data never leaves your machine. There is no server, cloud service, or login required.

**Can multiple people use this?**
Each person runs the tool independently on their own computer. There is no shared server or real-time collaboration — it is a personal productivity tool.

---

## Quick Reference

| Action | How |
|---|---|
| Launch the tool | `bash run.sh` (Mac/Linux) or `streamlit run app.py` (Windows) |
| Stop the tool | Press **Ctrl + C** in Terminal / Command Prompt |
| Explore with sample data | Click **"Load Sample Data"** on the home screen |
| Download the Excel template | Click **"Download Excel Template"** in the sidebar |
| Switch chart types | Click the **Gantt Chart** or **Block Diagram** tab |
| Change the color scheme | Select a different palette under **Chart Settings** in the sidebar |
| Export a chart | Click **Download PNG**, **Download PDF**, or **Download PowerPoint** below the chart |
| Start over with new data | Click **"Clear Data"** in the upper right corner |

---

---

## Changelog

### v1.0 — General Release
- Removed Sequencing Diagram — the tool now focuses on two core visualization modes: **Gantt Chart** and **Block Diagram**
- Upgraded Block Diagram to a true space-filling mosaic layout with proportional sizing and improved text rendering
- Updated version from 0.9 Beta to 1.0
- Cleaned up unused feature branches and codebase

### v0.9 Beta — Early Access Preview
- Initial release with Gantt Chart, Block Diagram, and Sequencing Diagram
- Excel upload with flexible column matching and multi-format date parsing
- Export to PowerPoint (editable shapes), PDF, and PNG (300 DPI)
- 6 color palettes: Ocean, Sunset, Forest, Corporate, Vibrant, Monochrome
- Sample data and Excel template generation

---

<div align="center">

**Transformation Office — Block & Gantt Creator Tool**

Version 1.0

*For internal use. Please direct feedback and feature requests to the Transformation Office team.*

</div>
