# Transformation Office — Block & Gantt Creator Tool

**Turn a simple spreadsheet into presentation-ready project visuals in under 60 seconds.**

No coding required. No complex project management software. Just your data in Excel and a browser.

---

## What This Tool Does

The Block & Gantt Creator Tool takes your project data from a basic Excel file and instantly generates two types of professional visualizations:

### 1. Swim Lane Gantt Chart
A timeline view where each team or workstream gets its own horizontal "lane." Tasks appear as color-coded bars stretched across their start and end dates. When tasks overlap within a team, the chart automatically stacks them so nothing gets hidden.

**Best for:** Showing the sequencing and timing of work across teams. Great for steering committee decks, status reviews, and planning workshops.

### 2. Space-Filling Block Diagram
Every work item is packed into a single slide, arranged by time and stacked to show how much work runs in parallel. The result is a dense, at-a-glance view of the entire program — similar to a heatmap of workload over time.

**Best for:** Executive summaries, one-pager overviews, and all-hands presentations where you need to show the full scope of work on a single slide.

---

## What You Need Before Starting

1. **A computer** (Mac or Windows)
2. **A web browser** (Chrome, Safari, Edge, or Firefox)
3. **Python 3** installed on your computer
   - **Mac:** Open Terminal (search for "Terminal" in Spotlight) and type `python3 --version`. If you see a version number, you're set. If not, download Python from [python.org/downloads](https://www.python.org/downloads/) — click the big yellow "Download" button, open the installer, and follow the prompts.
   - **Windows:** Open Command Prompt (search for "cmd" in the Start menu) and type `python --version`. If you don't have it, download from [python.org/downloads](https://www.python.org/downloads/) — **important:** during installation, check the box that says "Add Python to PATH."
4. **Your project data** in Excel format (.xlsx) — or just use the built-in sample data to try it out first.

---

## How to Set Up the Tool (One-Time)

### Step 1 — Download and Unzip

Download the `roadmap_pro.zip` file and unzip it. You will see a folder called `roadmap_pro` with these files inside:

```
roadmap_pro/
  app.py                  ← The main application
  models.py               ← Data definitions
  gantt_renderer.py       ← Gantt chart engine
  block_renderer.py       ← Block diagram engine
  pptx_export.py          ← PowerPoint export engine
  excel_io.py             ← Excel reader
  requirements.txt        ← List of software dependencies
  run.sh                  ← Quick-start script (Mac/Linux)
  sample_data/
    sample_roadmap.xlsx    ← Sample data to try the tool
```

### Step 2 — Run the Tool

#### On Mac

1. Open **Terminal** (search for "Terminal" in Spotlight, or find it in Applications → Utilities).
2. Drag the `roadmap_pro` folder into the Terminal window. This types the folder path for you. Then add `cd ` before it and press Enter. It will look something like:
   ```
   cd /Users/yourname/Downloads/roadmap_pro
   ```
3. Type the following and press Enter:
   ```
   bash run.sh
   ```
4. Wait about 30 seconds. The tool will install what it needs and then show:
   ```
   The app will open in your browser at http://localhost:8501
   ```
5. Your browser will open automatically. If it doesn't, open your browser and go to **http://localhost:8501**

#### On Windows

1. Open **Command Prompt** (search for "cmd" in the Start menu).
2. Navigate to the folder. If you unzipped to your Downloads folder, type:
   ```
   cd %USERPROFILE%\Downloads\roadmap_pro
   ```
3. Install the required packages (one-time only):
   ```
   pip install -r requirements.txt
   ```
4. Start the app:
   ```
   streamlit run app.py
   ```
5. Your browser will open automatically. If it doesn't, go to **http://localhost:8501**

#### What "Running" Looks Like

- The Terminal or Command Prompt window will show some log text — this is normal. **Do not close this window** while you are using the tool.
- To stop the tool when you're done, go back to the Terminal/Command Prompt and press **Ctrl + C**.
- To start it again later, repeat the steps above (you won't need to reinstall).

---

## How to Use the Tool

### Try It First With Sample Data

When the tool opens in your browser, you'll see the home screen. Click the blue **"Load Sample Data"** button to instantly see both chart types with a realistic product launch roadmap. This is a great way to explore the features before using your own data.

### Use Your Own Data

#### 1. Download the Template

Click **"Download Excel Template"** in the left sidebar. This gives you a pre-filled `.xlsx` file that shows the exact format the tool expects. Open it in Excel or Google Sheets.

#### 2. Replace the Sample Rows With Your Data

Delete the sample rows and add your own. Here is what each column does:

| Column | Required? | What to Enter | Example |
|---|---|---|---|
| **Title** | Yes | Name of the task, initiative, or deliverable | Website Redesign |
| **Start Date** | Yes | When work begins | 2025-01-15 |
| **End Date** | Yes | When work ends | 2025-03-31 |
| **Category** | No | Team, workstream, or department — this is what creates the color-coded groups | Marketing |
| **Description** | No | A short description (shows inside the blocks) | Full site overhaul with new branding |
| **Status** | No | One of: `planned`, `in_progress`, `done`, `at_risk` | in_progress |
| **Owner** | No | Person responsible | Sarah |
| **Label** | No | Short identifier shown on the block (e.g., "D1", "MVP") | M2 |

**Tips:**
- Only the first three columns (Title, Start Date, End Date) are required. Everything else is optional.
- Column names are flexible — "Name" works instead of "Title", "Team" works instead of "Category", "Due" works instead of "End Date", etc.
- Dates can be in almost any format: `2025-01-15`, `01/15/2025`, `1/15/25`, or just use Excel's date picker.
- If you want a milestone (a single point in time, not a range), set the Start Date and End Date to the same day.

#### 3. Upload Your File

Drag your `.xlsx` file into the upload area in the left sidebar, or click **"Browse files"** to find it. The tool reads your file instantly and shows you a summary of what it found.

#### 4. Customize the Look

Once your data is loaded, the sidebar shows these options:

- **Title & Subtitle** — Change the heading that appears on your chart.
- **Color Palette** — Choose from 6 built-in color schemes:
  - **Ocean** — Cool blues and teals
  - **Sunset** — Warm reds, oranges, and greens
  - **Forest** — Natural greens
  - **Corporate** — Professional blues and purples
  - **Vibrant** — Bold, high-contrast colors (great for presentations)
  - **Monochrome** — Shades of gray (clean and formal)
- **Show today line** — Draws a vertical line marking today's date on the chart.
- **Show status indicators** — Adds colored dots showing whether each item is planned, in progress, done, or at risk.
- **Show legend** — Displays a color key for the categories.
- **Date Range** — Adjust the start and end dates if you want to zoom in on a specific timeframe.

#### 5. Export Your Charts

Each chart tab (Gantt Chart and Block Diagram) has three download buttons:

| Format | What You Get | When to Use It |
|---|---|---|
| **PowerPoint (.pptx)** | Editable shapes on a widescreen slide. Every bar, label, and header is a native PowerPoint shape you can move, resize, and recolor. | When you need to make tweaks, add annotations, or incorporate the chart into an existing deck. |
| **PDF** | Vector-quality document. Scales to any size without getting blurry. | When you need a print-quality document or want to share a file that looks the same on every computer. |
| **PNG (300 DPI)** | High-resolution image. | When you need to paste the chart into an email, a Word document, Slack, or any tool that accepts images. |

---

## Understanding the Charts

### Gantt Chart Layout

```
                    Jan        Feb        Mar        Apr        May
                     │          │          │          │          │
  Marketing    ██████████████████
               Brand Refresh     ████████████████████████████
                                 Website Redesign
  ─────────────────────────────────────────────────────────────────
  Product                  ████████████████
                           UX/UI Design      ██████████████████
                                             Beta Testing
  ─────────────────────────────────────────────────────────────────
  Engineering       ████████████████████████████████████████████
                    Core Platform Build
```

- Each **row group** (Marketing, Product, Engineering) is a "swim lane" based on your Category column.
- Each **bar** is one work item, spanning from its Start Date to End Date.
- When items overlap within a category, they stack into **sublanes** so everything remains visible.
- The **today line** (if enabled) shows a dashed vertical line at today's date.

### Block Diagram Layout

```
  ┌──────────────────────────────────────────────────────────────┐
  │  Jan     │  Feb     │  Mar     │  Apr     │  May     │  Jun  │
  ├──────────┼──────────┼──────────┼──────────┼──────────┼───────┤
  │ Brand Refresh       │ Core Platform Build                    │
  │ Customer Research   │ UX/UI Design        │ Pilot Customers  │
  │ Vendor Selection │ Sales Playbook │ Mobile App Dev           │
  │ Hiring Plan Execution            │ Partner Onboarding       │
  │    │ Website Redesign  │ Product Launch Campaign │ Security  │
  └──────────────────────────────────────────────────────────────┘
```

- Every work item is packed into rows, sorted to minimize empty space.
- The result is a **dense, single-slide overview** of the entire program.
- Taller stacks mean more work happening in parallel during that time period.

---

## Troubleshooting

### "I can't find Terminal" (Mac)
Press **Cmd + Space** to open Spotlight, then type **Terminal** and press Enter.

### "Python is not recognized" (Windows)
You need to install Python. Go to [python.org/downloads](https://www.python.org/downloads/), download the installer, and — this is important — **check the box that says "Add Python to PATH"** during installation. Then close and reopen Command Prompt.

### "pip is not recognized"
Try using `pip3` instead of `pip`:
```
pip3 install -r requirements.txt
```

### "streamlit is not recognized"
After installing requirements, you may need to use the full path:
```
python3 -m streamlit run app.py
```

### The app opens but shows a blank page
- Try a different browser (Chrome works best).
- Make sure no other program is using port 8501. If so, the tool will show an error in Terminal — you can close that other program or run on a different port:
  ```
  streamlit run app.py --server.port 8502
  ```

### My file won't upload
- The file must be `.xlsx` format (not `.csv`, `.xls`, or `.numbers`).
- If you're using Google Sheets, download as Excel first: **File → Download → Microsoft Excel (.xlsx)**.
- File must be under 200 MB.

### "Missing required column" error
- Make sure your file has columns named **Title** (or Name/Task/Item), **Start Date** (or Start/Begin), and **End Date** (or End/Finish/Due).
- Column headers must be in **Row 1** of the spreadsheet.
- Don't merge cells in the header row.

### Dates aren't reading correctly
- Make sure date cells are formatted as **Date** in Excel, not as plain text.
- Supported formats: `2025-01-15`, `01/15/2025`, `1/15/25`, `15/01/2025`.
- Remove any blank rows — don't leave empty rows between data rows.
- If a Start Date is accidentally after the End Date, the tool will swap them automatically and show a warning.

### The chart looks squished or has too much empty space
- Use the **Date Range** controls in the sidebar to zoom into the relevant time period.
- For the Block Diagram, switch between **16:9** and **4:3** aspect ratio to see which fits your data better.

### Category colors are confusing
- Try a different **Color Palette** — "Vibrant" and "Corporate" tend to have the most distinct colors.
- Make sure your category names are spelled consistently (e.g., "Marketing" and "marketing" will be treated as two separate groups).

---

## Frequently Asked Questions

**Can I use Google Sheets instead of Excel?**
Yes. Create your spreadsheet in Google Sheets, then download it as Excel: **File → Download → Microsoft Excel (.xlsx)**. Upload that downloaded file.

**What if I only have three columns (Title, Start, End)?**
That's perfectly fine. The tool will assign all items to a single "General" category and use default settings for everything else.

**Can I show milestones (single-day events)?**
Yes. Set the Start Date and End Date to the same day. The tool will automatically display it as a diamond-shaped milestone marker.

**Can I edit the PowerPoint export?**
Yes. Every element in the exported PowerPoint — every bar, label, header, and shape — is a native, editable PowerPoint object. You can move, resize, recolor, or delete anything.

**How many items can it handle?**
The tool works well with up to several hundred items. For very large programs (500+), the Block Diagram may become dense — consider splitting into phases or filtering by team.

**Is my data sent anywhere?**
No. The tool runs entirely on your computer. Your data never leaves your machine. There is no server, no cloud, no login.

**Can I change the fonts or colors in the exported chart?**
In the PNG and PDF exports, the chart is rendered as-is. In the **PowerPoint export**, you can change fonts, colors, sizes, and positions of every element because they are native editable shapes.

**Can multiple people use this at the same time?**
Each person needs to run the tool on their own computer. There is no shared server or collaboration feature — it's a personal productivity tool.

---

## Quick Reference

| Action | How |
|---|---|
| Start the tool | `bash run.sh` (Mac) or `streamlit run app.py` (Windows) |
| Stop the tool | Press **Ctrl + C** in Terminal / Command Prompt |
| Try with sample data | Click **"Load Sample Data"** on the home screen |
| Get the Excel template | Click **"Download Excel Template"** in the sidebar |
| Switch between charts | Click the **Gantt Chart** or **Block Diagram** tab |
| Change colors | Select a different **Color Palette** in the sidebar |
| Export a chart | Click **Download PNG**, **Download PDF**, or **Download PowerPoint** below the chart |
| Start over | Click **"Clear Data"** in the top right corner of the visualization view |

---

## File Contents

| File | Purpose |
|---|---|
| `app.py` | Main application — the user interface and page layout |
| `models.py` | Data definitions — work items, chart settings, color palettes |
| `gantt_renderer.py` | Renders the Swim Lane Gantt Chart (PNG and PDF) |
| `block_renderer.py` | Renders the Space-Filling Block Diagram (PNG and PDF) |
| `pptx_export.py` | Creates editable PowerPoint exports for both chart types |
| `excel_io.py` | Reads your Excel file and creates the downloadable template |
| `requirements.txt` | Lists the Python packages the tool needs |
| `run.sh` | Quick-start script for Mac/Linux |
| `sample_data/sample_roadmap.xlsx` | Sample data with 25 items across 5 teams |

---

*Built for the Transformation Office. No coding required.*
