"""
Roadmap Pro — World-class project visualization tool.

Creates beautiful Gantt charts and space-filling block diagrams
from simple Excel input. Exports to PowerPoint, PDF, and PNG.
"""
import io
import traceback

import pandas as pd
import streamlit as st
from datetime import date, timedelta

from models import WorkItem, ChartConfig, PALETTES, DEFAULT_PALETTE, STATUS_LABELS
from excel_io import read_excel, create_template_bytes
from gantt_renderer import render_gantt, render_gantt_pdf
from block_renderer import render_block_diagram, render_block_pdf
from pptx_export import export_gantt_pptx, export_block_pptx


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Roadmap Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F8FAFC; }

    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #334155 100%);
        color: white;
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.15);
    }
    .main-header h1 {
        font-size: 2rem; font-weight: 700; margin: 0; letter-spacing: -0.02em;
    }
    .main-header p {
        color: #94A3B8; margin: 0.3rem 0 0 0; font-size: 1rem;
    }

    .card {
        background: white; border-radius: 12px; padding: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06); border: 1px solid #E2E8F0;
        margin-bottom: 1rem; height: 100%;
    }
    .card h3 { margin-top: 0; color: #0F172A; }
    .card p { color: #64748B; font-size: 0.95rem; line-height: 1.5; }

    .feature-icon {
        width: 48px; height: 48px; border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.5rem; margin-bottom: 0.75rem;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 8px 20px; font-weight: 600;
    }

    .stDownloadButton > button {
        border-radius: 8px; font-weight: 600; padding: 0.5rem 1.5rem;
    }

    section[data-testid="stSidebar"] {
        background-color: #FFFFFF; border-right: 1px solid #E2E8F0;
    }
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #0F172A; font-size: 0.85rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.05em;
    }

    [data-testid="stMetric"] {
        background: white; border-radius: 10px; padding: 1rem;
        border: 1px solid #E2E8F0; box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .step-num {
        display: inline-flex; align-items: center; justify-content: center;
        background: #2563EB; color: white; border-radius: 50%;
        width: 32px; height: 32px; font-weight: 700; font-size: 0.9rem;
        margin-right: 12px; flex-shrink: 0;
    }
    .step-row {
        display: flex; align-items: center; padding: 0.6rem 0;
        border-bottom: 1px solid #F1F5F9;
    }
    .step-row:last-child { border-bottom: none; }

    .help-box {
        background: #FFF7ED; border: 1px solid #FED7AA; border-radius: 10px;
        padding: 1.25rem; margin-top: 1rem;
    }
    .help-box h4 { margin: 0 0 0.5rem 0; color: #9A3412; }
    .help-box p, .help-box li { color: #78350F; font-size: 0.9rem; }

    .column-table {
        width: 100%; border-collapse: collapse; margin: 0.75rem 0;
    }
    .column-table th {
        background: #F1F5F9; padding: 0.5rem 0.75rem; text-align: left;
        font-size: 0.85rem; color: #475569; border-bottom: 2px solid #E2E8F0;
    }
    .column-table td {
        padding: 0.5rem 0.75rem; border-bottom: 1px solid #F1F5F9;
        font-size: 0.85rem; color: #334155;
    }
    .required { color: #DC2626; font-weight: 600; }
    .optional { color: #059669; }
</style>
""", unsafe_allow_html=True)


# ── Session state init ───────────────────────────────────────────────────────
if "items" not in st.session_state:
    st.session_state["items"] = None
if "config" not in st.session_state:
    st.session_state["config"] = ChartConfig()
if "warnings" not in st.session_state:
    st.session_state["warnings"] = []
if "load_error" not in st.session_state:
    st.session_state["load_error"] = None


# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Roadmap Pro</h1>
    <p>Create stunning project visualizations from simple spreadsheets</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📥 Step 1 — Get Template")
    template_bytes = create_template_bytes()
    st.download_button(
        label="Download Excel Template",
        data=template_bytes,
        file_name="roadmap_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Pre-filled template with sample data and instructions",
    )

    st.markdown("### 📤 Step 2 — Upload Your File")
    uploaded_file = st.file_uploader(
        "Drop your Excel file here",
        type=["xlsx", "xls"],
        help="Excel file with columns: Title, Start Date, End Date, Category (optional)",
    )

    if uploaded_file:
        try:
            loaded_items, loaded_config, loaded_warnings = read_excel(uploaded_file)
            if not loaded_items:
                st.session_state["load_error"] = "No work items found. Make sure your file has data rows with Title, Start Date, and End Date columns."
                st.session_state["items"] = None
            else:
                st.session_state["items"] = loaded_items
                st.session_state["config"] = loaded_config
                st.session_state["warnings"] = loaded_warnings
                st.session_state["load_error"] = None

                st.success(f"Loaded **{len(loaded_items)}** work items across **{len(set(it.category for it in loaded_items))}** categories")

                if loaded_warnings:
                    with st.expander(f"⚠️ {len(loaded_warnings)} warning(s)", expanded=False):
                        for w in loaded_warnings:
                            st.caption(w)
        except Exception as e:
            st.session_state["load_error"] = str(e)
            st.session_state["items"] = None
            st.error(f"Could not read file — see troubleshooting below")

    # Show load error details in sidebar
    if st.session_state["load_error"]:
        with st.expander("🔍 Error Details", expanded=True):
            st.code(st.session_state["load_error"], language=None)

    st.divider()

    # Settings — only when data is loaded
    if st.session_state["items"] is not None:
        st.markdown("### 🎨 Chart Settings")

        config = st.session_state["config"]
        config.title = st.text_input("Title", value=config.title)
        config.subtitle = st.text_input("Subtitle", value=config.subtitle)

        config.palette_name = st.selectbox(
            "Color Palette",
            options=list(PALETTES.keys()),
            index=list(PALETTES.keys()).index(config.palette_name)
                if config.palette_name in PALETTES else 0,
            help="Choose a color scheme for your visualization",
        )

        # Palette preview
        palette = PALETTES[config.palette_name]
        cols = st.columns(min(5, len(palette)))
        for i, col in enumerate(cols):
            if i < len(palette):
                col.color_picker(
                    f"C{i+1}", value=palette[i], key=f"cp_{i}",
                    disabled=True, label_visibility="collapsed",
                )

        config.show_today_line = st.toggle("Show today line", value=config.show_today_line)
        config.show_status = st.toggle("Show status indicators", value=config.show_status)
        config.show_legend = st.toggle("Show legend", value=config.show_legend)

        st.divider()
        st.markdown("### 📅 Date Range")
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            if config.start_date:
                config.start_date = st.date_input("Start", value=config.start_date)
        with d_col2:
            if config.end_date:
                config.end_date = st.date_input("End", value=config.end_date)

        st.session_state["config"] = config


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT AREA
# ══════════════════════════════════════════════════════════════════════════════

def show_homepage():
    """Beautiful landing page with clear instructions and troubleshooting."""

    # Feature cards
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #EFF6FF;">📋</div>
            <h3>Simple Input</h3>
            <p>Just <strong>3 columns</strong> needed: Title, Start Date, and End Date.
            Add Category to group by team. The app figures out the rest.</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #F0FDF4;">🎨</div>
            <h3>Two Powerful Views</h3>
            <p><strong>Gantt Chart</strong> — swim lanes showing each team's work over time.<br>
            <strong>Block Diagram</strong> — everything packed into one slide showing parallel workload.</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #FEF3C7;">📤</div>
            <h3>Export Anywhere</h3>
            <p>Download as <strong>PowerPoint</strong> (editable shapes), <strong>PDF</strong> (vector),
            or <strong>PNG</strong> (high-res image). Ready for presentations.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # ── Step-by-step guide ────────────────────────────────────────────────
    st.markdown("### How It Works")
    st.markdown("""
    <div class="card">
        <div class="step-row">
            <span class="step-num">1</span>
            <div>
                <strong>Download the template</strong> — Click "Download Excel Template" in the sidebar.
                It comes pre-filled with sample data so you can see the format.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">2</span>
            <div>
                <strong>Add your data</strong> — Open the template in Excel or Google Sheets.
                Replace the sample rows with your own projects. Only Title, Start Date, and End Date are required.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">3</span>
            <div>
                <strong>Upload your file</strong> — Drag your .xlsx file into the uploader in the sidebar,
                or click to browse. The app instantly reads and validates your data.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">4</span>
            <div>
                <strong>Customize & export</strong> — Choose your color palette, toggle options,
                then download your chart as PowerPoint, PDF, or PNG.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── Excel format reference ────────────────────────────────────────────
    st.markdown("### Excel Format Reference")
    st.markdown("""
    <div class="card">
        <p style="margin-bottom: 0.75rem; color: #475569;">
            Your Excel file needs one sheet with the following columns. The app is flexible with
            column names — for example, "Name" works instead of "Title", or "Team" instead of "Category".
        </p>
        <table class="column-table">
            <thead>
                <tr>
                    <th>Column</th>
                    <th>Required?</th>
                    <th>Description</th>
                    <th>Example</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Title</strong></td>
                    <td><span class="required">Required</span></td>
                    <td>Name of the work item</td>
                    <td>Website Redesign</td>
                </tr>
                <tr>
                    <td><strong>Start Date</strong></td>
                    <td><span class="required">Required</span></td>
                    <td>When work begins (any date format works)</td>
                    <td>2025-01-15</td>
                </tr>
                <tr>
                    <td><strong>End Date</strong></td>
                    <td><span class="required">Required</span></td>
                    <td>When work ends</td>
                    <td>2025-03-31</td>
                </tr>
                <tr>
                    <td><strong>Category</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Team or workstream (used for colors &amp; grouping)</td>
                    <td>Marketing</td>
                </tr>
                <tr>
                    <td><strong>Description</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Brief description shown inside blocks</td>
                    <td>Complete site overhaul</td>
                </tr>
                <tr>
                    <td><strong>Status</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>planned, in_progress, done, or at_risk</td>
                    <td>in_progress</td>
                </tr>
                <tr>
                    <td><strong>Owner</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Person responsible</td>
                    <td>Sarah</td>
                </tr>
                <tr>
                    <td><strong>Label</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Short label shown on the block (e.g., "D1", "MVP")</td>
                    <td>M2</td>
                </tr>
            </tbody>
        </table>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── Try sample data ──────────────────────────────────────────────────
    st.markdown("### Quick Start")
    st.markdown("Don't have a file ready? Try the app with built-in sample data:")
    if st.button("🚀 Load Sample Data", type="primary", use_container_width=False):
        try:
            sample_bytes = create_template_bytes()
            loaded_items, loaded_config, loaded_warnings = read_excel(io.BytesIO(sample_bytes))
            loaded_config.title = "Product Launch 2025"
            loaded_config.subtitle = "Cross-functional roadmap"
            loaded_config.palette_name = "Vibrant"
            st.session_state["items"] = loaded_items
            st.session_state["config"] = loaded_config
            st.session_state["warnings"] = loaded_warnings
            st.session_state["load_error"] = None
            st.rerun()
        except Exception as e:
            st.error(f"Error loading sample: {e}")

    st.markdown("")

    # ── Troubleshooting ──────────────────────────────────────────────────
    st.markdown("### Troubleshooting")
    st.markdown("""
    <div class="help-box">
        <h4>📁 File won't upload?</h4>
        <ul>
            <li>Make sure your file is <strong>.xlsx</strong> format (not .csv, .xls, or .numbers)</li>
            <li>The file must be under 200 MB</li>
            <li>If using Google Sheets, download as .xlsx first (File → Download → Microsoft Excel)</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>⚠️ "Missing required column" error?</h4>
        <ul>
            <li>Your file needs at least these columns: <strong>Title</strong>, <strong>Start Date</strong>, <strong>End Date</strong></li>
            <li>Column names are flexible — "Name", "Task", or "Item" all work for the Title column</li>
            <li>Make sure the column headers are in <strong>Row 1</strong> of your spreadsheet</li>
            <li>Check there are no merged cells in the header row</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>📅 Date parsing errors?</h4>
        <ul>
            <li>Supported formats: <code>2025-01-15</code>, <code>01/15/2025</code>, <code>1/15/25</code></li>
            <li>Make sure date cells in Excel are formatted as <strong>Date</strong>, not Text</li>
            <li>Avoid leaving date cells blank — remove entire rows you don't need</li>
            <li>If Start Date is after End Date, the app will swap them automatically</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>🎨 Chart looks wrong?</h4>
        <ul>
            <li>Use the <strong>Date Range</strong> controls in the sidebar to adjust the visible timeframe</li>
            <li>Try a different <strong>Color Palette</strong> if colors are hard to distinguish</li>
            <li>For the Block Diagram, try <strong>4:3 aspect ratio</strong> if blocks feel too short</li>
            <li>Make sure your Category names are consistent (e.g., "Marketing" vs "marketing" are treated as different)</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)


def show_visualizations():
    """Main visualization view with Gantt, Block, and Data tabs."""
    items = st.session_state["items"]
    config = st.session_state["config"]

    # Quick stats
    all_categories = list(set(it.category for it in items))
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Work Items", len(items))
    col2.metric("Categories", len(all_categories))
    min_date = min(it.start_date for it in items)
    max_date = max(it.end_date for it in items)
    span_months = max(1, (max_date - min_date).days // 30)
    col3.metric("Timeline", f"{span_months} months")
    in_progress_count = sum(1 for it in items if it.status == "in_progress")
    col4.metric("In Progress", in_progress_count)

    # Clear data button
    st.markdown("")
    with st.columns([6, 1])[1]:
        if st.button("🗑️ Clear Data", help="Remove loaded data and start over"):
            st.session_state["items"] = None
            st.session_state["config"] = ChartConfig()
            st.session_state["warnings"] = []
            st.session_state["load_error"] = None
            st.rerun()

    # Tabs
    tab_gantt, tab_block, tab_data = st.tabs([
        "📊 Gantt Chart",
        "🧱 Block Diagram",
        "📋 Data Preview",
    ])

    # ── GANTT CHART TAB ──────────────────────────────────────────────────
    with tab_gantt:
        st.markdown("#### Swim Lane Gantt Chart")
        st.caption("Each category gets its own swim lane. Tasks are arranged to avoid overlaps within a lane.")

        try:
            with st.spinner("Rendering Gantt chart..."):
                gantt_preview = render_gantt(items, config, dpi=150)
            st.image(gantt_preview, use_container_width=True)

            st.markdown("---")
            st.markdown("**Export Gantt Chart**")
            g1, g2, g3 = st.columns(3)

            with g1:
                hires_gantt = render_gantt(items, config, dpi=300)
                st.download_button(
                    "📷 Download PNG (300 DPI)", data=hires_gantt,
                    file_name="roadmap_gantt.png", mime="image/png",
                    use_container_width=True,
                )
            with g2:
                gantt_pdf = render_gantt_pdf(items, config, dpi=300)
                st.download_button(
                    "📄 Download PDF", data=gantt_pdf,
                    file_name="roadmap_gantt.pdf", mime="application/pdf",
                    use_container_width=True,
                )
            with g3:
                gantt_pptx = export_gantt_pptx(items, config)
                st.download_button(
                    "📊 Download PowerPoint", data=gantt_pptx,
                    file_name="roadmap_gantt.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Error rendering Gantt chart: {e}")
            with st.expander("Show full error"):
                st.code(traceback.format_exc())

    # ── BLOCK DIAGRAM TAB ────────────────────────────────────────────────
    with tab_block:
        st.markdown("#### Space-Filling Block Diagram")
        st.caption("All work items packed into a single slide showing parallel workload. Great for executive presentations.")

        aspect_ratio = st.selectbox(
            "Slide Aspect Ratio",
            options=["16:9", "4:3"],
            index=0,
            help="16:9 for widescreen monitors and modern presentations. 4:3 for older projectors.",
        )

        try:
            with st.spinner("Rendering block diagram..."):
                block_preview = render_block_diagram(items, config, dpi=150, slide_aspect=aspect_ratio)
            st.image(block_preview, use_container_width=True)

            st.markdown("---")
            st.markdown("**Export Block Diagram**")
            b1, b2, b3 = st.columns(3)

            with b1:
                hires_block = render_block_diagram(items, config, dpi=300, slide_aspect=aspect_ratio)
                st.download_button(
                    "📷 Download PNG (300 DPI)", data=hires_block,
                    file_name="roadmap_block.png", mime="image/png",
                    use_container_width=True,
                )
            with b2:
                block_pdf_data = render_block_pdf(items, config, dpi=300, slide_aspect=aspect_ratio)
                st.download_button(
                    "📄 Download PDF", data=block_pdf_data,
                    file_name="roadmap_block.pdf", mime="application/pdf",
                    use_container_width=True,
                )
            with b3:
                block_pptx = export_block_pptx(items, config)
                st.download_button(
                    "📊 Download PowerPoint", data=block_pptx,
                    file_name="roadmap_block.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Error rendering block diagram: {e}")
            with st.expander("Show full error"):
                st.code(traceback.format_exc())

    # ── DATA PREVIEW TAB ─────────────────────────────────────────────────
    with tab_data:
        st.markdown("#### Your Data")

        df = pd.DataFrame([
            {
                "Title": it.title,
                "Start": it.start_date.strftime("%Y-%m-%d"),
                "End": it.end_date.strftime("%Y-%m-%d"),
                "Category": it.category,
                "Status": STATUS_LABELS.get(it.status, it.status),
                "Owner": it.owner,
                "Label": it.label,
                "Days": it.duration_days,
            }
            for it in items
        ])

        st.dataframe(
            df,
            use_container_width=True,
            height=min(600, 40 + len(df) * 35),
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    options=list(STATUS_LABELS.values()),
                    required=True,
                ),
                "Days": st.column_config.NumberColumn(format="%d days"),
                "Start": st.column_config.DateColumn(),
                "End": st.column_config.DateColumn(),
            },
        )

        st.markdown("#### Summary by Category")
        summary = df.groupby("Category").agg(
            Items=("Title", "count"),
            Earliest=("Start", "min"),
            Latest=("End", "max"),
        ).reset_index()
        st.dataframe(summary, use_container_width=True, hide_index=True)

        if st.session_state["warnings"]:
            st.markdown("#### Warnings")
            for w in st.session_state["warnings"]:
                st.warning(w)


# ── Route to homepage or visualization ───────────────────────────────────────
if st.session_state["items"] is not None and len(st.session_state["items"]) > 0:
    show_visualizations()
else:
    show_homepage()
