"""
Transformation Office — Block & Gantt Creator Tool.

Creates presentation-ready Gantt charts and space-filling block diagrams
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


# ── Constants ────────────────────────────────────────────────────────────────
APP_NAME = "Block & Gantt Creator"
APP_FULL_NAME = "Transformation Office — Block & Gantt Creator Tool"
APP_VERSION = "0.9 Beta"


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=APP_FULL_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Global ─────────────────────────────────────────────────── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    .stApp {
        background-color: #F8FAFC;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* ── Header ─────────────────────────────────────────────────── */
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 40%, #1E3A5F 100%);
        color: white;
        padding: 2.25rem 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 0.5rem;
        box-shadow: 0 4px 24px rgba(15, 23, 42, 0.18);
        position: relative;
        overflow: hidden;
    }
    .main-header::before {
        content: '';
        position: absolute; top: 0; right: 0;
        width: 300px; height: 100%;
        background: radial-gradient(circle at 80% 50%, rgba(59,130,246,0.12) 0%, transparent 70%);
        pointer-events: none;
    }
    .main-header .header-label {
        font-size: 0.7rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.12em; color: #60A5FA; margin-bottom: 0.35rem;
    }
    .main-header h1 {
        font-size: 1.85rem; font-weight: 800; margin: 0; letter-spacing: -0.025em;
        line-height: 1.2;
    }
    .main-header .subtitle {
        color: #94A3B8; margin: 0.35rem 0 0 0; font-size: 0.95rem;
        font-weight: 400;
    }
    .main-header .version-badge {
        position: absolute; top: 1.25rem; right: 1.5rem;
        background: rgba(251, 191, 36, 0.15); color: #FBBF24;
        font-size: 0.65rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.08em; padding: 0.25rem 0.65rem;
        border-radius: 20px; border: 1px solid rgba(251, 191, 36, 0.3);
    }

    /* ── Beta banner ────────────────────────────────────────────── */
    .beta-banner {
        background: linear-gradient(90deg, #FFFBEB 0%, #FEF3C7 100%);
        border: 1px solid #F59E0B;
        border-left: 4px solid #F59E0B;
        border-radius: 8px;
        padding: 0.75rem 1.25rem;
        margin-bottom: 1.25rem;
        display: flex; align-items: center; gap: 0.75rem;
    }
    .beta-banner .beta-icon {
        font-size: 1.1rem; flex-shrink: 0;
    }
    .beta-banner .beta-text {
        font-size: 0.82rem; color: #92400E; line-height: 1.45;
    }
    .beta-banner .beta-text strong { color: #78350F; }

    /* ── Cards ───────────────────────────────────────────────────── */
    .card {
        background: white; border-radius: 12px; padding: 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05); border: 1px solid #E2E8F0;
        margin-bottom: 1rem; height: 100%;
        transition: box-shadow 0.2s ease, transform 0.2s ease;
    }
    .card:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transform: translateY(-1px);
    }
    .card h3 { margin-top: 0; color: #0F172A; font-weight: 700; font-size: 1.05rem; }
    .card p { color: #64748B; font-size: 0.9rem; line-height: 1.55; }

    .feature-icon {
        width: 48px; height: 48px; border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.5rem; margin-bottom: 0.75rem;
    }

    /* ── Tabs ────────────────────────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 8px 20px; font-weight: 600;
    }

    /* ── Buttons ─────────────────────────────────────────────────── */
    .stDownloadButton > button {
        border-radius: 8px; font-weight: 600; padding: 0.5rem 1.5rem;
    }

    /* ── Sidebar ─────────────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF; border-right: 1px solid #E2E8F0;
    }
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #0F172A; font-size: 0.8rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.06em;
    }

    .sidebar-brand {
        padding: 0.5rem 0 1rem 0; margin-bottom: 0.5rem;
        border-bottom: 1px solid #F1F5F9; text-align: center;
    }
    .sidebar-brand .brand-name {
        font-size: 0.75rem; font-weight: 700; color: #0F172A;
        text-transform: uppercase; letter-spacing: 0.08em;
    }
    .sidebar-brand .brand-sub {
        font-size: 0.65rem; color: #94A3B8; margin-top: 0.15rem;
    }

    /* ── Metrics ─────────────────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: white; border-radius: 10px; padding: 1rem;
        border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.75rem !important; font-weight: 600 !important;
        text-transform: uppercase; letter-spacing: 0.04em;
        color: #64748B !important;
    }
    [data-testid="stMetricValue"] {
        font-weight: 800 !important; color: #0F172A !important;
    }

    /* ── Hide Streamlit chrome ───────────────────────────────────── */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    /* ── Steps ───────────────────────────────────────────────────── */
    .step-num {
        display: inline-flex; align-items: center; justify-content: center;
        background: #2563EB; color: white; border-radius: 50%;
        width: 32px; height: 32px; font-weight: 700; font-size: 0.9rem;
        margin-right: 12px; flex-shrink: 0;
    }
    .step-row {
        display: flex; align-items: center; padding: 0.65rem 0;
        border-bottom: 1px solid #F1F5F9; font-size: 0.92rem;
        color: #334155; line-height: 1.5;
    }
    .step-row:last-child { border-bottom: none; }
    .step-row strong { color: #0F172A; }

    /* ── Help boxes ──────────────────────────────────────────────── */
    .help-box {
        background: #FFF7ED; border: 1px solid #FED7AA; border-radius: 10px;
        padding: 1.25rem; margin-top: 1rem;
    }
    .help-box h4 { margin: 0 0 0.5rem 0; color: #9A3412; font-size: 0.95rem; }
    .help-box p, .help-box li { color: #78350F; font-size: 0.85rem; line-height: 1.55; }

    /* ── Column reference table ──────────────────────────────────── */
    .column-table {
        width: 100%; border-collapse: collapse; margin: 0.75rem 0;
    }
    .column-table th {
        background: #F1F5F9; padding: 0.5rem 0.75rem; text-align: left;
        font-size: 0.8rem; color: #475569; border-bottom: 2px solid #E2E8F0;
        font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em;
    }
    .column-table td {
        padding: 0.5rem 0.75rem; border-bottom: 1px solid #F1F5F9;
        font-size: 0.85rem; color: #334155;
    }
    .column-table tr:hover { background: #F8FAFC; }
    .required { color: #DC2626; font-weight: 600; }
    .optional { color: #059669; }

    /* ── Section headers ─────────────────────────────────────────── */
    .section-header {
        font-size: 1.15rem; font-weight: 700; color: #0F172A;
        margin: 1.5rem 0 0.75rem 0; padding-bottom: 0.4rem;
        border-bottom: 2px solid #E2E8F0;
    }

    /* ── Footer ──────────────────────────────────────────────────── */
    .app-footer {
        text-align: center; padding: 1.5rem 0 0.5rem;
        border-top: 1px solid #E2E8F0; margin-top: 2rem;
    }
    .app-footer p {
        font-size: 0.75rem; color: #94A3B8; margin: 0.15rem 0;
    }
    .app-footer .footer-label {
        font-weight: 600; color: #64748B; text-transform: uppercase;
        letter-spacing: 0.06em; font-size: 0.65rem;
    }

    /* ── Export section ───────────────────────────────────────────── */
    .export-header {
        font-size: 0.85rem; font-weight: 700; color: #475569;
        text-transform: uppercase; letter-spacing: 0.05em;
        margin-bottom: 0.5rem; padding-top: 0.5rem;
    }
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
st.markdown(f"""
<div class="main-header">
    <div class="version-badge">{APP_VERSION}</div>
    <div class="header-label">Transformation Office</div>
    <h1>Block & Gantt Creator Tool</h1>
    <p class="subtitle">Generate presentation-ready project visualizations from a simple spreadsheet</p>
</div>
""", unsafe_allow_html=True)

# ── Beta notice ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="beta-banner">
    <span class="beta-icon">&#9432;</span>
    <div class="beta-text">
        <strong>Early Access Preview</strong> &mdash; This tool is currently in active development.
        We are gathering user feedback and refining functionality. You may encounter occasional
        formatting inconsistencies or limitations. Your input is invaluable &mdash; please share
        any feedback or issues with the Transformation Office team.
    </div>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-name">Transformation Office</div>
        <div class="brand-sub">Block & Gantt Creator Tool</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Step 1 — Get Template")
    template_bytes = create_template_bytes()
    st.download_button(
        label="Download Excel Template",
        data=template_bytes,
        file_name="transformation_office_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Pre-filled template with sample data and formatting instructions",
    )

    st.markdown("### Step 2 — Upload Your File")
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
                    with st.expander(f"Warnings ({len(loaded_warnings)})", expanded=False):
                        for w in loaded_warnings:
                            st.caption(w)
        except Exception as e:
            st.session_state["load_error"] = str(e)
            st.session_state["items"] = None
            st.error("Could not read file — see troubleshooting on the home page")

    # Show load error details in sidebar
    if st.session_state["load_error"]:
        with st.expander("Error Details", expanded=True):
            st.code(st.session_state["load_error"], language=None)

    st.divider()

    # Settings — only when data is loaded
    if st.session_state["items"] is not None:
        st.markdown("### Chart Settings")

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
        st.markdown("### Date Range")
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            if config.start_date:
                config.start_date = st.date_input("Start", value=config.start_date)
        with d_col2:
            if config.end_date:
                config.end_date = st.date_input("End", value=config.end_date)

        st.session_state["config"] = config

    # Sidebar footer
    st.divider()
    st.markdown(f"""
    <div style="text-align: center; padding: 0.25rem 0;">
        <div style="font-size: 0.6rem; color: #94A3B8; text-transform: uppercase;
             letter-spacing: 0.08em; font-weight: 600;">Version</div>
        <div style="font-size: 0.7rem; color: #64748B; font-weight: 500;">{APP_VERSION}</div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT AREA
# ══════════════════════════════════════════════════════════════════════════════

def show_homepage():
    """Landing page with clear instructions and troubleshooting."""

    # Feature cards
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #EFF6FF;">
                <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="#2563EB" stroke-width="2">
                    <path d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7"/>
                </svg>
            </div>
            <h3>Simple Input</h3>
            <p>Just <strong>3 columns</strong> needed: Title, Start Date, and End Date.
            Add Category to group by workstream. The tool handles the rest.</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #F0FDF4;">
                <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="#059669" stroke-width="2">
                    <rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/>
                    <rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>
                </svg>
            </div>
            <h3>Two Visualization Modes</h3>
            <p><strong>Gantt Chart</strong> — swim lanes showing each workstream over time.<br>
            <strong>Block Diagram</strong> — all items packed into one slide showing parallel workload.</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: #FEF3C7;">
                <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="#D97706" stroke-width="2">
                    <path d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                </svg>
            </div>
            <h3>Presentation-Ready Export</h3>
            <p>Download as <strong>PowerPoint</strong> (fully editable shapes), <strong>PDF</strong> (vector quality),
            or <strong>PNG</strong> (high-res image). Ready for leadership decks.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # ── Step-by-step guide ────────────────────────────────────────────────
    st.markdown('<div class="section-header">How It Works</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
        <div class="step-row">
            <span class="step-num">1</span>
            <div>
                <strong>Download the template</strong> — Click "Download Excel Template" in the sidebar.
                It comes pre-filled with sample data so you can see the expected format.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">2</span>
            <div>
                <strong>Add your data</strong> — Open the template in Excel or Google Sheets.
                Replace the sample rows with your own initiatives. Only Title, Start Date, and End Date are required.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">3</span>
            <div>
                <strong>Upload your file</strong> — Drag your .xlsx file into the uploader in the sidebar,
                or click "Browse files." The tool instantly reads and validates your data.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">4</span>
            <div>
                <strong>Customize & export</strong> — Select your color palette, adjust the date range,
                then download your chart as PowerPoint, PDF, or PNG.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── Excel format reference ────────────────────────────────────────────
    st.markdown('<div class="section-header">Excel Format Reference</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
        <p style="margin-bottom: 0.75rem; color: #475569; font-size: 0.88rem;">
            Your Excel file needs one sheet with the following columns. Column names are flexible &mdash;
            for example, "Name" works in place of "Title", or "Team" in place of "Category".
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
                    <td>Name of the initiative or work item</td>
                    <td>Website Redesign</td>
                </tr>
                <tr>
                    <td><strong>Start Date</strong></td>
                    <td><span class="required">Required</span></td>
                    <td>When work begins (any standard date format)</td>
                    <td>2025-01-15</td>
                </tr>
                <tr>
                    <td><strong>End Date</strong></td>
                    <td><span class="required">Required</span></td>
                    <td>When work is expected to complete</td>
                    <td>2025-03-31</td>
                </tr>
                <tr>
                    <td><strong>Category</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Workstream, team, or department (used for color coding &amp; swim lanes)</td>
                    <td>Marketing</td>
                </tr>
                <tr>
                    <td><strong>Description</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Brief description displayed inside blocks</td>
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
                    <td>Responsible individual or lead</td>
                    <td>Sarah</td>
                </tr>
                <tr>
                    <td><strong>Label</strong></td>
                    <td><span class="optional">Optional</span></td>
                    <td>Short identifier shown on the block (e.g., "D1", "MVP")</td>
                    <td>M2</td>
                </tr>
            </tbody>
        </table>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── Try sample data ──────────────────────────────────────────────────
    st.markdown('<div class="section-header">Quick Start</div>', unsafe_allow_html=True)
    st.markdown("Don't have a file ready? Load the built-in sample data to explore the tool:")
    if st.button("Load Sample Data", type="primary", use_container_width=False):
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
    st.markdown('<div class="section-header">Troubleshooting</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="help-box">
        <h4>File won't upload?</h4>
        <ul>
            <li>Ensure your file is in <strong>.xlsx</strong> format (not .csv, .xls, or .numbers)</li>
            <li>The file must be under 200 MB</li>
            <li>If using Google Sheets, export as .xlsx first: <em>File &rarr; Download &rarr; Microsoft Excel</em></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>"Missing required column" error?</h4>
        <ul>
            <li>Your file needs at least these columns: <strong>Title</strong>, <strong>Start Date</strong>, <strong>End Date</strong></li>
            <li>Column names are flexible &mdash; "Name", "Task", or "Item" all map to the Title column</li>
            <li>Ensure column headers are in <strong>Row 1</strong> of your spreadsheet</li>
            <li>Verify there are no merged cells in the header row</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>Date parsing errors?</h4>
        <ul>
            <li>Supported formats: <code>2025-01-15</code>, <code>01/15/2025</code>, <code>1/15/25</code></li>
            <li>Ensure date cells are formatted as <strong>Date</strong> in Excel, not as plain text</li>
            <li>Remove blank rows rather than leaving empty cells between data</li>
            <li>If Start Date is after End Date, the tool will swap them automatically and flag a warning</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>Chart looks wrong?</h4>
        <ul>
            <li>Adjust the <strong>Date Range</strong> in the sidebar to focus on the relevant time period</li>
            <li>Try a different <strong>Color Palette</strong> if categories are difficult to distinguish</li>
            <li>For the Block Diagram, experiment with the <strong>4:3 aspect ratio</strong> if blocks appear too narrow</li>
            <li>Ensure Category names are consistent (e.g., "Marketing" and "marketing" will be treated as separate groups)</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # ── Footer ────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="app-footer">
        <p class="footer-label">Transformation Office</p>
        <p>Block & Gantt Creator Tool &middot; {APP_VERSION}</p>
        <p>For internal use. Please direct feedback to the Transformation Office team.</p>
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
        if st.button("Clear Data", help="Remove loaded data and return to the home screen"):
            st.session_state["items"] = None
            st.session_state["config"] = ChartConfig()
            st.session_state["warnings"] = []
            st.session_state["load_error"] = None
            st.rerun()

    # Tabs
    tab_gantt, tab_block, tab_data = st.tabs([
        "Gantt Chart",
        "Block Diagram",
        "Data Preview",
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
            st.markdown('<div class="export-header">Export Gantt Chart</div>', unsafe_allow_html=True)
            g1, g2, g3 = st.columns(3)

            with g1:
                hires_gantt = render_gantt(items, config, dpi=300)
                st.download_button(
                    "Download PNG (300 DPI)", data=hires_gantt,
                    file_name="gantt_chart.png", mime="image/png",
                    use_container_width=True,
                )
            with g2:
                gantt_pdf = render_gantt_pdf(items, config, dpi=300)
                st.download_button(
                    "Download PDF", data=gantt_pdf,
                    file_name="gantt_chart.pdf", mime="application/pdf",
                    use_container_width=True,
                )
            with g3:
                gantt_pptx = export_gantt_pptx(items, config)
                st.download_button(
                    "Download PowerPoint", data=gantt_pptx,
                    file_name="gantt_chart.pptx",
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
        st.caption("All work items packed into a single slide showing parallel workload. Ideal for executive presentations.")

        aspect_ratio = st.selectbox(
            "Slide Aspect Ratio",
            options=["16:9", "4:3"],
            index=0,
            help="16:9 for widescreen displays and modern presentations. 4:3 for older projectors.",
        )

        try:
            with st.spinner("Rendering block diagram..."):
                block_preview = render_block_diagram(items, config, dpi=150, slide_aspect=aspect_ratio)
            st.image(block_preview, use_container_width=True)

            st.markdown("---")
            st.markdown('<div class="export-header">Export Block Diagram</div>', unsafe_allow_html=True)
            b1, b2, b3 = st.columns(3)

            with b1:
                hires_block = render_block_diagram(items, config, dpi=300, slide_aspect=aspect_ratio)
                st.download_button(
                    "Download PNG (300 DPI)", data=hires_block,
                    file_name="block_diagram.png", mime="image/png",
                    use_container_width=True,
                )
            with b2:
                block_pdf_data = render_block_pdf(items, config, dpi=300, slide_aspect=aspect_ratio)
                st.download_button(
                    "Download PDF", data=block_pdf_data,
                    file_name="block_diagram.pdf", mime="application/pdf",
                    use_container_width=True,
                )
            with b3:
                block_pptx = export_block_pptx(items, config)
                st.download_button(
                    "Download PowerPoint", data=block_pptx,
                    file_name="block_diagram.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Error rendering block diagram: {e}")
            with st.expander("Show full error"):
                st.code(traceback.format_exc())

    # ── DATA PREVIEW TAB ─────────────────────────────────────────────────
    with tab_data:
        st.markdown("#### Loaded Data")

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

    # ── Page footer ──────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="app-footer">
        <p class="footer-label">Transformation Office</p>
        <p>Block & Gantt Creator Tool &middot; {APP_VERSION}</p>
    </div>
    """, unsafe_allow_html=True)


# ── Route to homepage or visualization ───────────────────────────────────────
if st.session_state["items"] is not None and len(st.session_state["items"]) > 0:
    show_visualizations()
else:
    show_homepage()
