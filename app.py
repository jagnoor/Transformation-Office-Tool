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
APP_VERSION = "1.0"


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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Space+Grotesk:wght@500;600;700&display=swap');

    .stApp {
        background: linear-gradient(160deg, #FFF7ED 0%, #FDF2F8 25%, #EFF6FF 50%, #F0FDF4 75%, #FFFBEB 100%);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* ── Animated background blobs ──────────────────────────────── */
    @keyframes blob-float {
        0%, 100% { transform: translate(0, 0) scale(1); }
        33% { transform: translate(30px, -20px) scale(1.05); }
        66% { transform: translate(-15px, 15px) scale(0.95); }
    }

    /* ── Header ─────────────────────────────────────────────────── */
    .main-header {
        background: linear-gradient(135deg, #7C3AED 0%, #2563EB 30%, #0891B2 60%, #059669 100%);
        color: white;
        padding: 2.5rem 2.5rem 2.25rem;
        border-radius: 20px;
        margin-bottom: 0.75rem;
        box-shadow: 0 8px 32px rgba(124, 58, 237, 0.25), 0 2px 8px rgba(37, 99, 235, 0.15);
        position: relative;
        overflow: hidden;
    }
    .main-header::before {
        content: '';
        position: absolute; top: -50%; right: -20%;
        width: 400px; height: 400px;
        background: radial-gradient(circle, rgba(255,255,255,0.12) 0%, transparent 70%);
        border-radius: 50%;
        animation: blob-float 8s ease-in-out infinite;
        pointer-events: none;
    }
    .main-header::after {
        content: '';
        position: absolute; bottom: -30%; left: 10%;
        width: 300px; height: 300px;
        background: radial-gradient(circle, rgba(251,191,36,0.15) 0%, transparent 70%);
        border-radius: 50%;
        animation: blob-float 10s ease-in-out infinite reverse;
        pointer-events: none;
    }
    .main-header .header-label {
        font-size: 0.7rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.15em; color: #FDE68A; margin-bottom: 0.4rem;
        font-family: 'Space Grotesk', sans-serif;
    }
    .main-header h1 {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2rem; font-weight: 700; margin: 0; letter-spacing: -0.025em;
        line-height: 1.15;
        text-shadow: 0 2px 12px rgba(0,0,0,0.15);
    }
    .main-header .subtitle {
        color: rgba(255,255,255,0.85); margin: 0.4rem 0 0 0; font-size: 0.95rem;
        font-weight: 400;
    }
    .main-header .version-badge {
        position: absolute; top: 1.25rem; right: 1.5rem;
        background: rgba(255,255,255,0.2); color: white;
        font-size: 0.65rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.08em; padding: 0.25rem 0.75rem;
        border-radius: 20px; border: 1px solid rgba(255,255,255,0.3);
        backdrop-filter: blur(8px);
    }

    /* ── Cards ───────────────────────────────────────────────────── */
    .card {
        background: rgba(255,255,255,0.85); backdrop-filter: blur(12px);
        border-radius: 16px; padding: 1.6rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.04), 0 0 0 1px rgba(0,0,0,0.03);
        border: 1px solid rgba(255,255,255,0.8);
        margin-bottom: 1rem; height: 100%;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .card:hover {
        box-shadow: 0 12px 28px rgba(124,58,237,0.12), 0 0 0 1px rgba(124,58,237,0.08);
        transform: translateY(-4px);
    }
    .card h3 {
        margin-top: 0; font-weight: 700; font-size: 1.08rem;
        font-family: 'Space Grotesk', sans-serif;
        background: linear-gradient(135deg, #7C3AED, #2563EB);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .card p { color: #475569; font-size: 0.9rem; line-height: 1.6; }

    .feature-icon {
        width: 52px; height: 52px; border-radius: 14px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.6rem; margin-bottom: 0.85rem;
        box-shadow: 0 3px 10px rgba(0,0,0,0.06);
    }

    /* ── Tabs ────────────────────────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px; padding: 8px 22px; font-weight: 600;
    }

    /* ── Buttons ─────────────────────────────────────────────────── */
    .stDownloadButton > button {
        border-radius: 10px; font-weight: 600; padding: 0.5rem 1.5rem;
        transition: all 0.2s ease;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(124,58,237,0.2);
    }

    /* ── Sidebar ─────────────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FEFCE8 0%, #FFF7ED 30%, #FDF2F8 70%, #EFF6FF 100%);
        border-right: 1px solid rgba(0,0,0,0.06);
    }
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #6D28D9; font-size: 0.78rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.06em;
    }

    .sidebar-brand {
        padding: 0.75rem 0 1.25rem 0; margin-bottom: 0.5rem;
        border-bottom: 2px solid transparent;
        border-image: linear-gradient(90deg, #7C3AED, #2563EB, #0891B2, #059669) 1;
        text-align: center;
    }
    .sidebar-brand .brand-icon { font-size: 1.5rem; margin-bottom: 0.25rem; }
    .sidebar-brand .brand-name {
        font-size: 0.8rem; font-weight: 700;
        font-family: 'Space Grotesk', sans-serif;
        background: linear-gradient(135deg, #7C3AED, #2563EB);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
        text-transform: uppercase; letter-spacing: 0.08em;
    }
    .sidebar-brand .brand-sub {
        font-size: 0.65rem; color: #64748B; margin-top: 0.15rem;
    }

    /* ── Metrics ─────────────────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: rgba(255,255,255,0.8); backdrop-filter: blur(8px);
        border-radius: 14px; padding: 1rem;
        border: 1px solid rgba(255,255,255,0.6);
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.72rem !important; font-weight: 700 !important;
        text-transform: uppercase; letter-spacing: 0.05em;
        color: #7C3AED !important;
    }
    [data-testid="stMetricValue"] {
        font-weight: 800 !important;
        background: linear-gradient(135deg, #7C3AED, #2563EB) !important;
        -webkit-background-clip: text !important; -webkit-text-fill-color: transparent !important;
        background-clip: text !important;
    }

    /* ── Hide Streamlit chrome ───────────────────────────────────── */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    /* ── Steps ───────────────────────────────────────────────────── */
    @keyframes step-pulse {
        0%, 100% { box-shadow: 0 0 0 0 rgba(124,58,237,0.3); }
        50% { box-shadow: 0 0 0 6px rgba(124,58,237,0); }
    }
    .step-num {
        display: inline-flex; align-items: center; justify-content: center;
        background: linear-gradient(135deg, #7C3AED, #2563EB); color: white;
        border-radius: 50%;
        width: 36px; height: 36px; font-weight: 700; font-size: 0.95rem;
        margin-right: 14px; flex-shrink: 0;
        font-family: 'Space Grotesk', sans-serif;
    }
    .step-row:hover .step-num {
        animation: step-pulse 1.5s ease-in-out infinite;
    }
    .step-row {
        display: flex; align-items: center; padding: 0.75rem 0.5rem;
        border-bottom: 1px solid rgba(124,58,237,0.06); font-size: 0.92rem;
        color: #334155; line-height: 1.55;
        border-radius: 8px; margin: 2px 0;
        transition: background 0.2s ease;
    }
    .step-row:hover { background: rgba(124,58,237,0.03); }
    .step-row:last-child { border-bottom: none; }
    .step-row strong { color: #4C1D95; }

    /* ── Help boxes ──────────────────────────────────────────────── */
    .help-box {
        background: linear-gradient(135deg, rgba(255,247,237,0.9), rgba(254,243,199,0.7));
        backdrop-filter: blur(8px);
        border: 1px solid #FED7AA; border-radius: 14px;
        padding: 1.25rem 1.35rem; margin-top: 1rem;
        transition: all 0.2s ease;
    }
    .help-box:hover {
        box-shadow: 0 4px 16px rgba(217,119,6,0.1);
        transform: translateY(-1px);
    }
    .help-box h4 { margin: 0 0 0.5rem 0; color: #9A3412; font-size: 0.95rem;
        font-family: 'Space Grotesk', sans-serif;
    }
    .help-box p, .help-box li { color: #78350F; font-size: 0.85rem; line-height: 1.55; }

    /* ── Column reference table ──────────────────────────────────── */
    .column-table {
        width: 100%; border-collapse: collapse; margin: 0.75rem 0;
    }
    .column-table th {
        background: linear-gradient(135deg, #EDE9FE, #DBEAFE); padding: 0.6rem 0.85rem;
        text-align: left;
        font-size: 0.78rem; color: #4C1D95; border-bottom: 2px solid #C4B5FD;
        font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em;
        font-family: 'Space Grotesk', sans-serif;
    }
    .column-table td {
        padding: 0.55rem 0.85rem; border-bottom: 1px solid #F1F5F9;
        font-size: 0.85rem; color: #334155;
    }
    .column-table tr:hover { background: rgba(124,58,237,0.03); }
    .required {
        color: white; font-weight: 700; font-size: 0.72rem;
        background: linear-gradient(135deg, #DC2626, #E11D48);
        padding: 0.15rem 0.55rem; border-radius: 10px;
        text-transform: uppercase; letter-spacing: 0.03em;
    }
    .optional {
        color: white; font-weight: 700; font-size: 0.72rem;
        background: linear-gradient(135deg, #059669, #0D9488);
        padding: 0.15rem 0.55rem; border-radius: 10px;
        text-transform: uppercase; letter-spacing: 0.03em;
    }

    /* ── Section headers ─────────────────────────────────────────── */
    .section-header {
        font-size: 1.2rem; font-weight: 700;
        font-family: 'Space Grotesk', sans-serif;
        background: linear-gradient(135deg, #7C3AED, #2563EB);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 1.75rem 0 0.85rem 0; padding-bottom: 0.5rem;
        border-bottom: 2px solid transparent;
        border-image: linear-gradient(90deg, #7C3AED, #2563EB, transparent) 1;
    }

    /* ── Footer ──────────────────────────────────────────────────── */
    .app-footer {
        text-align: center; padding: 1.75rem 0 0.75rem;
        border-top: 2px solid transparent;
        border-image: linear-gradient(90deg, transparent, #7C3AED, #2563EB, #0891B2, transparent) 1;
        margin-top: 2.5rem;
    }
    .app-footer p {
        font-size: 0.75rem; color: #94A3B8; margin: 0.15rem 0;
    }
    .app-footer .footer-label {
        font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.08em; font-size: 0.65rem;
        background: linear-gradient(135deg, #7C3AED, #2563EB);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    /* ── Export section ───────────────────────────────────────────── */
    .export-header {
        font-size: 0.85rem; font-weight: 700; color: #6D28D9;
        text-transform: uppercase; letter-spacing: 0.05em;
        margin-bottom: 0.5rem; padding-top: 0.5rem;
        font-family: 'Space Grotesk', sans-serif;
    }

    /* ── Welcome hero ────────────────────────────────────────────── */
    .welcome-hero {
        text-align: center; padding: 1.5rem 1rem 0.5rem;
    }
    .welcome-hero .hero-emoji {
        font-size: 3rem; margin-bottom: 0.5rem;
        animation: blob-float 4s ease-in-out infinite;
        display: inline-block;
    }
    .welcome-hero h2 {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.5rem; font-weight: 700; margin: 0;
        background: linear-gradient(135deg, #7C3AED 0%, #2563EB 50%, #0891B2 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .welcome-hero p {
        color: #64748B; font-size: 0.92rem; margin: 0.35rem 0 0 0;
    }

    /* ── Stat pills on viz page ──────────────────────────────────── */
    .stat-ribbon {
        display: flex; gap: 0.75rem; flex-wrap: wrap;
        margin-bottom: 0.5rem;
    }
    .stat-pill {
        display: inline-flex; align-items: center; gap: 0.4rem;
        padding: 0.4rem 0.9rem; border-radius: 20px;
        font-size: 0.82rem; font-weight: 600;
        backdrop-filter: blur(8px);
    }
    .stat-pill .pill-icon { font-size: 1rem; }
    .stat-pill .pill-val { font-family: 'Space Grotesk', sans-serif; font-weight: 700; }

    /* ── Quick-start card ────────────────────────────────────────── */
    .qs-card {
        background: linear-gradient(135deg, #EDE9FE 0%, #DBEAFE 50%, #D1FAE5 100%);
        border-radius: 16px; padding: 1.5rem 1.75rem;
        border: 1px solid rgba(124,58,237,0.15);
        margin-bottom: 1rem;
        box-shadow: 0 4px 16px rgba(124,58,237,0.08);
    }
    .qs-card h3 {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 1.1rem; font-weight: 700; color: #4C1D95; margin: 0 0 0.35rem 0;
    }
    .qs-card p { color: #475569; font-size: 0.88rem; margin: 0; line-height: 1.5; }
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
    <div class="version-badge">v{APP_VERSION}</div>
    <div class="header-label">Transformation Office</div>
    <h1>Block & Gantt Creator Tool</h1>
    <p class="subtitle">Turn boring spreadsheets into stunning project visuals — in seconds, not hours</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-icon">&#x1F3A8;</div>
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
        <div style="font-size: 0.6rem; text-transform: uppercase;
             letter-spacing: 0.08em; font-weight: 700;
             background: linear-gradient(135deg, #7C3AED, #2563EB);
             -webkit-background-clip: text; -webkit-text-fill-color: transparent;
             background-clip: text;">Version</div>
        <div style="font-size: 0.7rem; color: #64748B; font-weight: 600;">v{APP_VERSION}</div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT AREA
# ══════════════════════════════════════════════════════════════════════════════

def show_homepage():
    """Landing page with clear instructions and troubleshooting."""

    # Welcome hero
    st.markdown("""
    <div class="welcome-hero">
        <div class="hero-emoji">&#x2728;</div>
        <h2>Welcome! Let's build something beautiful.</h2>
        <p>Upload a spreadsheet, pick your colors, and watch the magic happen.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # Feature cards
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: linear-gradient(135deg, #EDE9FE, #DBEAFE);">
                <span style="font-size: 1.5rem;">&#x1F4CB;</span>
            </div>
            <h3>Dead-Simple Input</h3>
            <p>Just <strong>3 columns</strong> — Title, Start Date, End Date — and you're off to the races.
            Add Category for color-coded workstreams. We handle the heavy lifting.</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: linear-gradient(135deg, #D1FAE5, #A7F3D0);">
                <span style="font-size: 1.5rem;">&#x1F3A8;</span>
            </div>
            <h3>Two Ways to Visualize</h3>
            <p><strong>Gantt Chart</strong> — swim lanes that show every workstream over time.<br>
            <strong>Block Diagram</strong> — a packed, bird's-eye view of your entire portfolio on one slide.</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="card">
            <div class="feature-icon" style="background: linear-gradient(135deg, #FEF3C7, #FDE68A);">
                <span style="font-size: 1.5rem;">&#x1F680;</span>
            </div>
            <h3>Export & Impress</h3>
            <p>Download as <strong>PowerPoint</strong> (fully editable shapes), <strong>PDF</strong> (crisp vectors),
            or <strong>PNG</strong> (high-res). Ready to drop into your next leadership deck.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # ── Step-by-step guide ────────────────────────────────────────────────
    st.markdown('<div class="section-header">Four Steps to Stunning Charts</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
        <div class="step-row">
            <span class="step-num">1</span>
            <div>
                <strong>Grab the template</strong> — Hit "Download Excel Template" in the sidebar.
                It's pre-loaded with sample data so you can see exactly what's expected.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">2</span>
            <div>
                <strong>Drop in your data</strong> — Open it in Excel or Google Sheets and swap the samples
                for your real initiatives. Three columns is all you need to get started.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">3</span>
            <div>
                <strong>Upload & go</strong> — Drag your .xlsx into the sidebar uploader. The tool reads,
                validates, and previews your data instantly. No waiting around.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">4</span>
            <div>
                <strong>Make it yours</strong> — Pick a palette, tweak the date range, then download your
                chart as PowerPoint, PDF, or PNG. Done!
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── Excel format reference ────────────────────────────────────────────
    st.markdown('<div class="section-header">What Goes in Your Spreadsheet</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
        <p style="margin-bottom: 0.75rem; color: #475569; font-size: 0.88rem;">
            One sheet, a few columns, endless possibilities. Column names are flexible &mdash;
            "Name" works in place of "Title", "Team" instead of "Category" &mdash; we'll figure it out.
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
    st.markdown('<div class="section-header">No Spreadsheet? No Problem!</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="qs-card">
        <h3>&#x26A1; Try it right now</h3>
        <p>Don't have a file handy? Load our built-in sample data with 25 initiatives across 6 workstreams and see the tool in action.</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Load Sample Data", type="primary", use_container_width=False):
        try:
            sample_bytes = create_template_bytes()
            loaded_items, loaded_config, loaded_warnings = read_excel(io.BytesIO(sample_bytes))
            loaded_config.title = "Transformation Roadmap 2025"
            loaded_config.subtitle = "25 initiatives across 6 workstreams"
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
    st.markdown('<div class="section-header">Something Not Working? We\'ve Got You.</div>', unsafe_allow_html=True)
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
        <p>Block & Gantt Creator Tool &middot; v{APP_VERSION}</p>
        <p>Made with care for the people who make transformation happen.</p>
    </div>
    """, unsafe_allow_html=True)


def show_visualizations():
    """Main visualization view with Gantt, Block, and Data tabs."""
    items = st.session_state["items"]
    config = st.session_state["config"]

    # Quick stats — colorful pills
    all_categories = list(set(it.category for it in items))
    min_date = min(it.start_date for it in items)
    max_date = max(it.end_date for it in items)
    span_months = max(1, (max_date - min_date).days // 30)
    in_progress_count = sum(1 for it in items if it.status == "in_progress")
    done_count = sum(1 for it in items if it.status == "done")
    milestones = sum(1 for it in items if it.is_milestone)

    st.markdown(f"""
    <div class="stat-ribbon">
        <span class="stat-pill" style="background: linear-gradient(135deg, #EDE9FE, #DDD6FE); color: #6D28D9;">
            <span class="pill-icon">&#x1F4CA;</span>
            <span class="pill-val">{len(items)}</span> items
        </span>
        <span class="stat-pill" style="background: linear-gradient(135deg, #DBEAFE, #BFDBFE); color: #1D4ED8;">
            <span class="pill-icon">&#x1F3AF;</span>
            <span class="pill-val">{len(all_categories)}</span> categories
        </span>
        <span class="stat-pill" style="background: linear-gradient(135deg, #D1FAE5, #A7F3D0); color: #047857;">
            <span class="pill-icon">&#x1F552;</span>
            <span class="pill-val">{span_months}</span> months
        </span>
        <span class="stat-pill" style="background: linear-gradient(135deg, #FEF3C7, #FDE68A); color: #B45309;">
            <span class="pill-icon">&#x1F525;</span>
            <span class="pill-val">{in_progress_count}</span> in progress
        </span>
        <span class="stat-pill" style="background: linear-gradient(135deg, #D1FAE5, #BBF7D0); color: #15803D;">
            <span class="pill-icon">&#x2705;</span>
            <span class="pill-val">{done_count}</span> done
        </span>
        {"<span class='stat-pill' style='background: linear-gradient(135deg, #FCE7F3, #FBCFE8); color: #BE185D;'><span class='pill-icon'>&#x1F48E;</span><span class='pill-val'>" + str(milestones) + "</span> milestones</span>" if milestones > 0 else ""}
    </div>
    """, unsafe_allow_html=True)

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
        "Gantt Chart  \u2003",
        "Block Diagram  \u2003",
        "Data Preview  \u2003",
    ])

    # ── GANTT CHART TAB ──────────────────────────────────────────────────
    with tab_gantt:
        st.markdown("#### Swim Lane Gantt Chart")
        st.caption("Every category gets its own lane. Tasks stack smartly to avoid collisions. Hover over lanes to follow the flow.")

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
        st.caption("Your entire portfolio packed into one gorgeous slide. Perfect for the exec who wants the big picture at a glance.")

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
        <p>Block & Gantt Creator Tool &middot; v{APP_VERSION}</p>
        <p>Made with care for the people who make transformation happen.</p>
    </div>
    """, unsafe_allow_html=True)


# ── Route to homepage or visualization ───────────────────────────────────────
if st.session_state["items"] is not None and len(st.session_state["items"]) > 0:
    show_visualizations()
else:
    show_homepage()
