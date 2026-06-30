"""
Transformation Office — Block & Gantt Creator Tool.

Creates presentation-ready Gantt charts and space-filling block diagrams
from simple Excel input. Exports to PowerPoint, PDF, and PNG.
"""
import io
import traceback
import zipfile

import pandas as pd
import streamlit as st
from datetime import date, timedelta

from models import WorkItem, ChartConfig, PALETTES, DEFAULT_PALETTE, STATUS_LABELS
from excel_io import read_excel, create_template_bytes, write_excel
from gantt_renderer import render_gantt, render_gantt_pdf
from block_renderer import render_block_diagram, render_block_pdf
from pptx_export import export_gantt_pptx, export_block_pptx


# ── Constants ────────────────────────────────────────────────────────────────
APP_NAME = "Block & Gantt Creator"
APP_FULL_NAME = "Transformation Office — Block & Gantt Creator Tool"
APP_VERSION = "1.2"


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=APP_FULL_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS — Airbnb-inspired clean design ────────────────────────────────
st.markdown("""
<style>
    /* ── Design tokens ──────────────────────────────────────────── */
    :root {
        --ink: #222222;
        --ink-2: #484848;
        --ink-3: #717171;
        --ink-4: #B0B0B0;
        --line: #EBEBEB;
        --line-soft: #F0F0F0;
        --bg: #FFFFFF;
        --bg-soft: #F7F7F7;
        --accent: #FF385C;
        --accent-2: #E00B41;
        --accent-soft: #FFF1F4;
        --teal: #00A699;
        --gold: #FC642D;
        --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
        --shadow-md: 0 6px 16px rgba(0,0,0,0.06);
        --shadow-lg: 0 12px 32px rgba(0,0,0,0.08);
        --radius-sm: 8px;
        --radius: 12px;
        --radius-lg: 16px;
    }

    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    /* ── Global ─────────────────────────────────────────────────── */
    .stApp {
        background: #FFFFFF;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        color: #222222;
    }

    /* ── Hide Streamlit chrome ───────────────────────────────────── */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1200px; }

    /* ── Header ─────────────────────────────────────────────────── */
    .main-header {
        padding: 0 0 1.5rem 0;
        margin-bottom: 1.75rem;
        border-bottom: 1px solid #EBEBEB;
        position: relative;
    }
    .main-header .header-label {
        font-size: 0.72rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.08em; color: #FF385C; margin-bottom: 0.5rem;
    }
    .main-header h1 {
        font-size: 2rem; font-weight: 700; margin: 0; color: #222222;
        letter-spacing: -0.02em; line-height: 1.2;
    }
    .main-header .subtitle {
        color: #717171; margin: 0.5rem 0 0 0; font-size: 1rem;
        font-weight: 400; line-height: 1.5;
    }
    .main-header .version-badge {
        position: absolute; top: 0; right: 0;
        background: #F7F7F7; color: #484848;
        font-size: 0.7rem; font-weight: 600;
        letter-spacing: 0.02em; padding: 0.3rem 0.7rem;
        border-radius: 999px; border: 1px solid #EBEBEB;
    }

    /* ── Cards ───────────────────────────────────────────────────── */
    .card {
        background: #FFFFFF;
        border: 1px solid #EBEBEB;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: box-shadow 0.2s ease, transform 0.2s ease, border-color 0.2s ease;
        height: 100%;
    }
    .card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.06);
        border-color: #DDDDDD;
    }
    .card h3 {
        margin: 0 0 0.5rem 0; font-weight: 600; font-size: 1.05rem;
        color: #222222;
    }
    .card p {
        color: #484848; font-size: 0.92rem; line-height: 1.55; margin: 0;
    }

    .feature-icon {
        width: 44px; height: 44px; border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.4rem; margin-bottom: 1rem;
        background: #FFF1F4; color: #FF385C;
    }
    .feature-icon.teal { background: #E0F7F5; color: #00A699; }
    .feature-icon.gold { background: #FEEEE5; color: #FC642D; }

    /* ── Markdown elements ───────────────────────────────────────── */
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {
        color: #222222 !important;
        font-weight: 600;
    }
    .stMarkdown p, .stMarkdown li {
        color: #484848;
    }
    .stCaption, [data-testid="stCaptionContainer"] {
        color: #717171 !important;
    }

    /* ── Tabs ────────────────────────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0; border-bottom: 1px solid #EBEBEB;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 0; padding: 12px 20px; font-weight: 600;
        color: #717171; background: transparent;
        border-bottom: 2px solid transparent;
        margin-bottom: -1px;
    }
    .stTabs [aria-selected="true"] {
        color: #222222 !important;
        border-bottom: 2px solid #222222 !important;
        background: transparent !important;
    }

    /* ── Buttons ─────────────────────────────────────────────────── */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px; font-weight: 600;
        padding: 0.55rem 1.2rem; font-size: 0.92rem;
        border: 1px solid #222222;
        background: #FFFFFF; color: #222222;
        transition: all 0.15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #222222; color: #FFFFFF;
    }
    .stButton > button[kind="primary"] {
        background: #FF385C; color: white; border-color: #FF385C;
    }
    .stButton > button[kind="primary"]:hover {
        background: #E00B41; border-color: #E00B41;
    }

    /* ── Sidebar ─────────────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid #EBEBEB;
    }
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #222222 !important; font-size: 0.8rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.06em;
        margin-top: 1rem;
    }
    .sidebar-brand {
        padding: 0.5rem 0 1.25rem 0; margin-bottom: 0.5rem;
        border-bottom: 1px solid #EBEBEB;
    }
    .sidebar-brand .brand-name {
        font-size: 0.95rem; font-weight: 700; color: #222222;
        letter-spacing: -0.01em;
    }
    .sidebar-brand .brand-sub {
        font-size: 0.78rem; color: #717171; margin-top: 0.15rem;
    }

    /* ── Metrics ─────────────────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #EBEBEB;
        border-radius: 12px; padding: 1rem;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.75rem !important; font-weight: 600 !important;
        color: #717171 !important;
    }
    [data-testid="stMetricValue"] {
        font-weight: 700 !important; color: #222222 !important;
    }

    /* ── Steps (numbered list on homepage) ───────────────────────── */
    .step-num {
        display: inline-flex; align-items: center; justify-content: center;
        background: #222222; color: white;
        border-radius: 50%;
        width: 32px; height: 32px; font-weight: 700; font-size: 0.92rem;
        margin-right: 14px; flex-shrink: 0;
    }
    .step-row {
        display: flex; align-items: flex-start; padding: 0.9rem 0;
        border-bottom: 1px solid #F0F0F0; font-size: 0.95rem;
        color: #484848; line-height: 1.55;
    }
    .step-row:last-child { border-bottom: none; }
    .step-row strong { color: #222222; font-weight: 600; }

    /* ── Help boxes ──────────────────────────────────────────────── */
    .help-box {
        background: #F7F7F7;
        border: 1px solid #EBEBEB;
        border-radius: 12px;
        padding: 1.2rem 1.35rem; margin-top: 0.75rem;
    }
    .help-box h4 {
        margin: 0 0 0.6rem 0; color: #222222; font-size: 0.98rem;
        font-weight: 600;
    }
    .help-box p, .help-box li {
        color: #484848; font-size: 0.9rem; line-height: 1.6;
    }
    .help-box ul { margin: 0; padding-left: 1.25rem; }

    /* ── Column reference table ──────────────────────────────────── */
    .column-table {
        width: 100%; border-collapse: collapse; margin: 0.5rem 0;
    }
    .column-table th {
        background: #F7F7F7; padding: 0.7rem 0.85rem;
        text-align: left; font-size: 0.78rem; color: #717171;
        border-bottom: 1px solid #EBEBEB;
        font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em;
    }
    .column-table td {
        padding: 0.7rem 0.85rem; border-bottom: 1px solid #F0F0F0;
        font-size: 0.92rem; color: #484848;
    }
    .column-table td strong { color: #222222; font-weight: 600; }
    .required {
        color: #FF385C; font-weight: 600; font-size: 0.78rem;
        background: #FFF1F4;
        padding: 0.18rem 0.55rem; border-radius: 999px;
    }
    .optional {
        color: #717171; font-weight: 600; font-size: 0.78rem;
        background: #F7F7F7;
        padding: 0.18rem 0.55rem; border-radius: 999px;
    }

    /* ── Section headers ─────────────────────────────────────────── */
    .section-header {
        font-size: 1.35rem; font-weight: 700; color: #222222;
        margin: 2rem 0 0.5rem 0; letter-spacing: -0.01em;
    }
    .section-sub {
        color: #717171; font-size: 0.92rem; margin: 0 0 1rem 0;
    }

    /* ── Footer ──────────────────────────────────────────────────── */
    .app-footer {
        padding: 2.5rem 0 0.75rem;
        border-top: 1px solid #EBEBEB;
        margin-top: 3rem;
    }
    .app-footer p {
        font-size: 0.85rem; color: #717171; margin: 0.2rem 0;
    }
    .app-footer .footer-label {
        font-weight: 700; color: #222222; font-size: 0.9rem;
    }

    /* ── Export section ───────────────────────────────────────────── */
    .export-header {
        font-size: 0.78rem; font-weight: 700; color: #717171;
        text-transform: uppercase; letter-spacing: 0.06em;
        margin: 0.5rem 0 0.75rem 0;
    }

    /* ── Welcome hero ────────────────────────────────────────────── */
    .welcome-hero {
        padding: 1rem 0 1.5rem;
    }
    .welcome-hero h2 {
        font-size: 1.85rem; font-weight: 700; margin: 0;
        color: #222222; letter-spacing: -0.02em; line-height: 1.2;
    }
    .welcome-hero p {
        color: #717171; font-size: 1rem; margin: 0.5rem 0 0 0;
        line-height: 1.55; max-width: 640px;
    }

    /* ── Stat pills on viz page ──────────────────────────────────── */
    .stat-ribbon {
        display: flex; gap: 0.5rem; flex-wrap: wrap;
        margin: 0.5rem 0 1.25rem 0;
    }
    .stat-pill {
        display: inline-flex; align-items: center; gap: 0.4rem;
        padding: 0.35rem 0.85rem; border-radius: 999px;
        font-size: 0.85rem; font-weight: 500;
        background: #F7F7F7; color: #484848;
        border: 1px solid #EBEBEB;
    }
    .stat-pill .pill-val { font-weight: 700; color: #222222; }
    .stat-pill.accent { background: #FFF1F4; border-color: #FFD0DA; }
    .stat-pill.accent .pill-val { color: #FF385C; }

    /* ── Quick-start card ────────────────────────────────────────── */
    .qs-card {
        background: #FFFFFF;
        border: 1px solid #222222;
        border-radius: 12px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1rem;
    }
    .qs-card h3 {
        font-size: 1.05rem; font-weight: 600; color: #222222;
        margin: 0 0 0.35rem 0;
    }
    .qs-card p {
        color: #717171; font-size: 0.92rem; margin: 0;
    }

    /* ── Inputs ──────────────────────────────────────────────────── */
    [data-testid="stTextInput"] input,
    [data-testid="stDateInput"] input,
    [data-baseweb="select"] > div {
        border-radius: 8px !important;
    }

    /* ── Compact preset button row ──────────────────────────────── */
    .preset-row .stButton > button {
        padding: 0.3rem 0.7rem;
        font-size: 0.78rem;
        font-weight: 500;
        border: 1px solid #EBEBEB;
        background: #F7F7F7;
        color: #484848;
    }
    .preset-row .stButton > button:hover {
        background: #222222;
        color: #FFFFFF;
        border-color: #222222;
    }
</style>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────────────

LABEL_TO_STATUS = {v: k for k, v in STATUS_LABELS.items()}


def _items_to_df(items):
    """Convert WorkItems to a DataFrame for st.data_editor."""
    return pd.DataFrame([
        {
            "Title": it.title,
            "Start": it.start_date,
            "End": it.end_date,
            "Category": it.category,
            "Description": it.description,
            "Status": STATUS_LABELS.get(it.status, "Planned"),
            "Owner": it.owner,
            "Label": it.label,
        }
        for it in items
    ])


def _df_to_items(df):
    """Convert an edited DataFrame back to a WorkItem list.

    Returns (items, warnings). Rows with missing/invalid required fields
    are skipped with a warning.
    """
    items = []
    warnings = []
    for idx, row in df.iterrows():
        title = str(row.get("Title", "") or "").strip()
        if not title or title.lower() == "nan":
            continue

        start = row.get("Start")
        end = row.get("End")
        if pd.isna(start) or pd.isna(end):
            warnings.append(f"Row {idx + 1} ('{title}'): missing date — skipped")
            continue

        if hasattr(start, "date") and not isinstance(start, date):
            start = start.date()
        if hasattr(end, "date") and not isinstance(end, date):
            end = end.date()

        if end < start:
            start, end = end, start

        status_label = str(row.get("Status", "Planned") or "Planned")
        status = LABEL_TO_STATUS.get(status_label, "planned")

        items.append(WorkItem(
            title=title,
            start_date=start,
            end_date=end,
            category=str(row.get("Category", "General") or "General").strip() or "General",
            description=str(row.get("Description", "") or "").strip(),
            status=status,
            owner=str(row.get("Owner", "") or "").strip(),
            label=str(row.get("Label", "") or "").strip(),
        ))
    return items, warnings


def _items_signature(items):
    """Hashable summary of items used to detect changes from the editor."""
    return tuple(
        (it.title, it.start_date, it.end_date, it.category, it.description,
         it.status, it.owner, it.label)
        for it in items
    )


def _quarter_bounds(year, quarter):
    """Return (start, end) date of the given calendar quarter."""
    start_month = (quarter - 1) * 3 + 1
    end_month = start_month + 2
    if end_month == 12:
        end = date(year, 12, 31)
    else:
        end = date(year, end_month + 1, 1) - timedelta(days=1)
    return date(year, start_month, 1), end


def _apply_date_preset(preset, items):
    """Compute (start, end) dates for a named preset."""
    today = date.today()
    item_start = min(it.start_date for it in items)
    item_end = max(it.end_date for it in items)

    if preset == "All":
        return item_start, item_end
    if preset == "This Year":
        return date(today.year, 1, 1), date(today.year, 12, 31)
    if preset == "This Quarter":
        q = (today.month - 1) // 3 + 1
        return _quarter_bounds(today.year, q)
    if preset == "Next 6 Months":
        new_month = today.month + 6
        new_year = today.year
        while new_month > 12:
            new_month -= 12
            new_year += 1
        return today, date(new_year, new_month, min(today.day, 28))
    return item_start, item_end


def _build_zip_export(items, config, slide_aspect="16:9"):
    """Bundle every export format into a single zip."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("gantt_chart.png", render_gantt(items, config, dpi=300))
        zf.writestr("gantt_chart.pdf", render_gantt_pdf(items, config, dpi=300))
        zf.writestr("gantt_chart.pptx", export_gantt_pptx(items, config))
        zf.writestr("block_diagram.png",
                    render_block_diagram(items, config, dpi=300, slide_aspect=slide_aspect))
        zf.writestr("block_diagram.pdf",
                    render_block_pdf(items, config, dpi=300, slide_aspect=slide_aspect))
        zf.writestr("block_diagram.pptx", export_block_pptx(items, config))
        zf.writestr("data.xlsx", write_excel(items))
    buf.seek(0)
    return buf.getvalue()


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
    <h1>Block &amp; Gantt Creator</h1>
    <p class="subtitle">Turn a spreadsheet into presentation-ready project visuals in under a minute.</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-name">Transformation Office</div>
        <div class="brand-sub">Block &amp; Gantt Creator</div>
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
        help="Pre-filled template with sample data and an instructions sheet",
    )

    st.markdown("### Step 2 — Upload Your File")
    uploaded_file = st.file_uploader(
        "Drop your Excel file here",
        type=["xlsx", "xls"],
        help="Required columns: Title, Start Date, End Date. Optional: Category, Description, Status, Owner, Label.",
    )

    if uploaded_file:
        try:
            loaded_items, loaded_config, loaded_warnings = read_excel(uploaded_file)
            if not loaded_items:
                st.session_state["load_error"] = (
                    "No work items found. Make sure your file has data rows with "
                    "Title, Start Date, and End Date columns."
                )
                st.session_state["items"] = None
            else:
                st.session_state["items"] = loaded_items
                st.session_state["config"] = loaded_config
                st.session_state["warnings"] = loaded_warnings
                st.session_state["load_error"] = None

                n_cats = len(set(it.category for it in loaded_items))
                st.success(f"Loaded **{len(loaded_items)}** items across **{n_cats}** categories")

                if loaded_warnings:
                    with st.expander(f"Warnings ({len(loaded_warnings)})", expanded=False):
                        for w in loaded_warnings:
                            st.caption(w)
        except Exception as e:
            st.session_state["load_error"] = str(e)
            st.session_state["items"] = None
            st.error("Could not read file — see the home page for troubleshooting tips.")

    if st.session_state["load_error"]:
        with st.expander("Error Details", expanded=True):
            st.code(st.session_state["load_error"], language=None)

    st.divider()

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
        n_preview = min(5, len(palette))
        if n_preview > 0:
            cols = st.columns(n_preview)
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

        # Quick presets
        st.markdown('<div class="preset-row">', unsafe_allow_html=True)
        p1, p2, p3, p4 = st.columns(4)
        preset_clicked = None
        with p1:
            if st.button("All", key="preset_all", use_container_width=True):
                preset_clicked = "All"
        with p2:
            if st.button("Year", key="preset_year", use_container_width=True,
                         help="This calendar year"):
                preset_clicked = "This Year"
        with p3:
            if st.button("Quarter", key="preset_quarter", use_container_width=True,
                         help="Current calendar quarter"):
                preset_clicked = "This Quarter"
        with p4:
            if st.button("6mo", key="preset_6mo", use_container_width=True,
                         help="Next 6 months from today"):
                preset_clicked = "Next 6 Months"
        st.markdown('</div>', unsafe_allow_html=True)

        if preset_clicked:
            new_start, new_end = _apply_date_preset(preset_clicked, st.session_state["items"])
            config.start_date = new_start
            config.end_date = new_end
            st.session_state["config"] = config
            st.rerun()

        d_col1, d_col2 = st.columns(2)
        new_start = config.start_date
        new_end = config.end_date
        with d_col1:
            if config.start_date:
                new_start = st.date_input("Start", value=config.start_date)
        with d_col2:
            if config.end_date:
                new_end = st.date_input("End", value=config.end_date)

        # Validate date range — auto-swap and warn rather than rendering empty chart
        if config.start_date and config.end_date:
            if new_start > new_end:
                st.warning("Start date is after end date — swapping for you.")
                new_start, new_end = new_end, new_start
            config.start_date = new_start
            config.end_date = new_end

        st.session_state["config"] = config

    # Sidebar footer
    st.divider()
    st.markdown(f"""
    <div style="text-align: center; padding: 0.25rem 0;">
        <div style="font-size: 0.7rem; color: #717171; font-weight: 600;">Version {APP_VERSION}</div>
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
        <h2>Build a roadmap from your spreadsheet.</h2>
        <p>Upload your project data, pick a color palette, and download presentation-ready charts. No installs, no logins, no data leaves your machine.</p>
    </div>
    """, unsafe_allow_html=True)

    # Feature cards
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="card">
            <div class="feature-icon">&#x1F4CB;</div>
            <h3>Simple input</h3>
            <p>Just three required columns &mdash; Title, Start Date, End Date. Add Category for color-coded workstreams. Flexible column names supported.</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
            <div class="feature-icon teal">&#x1F4CA;</div>
            <h3>Two chart types</h3>
            <p><strong>Gantt</strong> shows workstreams as swim lanes over time. <strong>Block diagram</strong> packs your whole portfolio into a single dense slide.</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="card">
            <div class="feature-icon gold">&#x1F4E4;</div>
            <h3>Three export formats</h3>
            <p><strong>PowerPoint</strong> with fully editable native shapes, <strong>PDF</strong> for documents, and <strong>PNG</strong> at 300 DPI for emails and decks.</p>
        </div>
        """, unsafe_allow_html=True)

    # ── Step-by-step guide ────────────────────────────────────────────────
    st.markdown('<div class="section-header">How it works</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Four steps from spreadsheet to slide.</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <div class="step-row">
            <span class="step-num">1</span>
            <div>
                <strong>Download the template.</strong> Use the &ldquo;Download Excel Template&rdquo; button in the sidebar. It comes pre-loaded with sample data so you can see the format.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">2</span>
            <div>
                <strong>Fill in your data.</strong> Open it in Excel or Google Sheets and swap the sample rows for your real initiatives. Three columns is enough to get started.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">3</span>
            <div>
                <strong>Upload.</strong> Drag your .xlsx file into the sidebar uploader. The tool validates and previews instantly.
            </div>
        </div>
        <div class="step-row">
            <span class="step-num">4</span>
            <div>
                <strong>Customize and export.</strong> Pick a palette, adjust the date range, then download as PowerPoint, PDF, or PNG.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Excel format reference ────────────────────────────────────────────
    st.markdown('<div class="section-header">Spreadsheet format</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Column names are flexible &mdash; &ldquo;Name&rdquo; works in place of &ldquo;Title&rdquo;, &ldquo;Team&rdquo; in place of &ldquo;Category&rdquo;, etc.</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
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
                    <td>Workstream, team, or department (color + lane grouping)</td>
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
                    <td>Short identifier shown on the block (e.g., &ldquo;D1&rdquo;, &ldquo;MVP&rdquo;)</td>
                    <td>M2</td>
                </tr>
            </tbody>
        </table>
    </div>
    """, unsafe_allow_html=True)

    # ── Try sample data ──────────────────────────────────────────────────
    st.markdown('<div class="section-header">Try it without a file</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="qs-card">
        <h3>Load sample data</h3>
        <p>Explore the tool with a 25-item product roadmap across 6 workstreams.</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Load Sample Data", type="primary"):
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

    # ── Troubleshooting ──────────────────────────────────────────────────
    st.markdown('<div class="section-header">Troubleshooting</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>File won't upload</h4>
        <ul>
            <li>Ensure your file is in <strong>.xlsx</strong> format (not .csv, .xls, or .numbers)</li>
            <li>The file must be under 200 MB</li>
            <li>If using Google Sheets, export as .xlsx first: <em>File &rarr; Download &rarr; Microsoft Excel</em></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>&ldquo;Missing required column&rdquo; error</h4>
        <ul>
            <li>Your file needs at least these columns: <strong>Title</strong>, <strong>Start Date</strong>, <strong>End Date</strong></li>
            <li>Column names are flexible &mdash; &ldquo;Name&rdquo;, &ldquo;Task&rdquo;, or &ldquo;Item&rdquo; all map to Title</li>
            <li>Column headers must be in <strong>Row 1</strong> of your spreadsheet</li>
            <li>Avoid merged cells in the header row</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>Date parsing errors</h4>
        <ul>
            <li>Supported formats: <code>2025-01-15</code>, <code>01/15/2025</code>, <code>1/15/25</code></li>
            <li>Date cells must be formatted as <strong>Date</strong> in Excel, not as plain text</li>
            <li>Remove blank rows rather than leaving empty cells between data</li>
            <li>If Start Date is after End Date, the tool will swap them and flag a warning</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="help-box">
        <h4>Chart doesn't look right</h4>
        <ul>
            <li>Adjust the <strong>Date Range</strong> in the sidebar to focus on the relevant period</li>
            <li>Try a different <strong>Color Palette</strong> if categories are hard to distinguish</li>
            <li>For the Block Diagram, try the <strong>4:3 aspect ratio</strong> if blocks appear too narrow</li>
            <li>Category names must match exactly &mdash; &ldquo;Marketing&rdquo; and &ldquo;marketing&rdquo; are treated as separate groups</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # ── Footer ────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="app-footer">
        <p class="footer-label">Transformation Office</p>
        <p>Block &amp; Gantt Creator &middot; v{APP_VERSION}</p>
        <p>Your data stays on your computer. Nothing is uploaded.</p>
    </div>
    """, unsafe_allow_html=True)


def show_visualizations():
    """Main visualization view with Gantt, Block, and Data tabs."""
    items = st.session_state["items"]
    config = st.session_state["config"]

    # Single-pass stats computation (was 6 separate O(n) passes)
    categories = set()
    min_date = items[0].start_date
    max_date = items[0].end_date
    in_progress_count = 0
    done_count = 0
    at_risk_count = 0
    milestones = 0
    for it in items:
        categories.add(it.category)
        if it.start_date < min_date:
            min_date = it.start_date
        if it.end_date > max_date:
            max_date = it.end_date
        if it.status == "in_progress":
            in_progress_count += 1
        elif it.status == "done":
            done_count += 1
        elif it.status == "at_risk":
            at_risk_count += 1
        if it.is_milestone:
            milestones += 1
    span_months = max(1, (max_date - min_date).days // 30)

    pills = [
        ('', len(items), 'items'),
        ('', len(categories), 'categories'),
        ('', span_months, 'months'),
        ('accent', in_progress_count, 'in progress'),
        ('', done_count, 'done'),
    ]
    if at_risk_count > 0:
        pills.append(('accent', at_risk_count, 'at risk'))
    if milestones > 0:
        pills.append(('', milestones, 'milestones'))

    pill_html = "".join(
        f'<span class="stat-pill {cls}"><span class="pill-val">{val}</span> {label}</span>'
        for cls, val, label in pills
    )
    st.markdown(f'<div class="stat-ribbon">{pill_html}</div>', unsafe_allow_html=True)

    # Top action row — batch export + clear data
    action_l, action_mid, action_r = st.columns([4, 2, 1])
    with action_mid:
        try:
            zip_bytes = _build_zip_export(items, config, slide_aspect="16:9")
            st.download_button(
                "Download All Formats (.zip)",
                data=zip_bytes,
                file_name="transformation_office_export.zip",
                mime="application/zip",
                help="One zip with PNG, PDF, PowerPoint, and cleaned Excel data",
                use_container_width=True,
            )
        except Exception as e:
            st.button(
                "Download All Formats (.zip)",
                disabled=True,
                help=f"Export failed: {e}",
                use_container_width=True,
            )
    with action_r:
        if st.button("Clear Data", help="Remove loaded data and return to the home screen",
                     use_container_width=True):
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
        st.markdown("#### Swim Lane Gantt")
        st.caption("Each category gets its own lane. Overlapping tasks stack automatically.")

        try:
            with st.spinner("Rendering Gantt chart..."):
                gantt_preview = render_gantt(items, config, dpi=150)
            st.image(gantt_preview, use_container_width=True)

            st.markdown("---")
            st.markdown('<div class="export-header">Export</div>', unsafe_allow_html=True)
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
            st.error(f"Couldn't render the Gantt chart: {e}")
            with st.expander("Show full error"):
                st.code(traceback.format_exc())

    # ── BLOCK DIAGRAM TAB ────────────────────────────────────────────────
    with tab_block:
        st.markdown("#### Space-Filling Block Diagram")
        st.caption("Your entire portfolio on a single slide. Great for executive summaries.")

        aspect_ratio = st.selectbox(
            "Slide Aspect Ratio",
            options=["16:9", "4:3"],
            index=0,
            help="16:9 for modern widescreen displays. 4:3 for older projectors.",
        )

        try:
            with st.spinner("Rendering block diagram..."):
                block_preview = render_block_diagram(items, config, dpi=150, slide_aspect=aspect_ratio)
            st.image(block_preview, use_container_width=True)

            st.markdown("---")
            st.markdown('<div class="export-header">Export</div>', unsafe_allow_html=True)
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
            st.error(f"Couldn't render the block diagram: {e}")
            with st.expander("Show full error"):
                st.code(traceback.format_exc())

    # ── DATA PREVIEW TAB ─────────────────────────────────────────────────
    with tab_data:
        st.markdown("#### Edit Data")
        st.caption(
            "Edit any cell, add rows with the + button below the table, or delete rows "
            "by selecting them and pressing Delete. Changes update the charts instantly."
        )

        editable_df = _items_to_df(items)
        edited_df = st.data_editor(
            editable_df,
            use_container_width=True,
            num_rows="dynamic",
            height=min(640, 60 + len(editable_df) * 36),
            column_config={
                "Title": st.column_config.TextColumn(
                    "Title", required=True, help="Name of the work item"),
                "Start": st.column_config.DateColumn(
                    "Start", required=True, help="When work begins"),
                "End": st.column_config.DateColumn(
                    "End", required=True, help="When work is expected to finish"),
                "Category": st.column_config.TextColumn(
                    "Category", help="Workstream or team — drives color and lane grouping"),
                "Description": st.column_config.TextColumn(
                    "Description", help="Optional short description shown on blocks"),
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=list(STATUS_LABELS.values()),
                    required=True,
                ),
                "Owner": st.column_config.TextColumn("Owner"),
                "Label": st.column_config.TextColumn(
                    "Label", help="Short identifier shown on the block (e.g., D1, MVP)"),
            },
            key="data_editor",
        )

        new_items, edit_warnings = _df_to_items(edited_df)
        if _items_signature(new_items) != _items_signature(items):
            if not new_items:
                st.warning(
                    "All rows have been removed. Add at least one item or click "
                    "**Clear Data** to return to the home page."
                )
            else:
                st.session_state["items"] = new_items
                if edit_warnings:
                    st.session_state["warnings"] = edit_warnings
                st.rerun()

        # Download cleaned data back to Excel
        dl_col, _spacer = st.columns([1, 3])
        with dl_col:
            try:
                xlsx_bytes = write_excel(items)
                st.download_button(
                    "Download as Excel",
                    data=xlsx_bytes,
                    file_name="transformation_office_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Save your edited data back to a .xlsx file you can share",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Couldn't generate Excel export: {e}")

        # Summary by category
        st.markdown("#### Summary by Category")
        summary_df = pd.DataFrame([
            {
                "Category": it.category,
                "Start": it.start_date.strftime("%Y-%m-%d"),
                "End": it.end_date.strftime("%Y-%m-%d"),
            }
            for it in items
        ])
        if not summary_df.empty:
            summary = summary_df.groupby("Category").agg(
                Items=("Category", "count"),
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
        <p>Block &amp; Gantt Creator &middot; v{APP_VERSION}</p>
    </div>
    """, unsafe_allow_html=True)


# ── Route to homepage or visualization ───────────────────────────────────────
if st.session_state["items"] is not None and len(st.session_state["items"]) > 0:
    show_visualizations()
else:
    show_homepage()
