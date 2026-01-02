from __future__ import annotations


# =============================================================================
# Teaching notes: app.py (Streamlit UI)
#
# This file is the "front door" of the application. It defines the Streamlit UI:
#   - Upload an Excel workbook (source of truth)
#   - Edit Settings / Workstreams / Tasks in friendly tables
#   - Preview the roadmap chart
#   - Export to PDF, PNG, and editable PPTX
#
# Beginner reading tip:
#   1) Skim the helper functions at the top (_df_clean, _is_blank, etc.).
#   2) Jump to main() near the bottom to see how the UI is assembled.
#   3) When you see a call like export_png_bytes(...), follow it into export.py
#      and then renderer.py / pptx_export.py.
#
# Safety rule for this app:
#   - When a new workbook is uploaded, we hard-replace ALL in-app data.
#     No merging, no leftovers. That keeps the workflow predictable.
# =============================================================================

# ---------------------------------------------------------------------------
# Imports
#
# A quick note for beginners:
# - "standard library" modules ship with Python (no extra install needed)
# - "third-party" modules are installed via requirements.txt (pip install ...)
# - "local" imports come from files in this repo (excel_io.py, renderer.py, etc.)
# ---------------------------------------------------------------------------
import hashlib
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from pydantic import ValidationError

from excel_io import (
    ExcelPayload,
    TASK_COLUMNS,
    WORKSTREAM_COLUMNS,
    read_roadmap_excel,
    template_bytes,
    write_roadmap_excel_bytes,
)
from export import export_pdf_bytes, export_png_bytes, export_pptx_bytes, preview_png_bytes
from roadmap_models import COLOR_DROPDOWN_VALUES, COLOR_NAME_TO_HEX, Settings, Task, Workstream
from scheduler import schedule_by_workstream


APP_TITLE = "Transformation Office Block & Gantt Creator Tool"
APP_SUBTITLE = "Excel → executive-ready blocked swimlane roadmap → export PDF + high‑res PNG + editable PPTX"


# ----------------------------
# Caching helpers
# ----------------------------


@st.cache_data(show_spinner=False)
def _cached_read_excel(excel_bytes: bytes) -> ExcelPayload:
    """Cache wrapper: parse uploaded Excel bytes into our internal payload object."""
    return read_roadmap_excel(excel_bytes)


@st.cache_data(show_spinner=False)
def _cached_template_bytes() -> bytes:
    """Cache wrapper: build the downloadable Excel template once and reuse it."""
    return template_bytes()


@st.cache_data(show_spinner=False)
def _cached_preview(
    settings_dump: Dict[str, Any],
    workstreams_dump: List[Dict[str, Any]],
    tasks_dump: List[Dict[str, Any]],
    *,
    include_out_of_range: bool,
    dpi: int,
) -> bytes:
    """Cache wrapper: build a small preview PNG for the current in-app data."""
    settings = Settings(**settings_dump)
    workstreams = [Workstream(**w) for w in workstreams_dump]
    tasks = [Task(**t) for t in tasks_dump]
    return preview_png_bytes(settings, workstreams, tasks, include_out_of_range=include_out_of_range, dpi=dpi)


@st.cache_data(show_spinner=False)
def _cached_export_pdf(
    settings_dump: Dict[str, Any],
    workstreams_dump: List[Dict[str, Any]],
    tasks_dump: List[Dict[str, Any]],
    *,
    include_out_of_range: bool,
) -> bytes:
    """Cache wrapper: render and export a PDF (vector) for the current in-app data."""
    settings = Settings(**settings_dump)
    workstreams = [Workstream(**w) for w in workstreams_dump]
    tasks = [Task(**t) for t in tasks_dump]
    return export_pdf_bytes(settings, workstreams, tasks, include_out_of_range=include_out_of_range)


@st.cache_data(show_spinner=False)
def _cached_export_png(
    settings_dump: Dict[str, Any],
    workstreams_dump: List[Dict[str, Any]],
    tasks_dump: List[Dict[str, Any]],
    *,
    include_out_of_range: bool,
    dpi: int,
) -> bytes:
    """Cache wrapper: render and export a high-resolution PNG for the current in-app data."""
    settings = Settings(**settings_dump)
    workstreams = [Workstream(**w) for w in workstreams_dump]
    tasks = [Task(**t) for t in tasks_dump]
    return export_png_bytes(settings, workstreams, tasks, include_out_of_range=include_out_of_range, dpi=dpi)


@st.cache_data(show_spinner=False)
def _cached_export_pptx(
    settings_dump: Dict[str, Any],
    workstreams_dump: List[Dict[str, Any]],
    tasks_dump: List[Dict[str, Any]],
    *,
    include_out_of_range: bool,
) -> bytes:
    """Cache wrapper: render and export an editable PPTX for the current in-app data."""
    settings = Settings(**settings_dump)
    workstreams = [Workstream(**w) for w in workstreams_dump]
    tasks = [Task(**t) for t in tasks_dump]
    return export_pptx_bytes(settings, workstreams, tasks, include_out_of_range=include_out_of_range)


# ----------------------------
# Data prep / validation
# ----------------------------


def _df_clean(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize a DataFrame so missing values are plain Python None (safer for UI + validation)."""
    df2 = df.copy()
    # Normalize pandas missing sentinels to Python None.
    # NOTE: pd.NA is "poisonous" in boolean contexts (e.g., `x or ""`,
    # `x in (...)`) and will raise: "boolean value of NA is ambiguous".
    # Converting early prevents accidental truthiness checks from exploding.
    df2 = df2.replace({pd.NA: None})

    # Also convert numpy NaN/NaT to None for consistent downstream logic.
    # This will upcast some extension dtypes to object, which is fine for our
    # UI-centric workflow.
    try:
        df2 = df2.where(pd.notna(df2), None)
    except Exception:
        pass

    df2 = df2.dropna(how="all")
    return df2.reset_index(drop=True)


def _is_blank(value: Any) -> bool:
    """True if value is None/NaN/NaT/pd.NA or an empty/whitespace string."""

    if value is None:
        return True
    # pandas missing (covers pd.NA, np.nan, NaT)
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str):
        return value.strip() == ""
    return False


def _to_str(value: Any) -> str:
    """Convert a value to a safe string for UI display (never returns None)."""
    if _is_blank(value):
        return ""
    return str(value)


def _to_str_or_none(value: Any) -> Optional[str]:
    """Convert a value to a string, but return None if the input is blank."""
    s = _to_str(value).strip()
    return s or None


def _ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Make sure the DataFrame has the expected columns (adds missing columns as blank)."""
    df2 = df.copy()
    for c in cols:
        if c not in df2.columns:
            df2[c] = pd.NA
    return df2[cols]


def _coerce_text_cols_for_editor(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Force text-like columns to a real editable string dtype.

    Pandas will often infer an all-empty Excel column as float (all NaNs).
    Streamlit's data_editor then refuses TextColumn/SelectboxColumn edits.

    This function guarantees the column kind is string-compatible.
    """
    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            # pandas StringDtype keeps missing as <NA>, but we want blanks for editing
            df2[c] = df2[c].astype("string").fillna("")
    return df2


# ----------------------------
# Color dropdown coercion
# ----------------------------


def _normalize_color_token(value: str) -> str:
    """Normalize a color name token (trim, lowercase, collapse whitespace)."""
    s = (value or "").strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    return s


# Canonical dropdown names (for display) keyed by normalized token.
_COLOR_DISPLAY_CANONICAL: Dict[str, str] = {
    _normalize_color_token(name): name for name in COLOR_DROPDOWN_VALUES
}


def _hex_upper(s: str) -> str:
    """Normalize a hex color string (ensure leading # and uppercase)."""
    s = (s or "").strip()
    if not s:
        return ""
    if not s.startswith("#"):
        s = "#" + s
    return s.upper()


# Map known palette hex values back to the dropdown display names.
_HEX_TO_DISPLAY: Dict[str, str] = {}
for _name in COLOR_DROPDOWN_VALUES:
    token = _normalize_color_token(_name)
    hx = COLOR_NAME_TO_HEX.get(token)
    if hx:
        _HEX_TO_DISPLAY[_hex_upper(hx)] = _name


def _coerce_color_cols_for_editor(df: pd.DataFrame, cols: List[str]) -> tuple[pd.DataFrame, List[str]]:
    """Force color columns to the supported dropdown values.

    We only allow the dropdown choices in the Streamlit editor to keep the
    UX simple for non-technical users.

    If a workbook contains:
      - blanks -> becomes "Auto"
      - legacy tokens (Primary/Secondary/Accent/Neutral) -> mapped to nearest
        dropdown color name
      - hex colors -> mapped when the hex matches a known palette color;
        otherwise becomes "Auto" and a warning is emitted
    """

    import re

    HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")

    warnings: List[str] = []
    replaced_custom_hex = 0

    df2 = df.copy()

    for c in cols:
        if c not in df2.columns:
            continue

        def _coerce(v: Any) -> str:
            nonlocal replaced_custom_hex
            if _is_blank(v):
                return "Auto"
            s = str(v).strip()
            if not s:
                return "Auto"

            token = _normalize_color_token(s)
            if token in _COLOR_DISPLAY_CANONICAL:
                return _COLOR_DISPLAY_CANONICAL[token]

            # Legacy friendly names and other tokens understood by validators.
            if token in COLOR_NAME_TO_HEX:
                hx = COLOR_NAME_TO_HEX.get(token)
                if not hx:
                    return "Auto"
                disp = _HEX_TO_DISPLAY.get(_hex_upper(hx))
                return disp or "Auto"

            # Hex colors: map if it matches a known palette color; otherwise fall back.
            if HEX_RE.match(s):
                hx = _hex_upper(s)
                if hx in _HEX_TO_DISPLAY:
                    return _HEX_TO_DISPLAY[hx]
                replaced_custom_hex += 1
                return "Auto"

            return "Auto"

        df2[c] = df2[c].apply(_coerce).astype("string").fillna("Auto")

    if replaced_custom_hex:
        warnings.append(
            f"Found {replaced_custom_hex} custom hex color(s) in the uploaded file. "
            "They were replaced with 'Auto' in the editor. Choose a named color from the dropdown if you want something specific."
        )

    return df2, warnings


def _coerce_date_cols_for_editor(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Ensure date columns are object dtype containing python date or None."""

    def _to_date(v: Any) -> Optional[date]:
        if v is None:
            return None
        if isinstance(v, float) and pd.isna(v):
            return None
        if isinstance(v, date) and not isinstance(v, datetime):
            return v
        if isinstance(v, datetime):
            return v.date()
        # pandas Timestamp
        if hasattr(v, "to_pydatetime"):
            try:
                dt = v.to_pydatetime()
                if isinstance(dt, datetime):
                    return dt.date()
            except Exception:
                return None
        return None

    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            df2[c] = df2[c].apply(_to_date).astype(object)
    return df2



def _coerce_date_cols_to_text_for_editor(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Represent date columns as ISO strings for easier copy/paste and deletion.

    Streamlit's DateColumn editing can be finicky across browsers/platforms.
    Using text dates (YYYY-MM-DD) is robust and still validates strictly before export.
    """

    def _fmt(v: Any) -> str:
        if _is_blank(v):
            return ""
        if isinstance(v, datetime):
            return v.date().isoformat()
        if isinstance(v, date) and not isinstance(v, datetime):
            return v.isoformat()
        # pandas Timestamp
        if hasattr(v, "to_pydatetime"):
            try:
                dt = v.to_pydatetime()
                if isinstance(dt, datetime):
                    return dt.date().isoformat()
            except Exception:
                return ""
        if isinstance(v, str):
            return v.strip()
        try:
            return str(v)
        except Exception:
            return ""

    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            df2[c] = df2[c].apply(_fmt).astype("string").fillna("")
    return df2


def _stable_auto_id(row: Dict[str, Any], *, used: set[str]) -> str:
    """Deterministic-ish ID for blank rows (stable across re-renders)."""

    # IMPORTANT: avoid `value or ""` here because pandas' pd.NA is not
    # truthy/falsey (it raises "boolean value of NA is ambiguous").
    parts = [
        _to_str(row.get("workstream")).strip(),
        _to_str(row.get("title")).strip(),
        _to_str(row.get("description")).strip(),
        _to_str(row.get("start_date")).strip(),
        _to_str(row.get("end_date")).strip(),
        _to_str(row.get("status")).strip(),
        _to_str(row.get("owner")).strip(),
        _to_str(row.get("type")).strip(),
        _to_str(row.get("hyperlink")).strip(),
    ]
    base = "|".join(parts).encode("utf-8", errors="ignore")
    h = hashlib.sha1(base).hexdigest()[:10].upper()
    candidate = f"AUTO-{h}"
    if candidate not in used:
        used.add(candidate)
        return candidate
    i = 2
    while f"{candidate}-{i}" in used:
        i += 1
    candidate2 = f"{candidate}-{i}"
    used.add(candidate2)
    return candidate2


def _build_models(
    settings_raw: Dict[str, Any],
    workstreams_df: pd.DataFrame,
    tasks_df: pd.DataFrame,
) -> Tuple[Optional[Settings], List[Workstream], List[Task], List[str], List[str]]:
    """Returns: (settings, workstreams, tasks, errors, warnings)."""

    errors: List[str] = []
    warnings: List[str] = []

    # -------------------------
    # Settings
    # -------------------------
    s_dict: Dict[str, Any] = dict(settings_raw or {})

    # Normalize output_dpi from Excel numeric/text
    if isinstance(s_dict.get("output_dpi"), str) and s_dict["output_dpi"].strip().isdigit():
        s_dict["output_dpi"] = int(s_dict["output_dpi"].strip())
    if isinstance(s_dict.get("output_dpi"), float) and not pd.isna(s_dict["output_dpi"]):
        s_dict["output_dpi"] = int(s_dict["output_dpi"])

    def _norm_enum(v: Any) -> Any:
        if v is None:
            return v
        s = str(v).strip()
        if not s:
            return None
        if s.lower() == "monday":
            return "Mon"
        if s.lower() == "sunday":
            return "Sun"
        if s.lower() in {"mon", "sun"}:
            return s[:1].upper() + s[1:].lower()
        if s.lower() in {"weekly", "monthly"}:
            return s[:1].upper() + s[1:].lower()
        if s.lower() in {"a3", "a4"}:
            return s.upper()
        return s

    for k in ("week_start_day", "time_granularity", "page_size"):
        if k in s_dict:
            s_dict[k] = _norm_enum(s_dict.get(k))

    try:
        settings = Settings(**s_dict)
    except ValidationError as ve:
        settings = None
        for err in ve.errors():
            loc = ".".join([str(x) for x in err.get("loc", [])])
            msg = err.get("msg", "Invalid value")
            errors.append(f"Settings: {loc} — {msg}")
    except Exception as e:
        settings = None
        errors.append(f"Settings: {e}")

    # -------------------------
    # Workstreams
    # -------------------------
    ws_df = _df_clean(_ensure_columns(workstreams_df, WORKSTREAM_COLUMNS))
    workstreams: List[Workstream] = []
    ws_names: List[str] = []

    for idx, row in ws_df.iterrows():
        rec = {c: row.get(c) for c in WORKSTREAM_COLUMNS}
        ws_name = _to_str(rec.get("workstream")).strip()
        if not ws_name:
            continue
        try:
            order_val = rec.get("order")
            order = None
            if not _is_blank(order_val):
                try:
                    order = int(order_val)
                except Exception:
                    # Common case: user typed a number as text
                    s = _to_str(order_val).strip()
                    if s.isdigit():
                        order = int(s)
                    else:
                        raise ValueError("order must be an integer")

            color_val = rec.get("color")
            color = None
            color = _to_str_or_none(color_val)

            ws = Workstream(workstream=ws_name, order=order, color=color)
            workstreams.append(ws)
            ws_names.append(ws.workstream)
        except ValidationError as ve:
            for err in ve.errors():
                loc = ".".join([str(x) for x in err.get("loc", [])])
                msg = err.get("msg", "Invalid value")
                errors.append(f"Workstreams row {idx+2}: {loc} — {msg}")
        except Exception as e:
            errors.append(f"Workstreams row {idx+2}: {e}")

    dups = sorted({n for n in ws_names if ws_names.count(n) > 1})
    if dups:
        errors.append(f"Workstreams: workstream names must be unique. Duplicates: {', '.join(dups)}")

    any_order = any(ws.order is not None for ws in workstreams)
    if any_order:
        workstreams = sorted(workstreams, key=lambda w: (w.order if w.order is not None else 9999, w.workstream.lower()))
    else:
        workstreams = sorted(workstreams, key=lambda w: w.workstream.lower())

    ws_set = {w.workstream for w in workstreams}

    # -------------------------
    # Tasks
    # -------------------------
    t_df = _df_clean(_ensure_columns(tasks_df, TASK_COLUMNS))

    if not ws_set and not t_df.empty:
        errors.append("Workstreams: define at least one workstream. Tasks must reference Workstreams.workstream.")

    used_ids: set[str] = set()
    for _, row in t_df.iterrows():
        rid = row.get("id")
        if _is_blank(rid):
            continue
        s = _to_str(rid).strip()
        if s:
            used_ids.add(s)

    tasks: List[Task] = []
    for idx, row in t_df.iterrows():
        rec = {c: row.get(c) for c in TASK_COLUMNS}

        # Skip fully blank-ish rows
        if all(_is_blank(rec.get(c)) for c in ("workstream", "title", "start_date", "end_date", "id")):
            continue

        rid = rec.get("id")
        rid_s = _to_str(rid).strip()
        if not rid_s:
            rid_s = _stable_auto_id(rec, used=used_ids)

        ws_name = _to_str(rec.get("workstream")).strip()
        if not ws_name:
            errors.append(f"Tasks row {idx+2}: workstream is required.")
            continue
        if ws_set and ws_name not in ws_set:
            errors.append(f"Tasks row {idx+2}: unknown workstream '{ws_name}'. Add it to Workstreams.")
            continue

        def _coerce_date_val(v: Any) -> Optional[date]:
            if _is_blank(v):
                return None
            if isinstance(v, date) and not isinstance(v, datetime):
                return v
            if isinstance(v, datetime):
                return v.date()
            if hasattr(v, "to_pydatetime"):
                try:
                    dt = v.to_pydatetime()
                    if isinstance(dt, datetime):
                        return dt.date()
                except Exception:
                    return None
            if isinstance(v, str):
                s = v.strip()
                if not s:
                    return None
                if s.lower() in {"none", "null", "nan", "nat"}:
                    return None
                for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%b %d %Y"):
                    try:
                        return datetime.strptime(s, fmt).date()
                    except ValueError:
                        continue
            return None

        sd = _coerce_date_val(rec.get("start_date"))
        ed = _coerce_date_val(rec.get("end_date"))
        if sd is None:
            errors.append(f"Tasks row {idx+2}: start_date is required. Enter a date like 2026-01-15.")
            continue
        if ed is None:
            errors.append(f"Tasks row {idx+2}: end_date is required. Enter a date like 2026-01-15.")
            continue

        try:
            desc = _to_str_or_none(rec.get("description"))

            status = _to_str_or_none(rec.get("status"))
            if status is not None:
                status = status.strip().lower()

            owner = _to_str_or_none(rec.get("owner"))

            color_override = _to_str_or_none(rec.get("color_override"))

            type_val = _to_str_or_none(rec.get("type"))
            ttype = "block" if type_val is None else type_val.strip().lower()

            hyperlink = _to_str_or_none(rec.get("hyperlink"))

            t = Task(
                id=rid_s,
                workstream=ws_name,
                title=_to_str(rec.get("title")).strip(),
                description=desc,
                start_date=sd,
                end_date=ed,
                status=status,
                owner=owner,
                color_override=color_override,
                type=ttype,  # validated by pydantic
                hyperlink=hyperlink,
            )
            tasks.append(t)
        except ValidationError as ve:
            for err in ve.errors():
                loc = ".".join([str(x) for x in err.get("loc", [])])
                msg = err.get("msg", "Invalid value")
                errors.append(f"Tasks row {idx+2}: {loc} — {msg}")
        except Exception as e:
            errors.append(f"Tasks row {idx+2}: {e}")

    if not tasks and t_df.empty:
        warnings.append("No tasks found. Add rows to the Tasks sheet.")

    # -------------------------
    # Cross checks + scheduling
    # -------------------------
    if settings is not None and tasks:
        for t in tasks:
            if t.end_date < settings.overall_start_date or t.start_date > settings.overall_end_date:
                warnings.append(f"Task {t.id} is completely outside the overall date range (hidden by default).")
            elif t.start_date < settings.overall_start_date or t.end_date > settings.overall_end_date:
                warnings.append(f"Task {t.id} is partially outside the overall date range (will be clamped).")

    if not errors and tasks:
        scheduled = schedule_by_workstream(tasks, touching_counts_as_overlap=True)
        by_id = {t.id: t for ws_tasks in scheduled.values() for t in ws_tasks}
        tasks = [by_id.get(t.id, t) for t in tasks]

    return settings, workstreams, tasks, errors, warnings


# ----------------------------
# UI
# ----------------------------


def _inject_css() -> None:
    """Inject small CSS tweaks to make the Streamlit UI feel more polished."""
    st.markdown(
        """
<style>
  /* Tighter top padding */
  div.block-container { padding-top: 1.2rem; }

  /* Make sidebar feel a bit more like a control panel */
  [data-testid="stSidebar"] { border-right: 1px solid rgba(49, 51, 63, 0.10); }

  /* Slightly larger buttons for non-technical users */
  .stDownloadButton button, .stButton button { padding: 0.55rem 0.9rem; }

  /* Hide Streamlit footer */
  footer { visibility: hidden; }
</style>
        """,
        unsafe_allow_html=True,
    )


def _current_editor_keys() -> tuple[str, str]:
    """Return stable widget keys for the currently loaded dataset.

    Important: st.data_editor stores its state by key. If a user loads a new
    workbook, we must guarantee a brand-new key (and purge old ones), otherwise
    Streamlit can "replay" prior edits onto the new dataframe.

    We use a monotonically increasing epoch that bumps on every load/reset.
    """

    epoch = int(st.session_state.get("_data_epoch", 0))
    return (f"ws_editor_{epoch}", f"tasks_editor_{epoch}")


def _purge_editor_state() -> None:
    """Remove all cached editor widget states from session_state.

    This is the defensive sledgehammer that ensures *zero* rows from a previous
    workbook can leak into the newly uploaded workbook.
    """

    for k in list(st.session_state.keys()):
        if k.startswith("ws_editor_") or k.startswith("tasks_editor_"):
            st.session_state.pop(k, None)
    # Legacy keys from earlier builds
    st.session_state.pop("ws_editor", None)
    st.session_state.pop("tasks_editor", None)


def _bump_epoch() -> int:
    """Increment the internal epoch counter so editor widgets re-mount with fresh state."""
    st.session_state["_data_epoch"] = int(st.session_state.get("_data_epoch", 0)) + 1
    return int(st.session_state["_data_epoch"])


def _load_payload_into_state(payload: ExcelPayload, *, source_name: str, upload_hash: str) -> None:
    """Hard-replace all editable app data with the uploaded workbook payload."""

    # 1) Purge any editor/widget state so nothing can survive the swap.
    _purge_editor_state()

    # 2) Bump epoch so new editors always get brand-new keys.
    _bump_epoch()

    # 3) Replace the active workbook identity
    st.session_state["_last_upload_hash"] = upload_hash
    st.session_state["_active_workbook_name"] = source_name
    st.session_state["_loaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 4) Replace data (no merging, no appending)
    st.session_state["settings_raw"] = dict(payload.settings)
    st.session_state["workstreams_df"] = payload.workstreams_df.copy()
    st.session_state["tasks_df"] = payload.tasks_df.copy()

    # 5) Originals for reset
    st.session_state["_orig_settings_raw"] = dict(payload.settings)
    st.session_state["_orig_workstreams_df"] = payload.workstreams_df.copy()
    st.session_state["_orig_tasks_df"] = payload.tasks_df.copy()

    # 6) Clear derived models + exports (forces preview/export to use new data)
    st.session_state.pop("_last_models", None)
    st.session_state.pop("last_pdf", None)
    st.session_state.pop("last_png", None)
    st.session_state.pop("last_pptx", None)


@st.cache_data(show_spinner=False)
def _read_local_file(path: str) -> bytes:
    """Read a bundled file (like the sample Excel) from disk and return raw bytes."""
    return Path(path).read_bytes()


def _mark_upload_changed() -> None:
    """Marks that the user selected a file in the uploader.

    Streamlit reruns the script frequently. Without an explicit "upload changed"
    flag, it's easy to accidentally reload/wipe user edits on any rerun.
    """

    st.session_state["_upload_changed"] = True


def main() -> None:
    """Streamlit entry point: builds the full app UI and wires together upload/edit/preview/export."""
    st.set_page_config(page_title=APP_TITLE, page_icon="🗺️", layout="wide")
    _inject_css()

    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    tabs = st.tabs(["0 Instructions", "1 Upload", "2 Edit", "3 Preview & Export"])

    # -----------------
    # Instructions tab
    # -----------------
    with tabs[0]:
        st.header("How to use this tool (non-technical step-by-step)")

        st.markdown(
            """
This app turns a simple Excel table into an executive-ready blocked swimlane roadmap and exports:

- PDF (print-ready)
- High-resolution PNG (150/300/600 DPI)
- Editable PowerPoint (PPTX)

The workflow is always:
1) Download template → 2) Fill it in Excel → 3) Upload → 4) Preview → 5) Export.
            """
        )

        st.subheader("Step 1 — Download the Excel template")
        st.markdown(
            """
Go to the Upload tab and click Download template.

The template contains 3 sheets:
- Settings (controls the time window + output preferences)
- Workstreams (the swimlanes)
- Tasks (the blocks/milestones)
            """
        )

        st.subheader("Step 2 — Fill out Settings")
        st.markdown(
            """
Open the template in Excel and edit the Settings sheet.

Required:
- `overall_start_date` and `overall_end_date`
- `chart_title`

Helpful tips:
- Keep the overall range tight (a full year is fine; multiple years can get dense).
- If your roadmap spans:
  - under ~4 months → the chart will show weeks
  - 4–12 months → the chart will show months
  - 1–2 years → the chart will show quarters
  - over 2 years → the chart will show years + quarters

Recommended defaults:
- `timezone`: `America/Chicago`
- `output_dpi`: `300`
- `page_size`: `A3` for internal planning; `A4` for quick sharing
            """
        )

        st.subheader("Step 3 — Define Workstreams (swimlanes)")
        st.markdown(
            """
In the Workstreams sheet:

Required:
- `workstream`: the swimlane name (must be unique)

Optional:
- `order`: a number to control lane ordering (smaller = higher on the chart)
- `color`: choose from the dropdown (no hex codes required)

Best practices:
- Keep workstream names short (1–4 words).
- Use `color = Auto` unless you have a strong reason to force a specific color.
            """
        )

        st.subheader("Step 4 — Add Tasks")
        st.markdown(
            """
In the Tasks sheet, each row becomes a block (or milestone).

Required columns:
- `workstream` (must exactly match a Workstreams.workstream name)
- `title`
- `start_date` and `end_date`

Optional columns:
- `description` (shown when space allows)
- `status` (`planned`, `in_progress`, `done`, `risk`)
- `type` (`block` or `milestone`)
- `hyperlink` (PPTX shapes become clickable)

Date entry tips:
- Preferred format: `YYYY-MM-DD` (example: `2026-01-15`)
- Also accepted: `M/D/YYYY` (example: `1/15/2026`)

Overlap handling (automatic):
- If tasks overlap within the same workstream, the app automatically stacks them into sub-lanes so blocks never overlap.

How status appears on the chart:
- planned: neutral stripe
- in_progress: blue stripe + dashed outline
- done: green stripe + lighter fill
- risk: red stripe + stronger red outline
            """
        )

        st.subheader("Step 5 — Upload and preview")
        st.markdown(
            """
In the Upload tab:
1) Upload your filled workbook (`.xlsx`).
2) Go to Edit to make quick fixes.
3) Go to Preview & Export to confirm the chart looks right.

Important:
- Uploading a workbook replaces all prior data in the app.
            """
        )

        st.subheader("Step 6 — Export")
        st.markdown(
            """
Use Preview & Export:
- PDF for printing / sharing
- PNG for email, docs, and crisp images
- PPTX to make last-mile edits in PowerPoint

PowerPoint tips:
- Each task is an editable shape.
- You can nudge blocks, change text, or recolor items directly in PPTX.
            """
        )

        st.subheader("Best practices for executive-ready roadmaps")
        st.markdown(
            """
- Keep task titles action-oriented (verb + object). Example: “Migrate Payroll Engine”.
- Use milestones for key dates (go-live, close, launch).
- Keep descriptions short (1 sentence). If you need detail, use `hyperlink` to a doc.
- Use `risk` sparingly — only for items that leadership should notice.
- If the chart gets too dense, switch to Monthly granularity and/or narrow the date range.
            """
        )

        st.subheader("Troubleshooting")
        st.markdown(
            """
- Unknown workstream: The task’s `workstream` doesn’t match any Workstreams row.
- Dates won’t validate: Use `YYYY-MM-DD` and make sure `end_date` is on/after `start_date`.
- Task missing from chart: It may be outside the overall date range (toggle “Include out-of-range”).
- Chart looks cramped: Reduce text, use Monthly view, or export to A3.
            """
        )

    # -----------------
    # Upload tab
    # -----------------
    with tabs[1]:
        c1, c2, c3, c4 = st.columns([1.0, 1.0, 1.2, 2.8])

        with c1:
            st.download_button(
                "Download template",
                data=_cached_template_bytes(),
                file_name="Roadmap_Input_TEMPLATE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with c2:
            try:
                sample_bytes = _read_local_file("sample_inputs/Roadmap_Sample.xlsx")
                st.download_button(
                    "Download sample",
                    data=sample_bytes,
                    file_name="Roadmap_Sample.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception:
                st.write("")

        with c3:
            try_sample = st.button("Try sample now", use_container_width=True)

        with c4:
            st.info("Upload your Excel, or click 'Try sample now' for a working demo.")

        # IMPORTANT: the uploader uses a widget callback to detect a *user* upload
        # event (vs a normal Streamlit rerun). This prevents stale rows from
        # previous uploads leaking into the editor.
        uploader_key = f"excel_uploader_{st.session_state.get('_uploader_epoch', 0)}"
        uploaded = st.file_uploader(
            "Upload an Excel workbook (.xlsx)",
            type=["xlsx"],
            key=uploader_key,
            on_change=_mark_upload_changed,
        )

        excel_bytes: Optional[bytes] = None
        source_name = ""

        if uploaded is not None:
            excel_bytes = uploaded.getvalue()
            source_name = uploaded.name
        elif try_sample:
            try:
                excel_bytes = _read_local_file("sample_inputs/Roadmap_Sample.xlsx")
                source_name = "Roadmap_Sample.xlsx"
            except Exception as e:
                st.error(f"Unable to load the bundled sample: {e}")
                excel_bytes = None

        # We consider the app "loaded" once we have a dataset in session_state,
        # even if the uploader widget is currently empty.
        has_loaded = bool(st.session_state.get("_last_upload_hash")) and (
            "workstreams_df" in st.session_state and "tasks_df" in st.session_state
        )

        if excel_bytes is not None:
            # Parse the uploaded bytes
            try:
                payload = _cached_read_excel(excel_bytes)
            except Exception as e:
                st.error(str(e))
                st.stop()

            upload_hash = hashlib.md5(excel_bytes).hexdigest()

            force_reload = bool(st.session_state.get("_force_reload_once", False))
            if force_reload:
                st.session_state["_force_reload_once"] = False

            upload_changed = bool(st.session_state.get("_upload_changed", False))
            if upload_changed:
                st.session_state["_upload_changed"] = False

            should_load = (
                force_reload
                or upload_changed
                or (st.session_state.get("_last_upload_hash") != upload_hash)
                or (not has_loaded)
            )

            if should_load:
                # Hard replace everything with the uploaded workbook content.
                _load_payload_into_state(payload, source_name=source_name, upload_hash=upload_hash)

                # Clear the uploader widget so the next upload is always a
                # "real" change event (prevents browser/file-input caching
                # quirks from keeping old bytes around).
                st.session_state["_uploader_epoch"] = int(st.session_state.get("_uploader_epoch", 0)) + 1

                # Rerun so the Edit tab grids are created with fresh widget keys.
                st.rerun()

            # If we didn't reload (same bytes and no force), still show counts
            # to confirm what file is selected.
            ws_count = int(_df_clean(payload.workstreams_df).shape[0])
            task_count = int(_df_clean(payload.tasks_df).shape[0])

            m1, m2, m3 = st.columns([1, 1, 2])
            m1.metric("Workbook", source_name)
            m2.metric("Workstreams", ws_count)
            m3.metric("Tasks", task_count)

            st.caption(f"Loaded hash: {upload_hash[:8]} • Loaded at: {st.session_state.get('_loaded_at', '')}")
            st.success("Loaded. Move to the 'Edit' tab to review and adjust data.")

        else:
            # No new file selection this run.
            if not has_loaded:
                st.warning("No workbook loaded yet. Upload an Excel file or click 'Try sample now'.")
                st.stop()

            # Show status for the already-loaded dataset (no re-parse required)
            ws_count = int(_df_clean(st.session_state.get("workstreams_df", pd.DataFrame())).shape[0])
            task_count = int(_df_clean(st.session_state.get("tasks_df", pd.DataFrame())).shape[0])
            wb_name = st.session_state.get("_active_workbook_name", "(loaded)")
            upload_hash = str(st.session_state.get("_last_upload_hash", ""))
            loaded_at = st.session_state.get("_loaded_at", "")

            m1, m2, m3 = st.columns([1, 1, 2])
            m1.metric("Workbook", wb_name)
            m2.metric("Workstreams", ws_count)
            m3.metric("Tasks", task_count)

            if upload_hash:
                st.caption(f"Loaded hash: {upload_hash[:8]} • Loaded at: {loaded_at}")
            st.success("Workbook already loaded. Upload another .xlsx anytime to replace it.")

        # Optional emergency switches (useful if the browser doesn't fire a
        # change event when re-selecting the same filename).
        with st.expander("Having trouble?", expanded=False):
            st.caption("Use these controls if your browser cached the file picker selection.")

            if st.button("Force reload workbook (replace everything)", use_container_width=True):
                st.session_state["_force_reload_once"] = True
                st.rerun()

            if st.button("Clear uploaded file selection", use_container_width=True):
                st.session_state["_uploader_epoch"] = int(st.session_state.get("_uploader_epoch", 0)) + 1
                st.rerun()

    # If we got here, we have loaded session state.
    settings_raw = dict(st.session_state.get("settings_raw", {}))
    workstreams_df = st.session_state.get("workstreams_df", pd.DataFrame(columns=WORKSTREAM_COLUMNS))
    tasks_df = st.session_state.get("tasks_df", pd.DataFrame(columns=TASK_COLUMNS))

    # -----------------
    # Sidebar settings
    # -----------------
    with st.sidebar:
        st.header("Settings")

        wb_name = st.session_state.get("_active_workbook_name")
        if wb_name:
            st.caption(f"Workbook: {wb_name}")

        # Quick actions
        a1, a2 = st.columns(2)
        with a1:
            if st.button("Reset", use_container_width=True, help="Revert edits back to the uploaded workbook"):
                st.session_state["settings_raw"] = dict(st.session_state.get("_orig_settings_raw", settings_raw))
                st.session_state["workstreams_df"] = st.session_state.get("_orig_workstreams_df", workstreams_df).copy()
                st.session_state["tasks_df"] = st.session_state.get("_orig_tasks_df", tasks_df).copy()

                # Full reset: purge all editor states and bump epoch so the
                # grids render from a clean slate.
                _purge_editor_state()
                _bump_epoch()

                # Clear derived models + exports
                st.session_state.pop("_last_models", None)
                st.session_state.pop("last_pdf", None)
                st.session_state.pop("last_png", None)
                st.session_state.pop("last_pptx", None)
                st.rerun()

        with a2:
            # Download a round-trippable workbook reflecting current edits
            try:
                xlsx_bytes = write_roadmap_excel_bytes(settings_raw, workstreams_df, tasks_df)
            except Exception:
                xlsx_bytes = b""

            st.download_button(
                "Download Excel",
                data=xlsx_bytes,
                file_name="Roadmap_Edited.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=(xlsx_bytes == b""),
            )

        st.divider()

        def _get(key: str, default: Any) -> Any:
            v = settings_raw.get(key)
            return default if _is_blank(v) else v

        settings_raw["chart_title"] = st.text_input("Chart title", value=str(_get("chart_title", "Roadmap")))
        settings_raw["chart_subtitle"] = st.text_input("Chart subtitle (optional)", value=str(_get("chart_subtitle", "")))
        settings_raw["confidentiality_label"] = st.text_input(
            "Confidentiality label (optional)", value=str(_get("confidentiality_label", "Dayforce Confidential"))
        )

        s_start = _get("overall_start_date", date.today())
        s_end = _get("overall_end_date", date.today())
        settings_raw["overall_start_date"] = st.date_input("Overall start date", value=s_start)
        settings_raw["overall_end_date"] = st.date_input("Overall end date", value=s_end)

        settings_raw["timezone"] = st.text_input("Timezone", value=str(_get("timezone", "America/Chicago")))

        settings_raw["week_start_day"] = st.selectbox(
            "Week start day",
            options=["Mon", "Sun"],
            index=0 if _get("week_start_day", "Mon") == "Mon" else 1,
        )
        settings_raw["time_granularity"] = st.selectbox(
            "Time granularity",
            options=["Weekly", "Monthly"],
            index=0 if _get("time_granularity", "Weekly") == "Weekly" else 1,
        )

        settings_raw["page_size"] = st.selectbox(
            "Page size",
            options=["A3", "A4"],
            index=0 if _get("page_size", "A3") == "A3" else 1,
        )
        settings_raw["font_family"] = st.text_input("Font family", value=str(_get("font_family", "Calibri")))

        settings_raw["show_today_line"] = st.checkbox("Show today line", value=bool(_get("show_today_line", True)))

        override_on = st.checkbox(
            "Override today line date",
            value=not _is_blank(settings_raw.get("today_line_date")),
            help="If off, the app uses today's date in the selected timezone.",
        )
        if override_on:
            settings_raw["today_line_date"] = st.date_input("Today line date", value=_get("today_line_date", date.today()))
        else:
            settings_raw["today_line_date"] = None

        settings_raw["output_dpi"] = st.selectbox(
            "Default output DPI",
            options=[150, 300, 600],
            index=[150, 300, 600].index(int(_get("output_dpi", 300))),
        )

        st.divider()
        include_out_of_range = st.toggle(
            "Include out-of-range tasks",
            value=bool(st.session_state.get("include_out_of_range", False)),
            help="If a task is completely outside the overall date range, hide it by default (recommended).",
        )
        st.session_state["include_out_of_range"] = include_out_of_range

        st.session_state["settings_raw"] = settings_raw

    # -----------------
    # Edit tab
    # -----------------
    with tabs[2]:
        st.subheader("Edit data")
        st.caption("Tip: Dates use YYYY-MM-DD (e.g., 2026-01-15). To delete rows, tick the delete checkbox and click Delete checked.")

        ws_editor_key, tasks_editor_key = _current_editor_keys()

        # Prepare editor dataframes (fixes Streamlit dtype mismatches)
        ws_for_editor = _df_clean(_ensure_columns(workstreams_df, WORKSTREAM_COLUMNS))
        ws_for_editor = _coerce_text_cols_for_editor(ws_for_editor, ["workstream", "color"])
        ws_for_editor, color_warnings_ws = _coerce_color_cols_for_editor(ws_for_editor, ["color"])

        # Add a delete checkbox column for user-friendly row removal
        if "__delete__" not in ws_for_editor.columns:
            ws_for_editor.insert(0, "__delete__", False)

        tasks_for_editor = _df_clean(_ensure_columns(tasks_df, TASK_COLUMNS))
        tasks_for_editor = _coerce_text_cols_for_editor(
            tasks_for_editor,
            [
                "id",
                "workstream",
                "title",
                "description",
                "status",
                "owner",
                "color_override",
                "type",
                "hyperlink",
            ],
        )
        tasks_for_editor, color_warnings_tasks = _coerce_color_cols_for_editor(tasks_for_editor, ["color_override"])
        tasks_for_editor = _coerce_date_cols_to_text_for_editor(tasks_for_editor, ["start_date", "end_date"])

        # Surface any color coercion warnings once (non-blocking).
        for w in (color_warnings_ws + color_warnings_tasks):
            st.info(w)

        # Add a delete checkbox column for user-friendly row removal
        if "__delete__" not in tasks_for_editor.columns:
            tasks_for_editor.insert(0, "__delete__", False)

        delete_ws_clicked = False
        delete_tasks_clicked = False

        left, right = st.columns([1, 2])
        with left:
            ws_h1, ws_h2, ws_h3 = st.columns([0.58, 0.21, 0.21])
            with ws_h1:
                st.markdown("Workstreams")
            with ws_h2:
                if st.button(
                    "Clear",
                    key=f"clear_workstreams_{st.session_state.get('_data_epoch', 0)}",
                    help="Clear all workstreams in the grid.",
                    use_container_width=True,
                ):
                    st.session_state["workstreams_df"] = pd.DataFrame(columns=WORKSTREAM_COLUMNS)
                    # Also clear tasks because they reference workstreams.
                    st.session_state["tasks_df"] = pd.DataFrame(columns=TASK_COLUMNS)
                    _purge_editor_state()
                    _bump_epoch()
                    st.rerun()
            with ws_h3:
                delete_ws_clicked = st.button(
                    "Delete checked",
                    key=f"delete_ws_{st.session_state.get('_data_epoch', 0)}",
                    help="Tick rows using the delete checkbox, then click to remove them.",
                    use_container_width=True,
                )

            edited_ws = st.data_editor(
                ws_for_editor,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "__delete__": st.column_config.CheckboxColumn("delete", help="Tick then click 'Delete checked'."),
                    "workstream": st.column_config.TextColumn("workstream"),
                    "order": st.column_config.NumberColumn("order", help="Optional: lower numbers appear first."),
                    "color": st.column_config.SelectboxColumn(
                        "color",
                        options=COLOR_DROPDOWN_VALUES,
                        help="Optional. Choose a named color. Auto picks an accessible default.",
                    ),
                },
                key=ws_editor_key,
            )
            edited_ws_clean = edited_ws.drop(columns=["__delete__"], errors="ignore")
            st.session_state["workstreams_df"] = edited_ws_clean

        with right:
            t_h1, t_h2, t_h3 = st.columns([0.66, 0.17, 0.17])
            with t_h1:
                st.markdown("Tasks")
            with t_h2:
                if st.button(
                    "Clear",
                    key=f"clear_tasks_{st.session_state.get('_data_epoch', 0)}",
                    help="Clear all tasks in the grid.",
                    use_container_width=True,
                ):
                    st.session_state["tasks_df"] = pd.DataFrame(columns=TASK_COLUMNS)
                    _purge_editor_state()
                    _bump_epoch()
                    st.rerun()
            with t_h3:
                delete_tasks_clicked = st.button(
                    "Delete checked",
                    key=f"delete_tasks_{st.session_state.get('_data_epoch', 0)}",
                    help="Tick rows using the delete checkbox, then click to remove them.",
                    use_container_width=True,
                )

            edited_tasks = st.data_editor(
                tasks_for_editor,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "__delete__": st.column_config.CheckboxColumn("delete", help="Tick then click 'Delete checked'."),
                    "id": st.column_config.TextColumn("id", help="Leave blank to auto-generate a stable ID."),
                    "workstream": st.column_config.TextColumn("workstream", help="Must match a Workstream."),
                    "title": st.column_config.TextColumn("title"),
                    "description": st.column_config.TextColumn("description"),
                    "start_date": st.column_config.TextColumn(
                        "start_date",
                        help="Enter a date like 2026-01-15 (YYYY-MM-DD). You can also paste 1/15/2026.",
                    ),
                    "end_date": st.column_config.TextColumn(
                        "end_date",
                        help="Enter a date like 2026-02-01 (YYYY-MM-DD). You can also paste 2/1/2026.",
                    ),
                    "status": st.column_config.SelectboxColumn(
                        "status",
                        options=["", "planned", "in_progress", "done", "risk"],
                        help="Optional. Used for subtle border styling + legend.",
                    ),
                    "owner": st.column_config.TextColumn("owner"),
                    "color_override": st.column_config.SelectboxColumn(
                        "color_override",
                        options=COLOR_DROPDOWN_VALUES,
                        help="Optional. Choose a named color. Auto inherits the workstream color.",
                    ),
                    "type": st.column_config.SelectboxColumn(
                        "type",
                        options=["", "block", "milestone"],
                        help="milestone renders as a diamond at start_date",
                    ),
                    "hyperlink": st.column_config.TextColumn("hyperlink"),
                },
                key=tasks_editor_key,
            )
            edited_tasks_clean = edited_tasks.drop(columns=["__delete__"], errors="ignore")
            st.session_state["tasks_df"] = edited_tasks_clean

# Apply deletions (checkbox-based row removal)
        if delete_ws_clicked:
            ws_del_mask = edited_ws.get("__delete__", pd.Series([False] * len(edited_ws))).apply(lambda v: v is True)
            delete_ws_names = set(
                edited_ws.loc[ws_del_mask, "workstream"].astype(str).apply(lambda s: s.strip()).tolist()
            )

            new_ws_df = (
                edited_ws.loc[~ws_del_mask]
                .drop(columns=["__delete__"], errors="ignore")
                .reset_index(drop=True)
            )

            # When workstreams are deleted, also delete any tasks that reference them.
            new_tasks_df = edited_tasks.drop(columns=["__delete__"], errors="ignore")
            if delete_ws_names and "workstream" in new_tasks_df.columns:
                new_tasks_df = new_tasks_df[
                    ~new_tasks_df["workstream"].astype(str).apply(lambda s: s.strip()).isin(delete_ws_names)
                ].reset_index(drop=True)

            st.session_state["workstreams_df"] = new_ws_df
            st.session_state["tasks_df"] = new_tasks_df
            _purge_editor_state()
            _bump_epoch()
            st.rerun()

        if delete_tasks_clicked:
            t_del_mask = edited_tasks.get("__delete__", pd.Series([False] * len(edited_tasks))).apply(lambda v: v is True)
            new_tasks_df = (
                edited_tasks.loc[~t_del_mask]
                .drop(columns=["__delete__"], errors="ignore")
                .reset_index(drop=True)
            )
            st.session_state["tasks_df"] = new_tasks_df
            _purge_editor_state()
            _bump_epoch()
            st.rerun()

        st.divider()

        settings, workstreams, tasks, errors, warnings = _build_models(settings_raw, edited_ws_clean, edited_tasks_clean)

        # Show validation summary
        if errors:
            st.error(f"{len(errors)} issue(s) to fix before export")
            for e in errors:
                st.write(f"- {e}")
        else:
            st.success("No blocking issues found.")

        if warnings:
            with st.expander(f"Warnings ({len(warnings)})", expanded=False):
                seen = set()
                for w in warnings:
                    if w in seen:
                        continue
                    seen.add(w)
                    st.write(f"- {w}")

        st.session_state["_last_models"] = {
            "settings": settings.model_dump() if settings else None,
            "workstreams": [w.model_dump() for w in workstreams],
            "tasks": [t.model_dump() for t in tasks],
            "errors": errors,
            "warnings": warnings,
        }

    # -----------------
    # Preview & Export tab
    # -----------------
    with tabs[3]:
        st.subheader("Preview & export")

        models = st.session_state.get("_last_models")
        if not models:
            st.info("Open the 'Edit' tab first to validate your data.")
            st.stop()

        errors = models.get("errors") or []
        settings_dump = models.get("settings")
        ws_dump = models.get("workstreams") or []
        tasks_dump = models.get("tasks") or []

        if errors or settings_dump is None:
            st.info("Fix issues in the 'Edit' tab to enable preview and exports.")
            st.stop()

        include_out_of_range = bool(st.session_state.get("include_out_of_range", False))

        # Preview
        try:
            preview_bytes = _cached_preview(
                settings_dump,
                ws_dump,
                tasks_dump,
                include_out_of_range=include_out_of_range,
                dpi=140,
            )
            st.image(preview_bytes, caption="Live preview (lower resolution)", use_container_width=True)
        except Exception as e:
            st.error(f"Preview failed: {e}")
            st.stop()

        st.divider()

        # Export controls
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("Generate PDF", use_container_width=True):
                try:
                    pdf_bytes = _cached_export_pdf(
                        settings_dump,
                        ws_dump,
                        tasks_dump,
                        include_out_of_range=include_out_of_range,
                    )
                    st.session_state["last_pdf"] = pdf_bytes
                except Exception as e:
                    st.error(f"PDF export failed: {e}")

            st.download_button(
                "Download PDF",
                data=st.session_state.get("last_pdf", b""),
                file_name="roadmap.pdf",
                mime="application/pdf",
                use_container_width=True,
                disabled=("last_pdf" not in st.session_state),
            )

        with col2:
            png_dpi = st.selectbox("PNG DPI", options=[150, 300, 600], index=1)

            if st.button("Generate PNG", use_container_width=True):
                try:
                    png_bytes = _cached_export_png(
                        settings_dump,
                        ws_dump,
                        tasks_dump,
                        include_out_of_range=include_out_of_range,
                        dpi=int(png_dpi),
                    )
                    st.session_state["last_png"] = png_bytes
                except Exception as e:
                    st.error(f"PNG export failed: {e}")

            st.download_button(
                "Download PNG",
                data=st.session_state.get("last_png", b""),
                file_name=f"roadmap_{png_dpi}dpi.png",
                mime="image/png",
                use_container_width=True,
                disabled=("last_png" not in st.session_state),
            )

        with col3:
            if st.button("Generate PPTX", use_container_width=True):
                try:
                    pptx_bytes = _cached_export_pptx(
                        settings_dump,
                        ws_dump,
                        tasks_dump,
                        include_out_of_range=include_out_of_range,
                    )
                    st.session_state["last_pptx"] = pptx_bytes
                except Exception as e:
                    st.error(f"PPTX export failed: {e}")

            st.download_button(
                "Download PPTX",
                data=st.session_state.get("last_pptx", b""),
                file_name="roadmap_editable.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                disabled=("last_pptx" not in st.session_state),
            )

        st.caption("Tip: for print decks, use PDF. For slides or emails, a 300 DPI PNG is usually perfect. PPTX gives you fully editable shapes.")


if __name__ == "__main__":
    main()
