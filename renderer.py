from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

import matplotlib
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon, Rectangle
from zoneinfo import ZoneInfo

from date_utils import block_span_inclusive, date_to_x

from roadmap_models import Settings, Task, Workstream


DEFAULT_PALETTE = [
    "#1F77B4",  # blue
    "#FF7F0E",  # orange
    "#2CA02C",  # green
    "#D62728",  # red
    "#9467BD",  # purple
    "#8C564B",  # brown
    "#E377C2",  # pink
    "#7F7F7F",  # gray
    "#BCBD22",  # olive
    "#17BECF",  # cyan
    "#AEC7E8",  # sky blue (light)
    "#FFBB78",  # peach (light)
]


@dataclass(frozen=True)
class LayoutRow:
    workstream: str
    sublane: int
    y0: float
    y1: float


@dataclass(frozen=True)
class WorkstreamBand:
    workstream: str
    y0: float
    y1: float
    color: str


def _font_family_available(family: str) -> bool:
    family = (family or "").strip()
    if not family:
        return False
    # Matplotlib stores font names; check case-insensitively.
    fam_lower = family.lower()
    from matplotlib import font_manager as fm
    for f in fm.fontManager.ttflist:
        if f.name.lower() == fam_lower:
            return True
    return False


def resolve_font_family(preferred: str) -> str:
    """
    Returns a font family name that matplotlib can actually render.
    Priority:
      1) preferred, if available
      2) Arial, if available
      3) DejaVu Sans (matplotlib default)
    """
    preferred = (preferred or "").strip()
    if preferred and _font_family_available(preferred):
        return preferred
    if _font_family_available("Arial"):
        return "Arial"
    return "DejaVu Sans"


def _start_of_week(d: date, week_start_day: str) -> date:
    # week_start_day is "Mon" or "Sun"
    weekday = d.weekday()  # Mon=0..Sun=6
    if week_start_day == "Mon":
        delta = weekday
    else:  # Sun
        # Convert to Sun=0..Sat=6
        sun_based = (weekday + 1) % 7
        delta = sun_based
    return d - timedelta(days=delta)


def _iter_week_starts(start: date, end: date, week_start_day: str) -> List[date]:
    cur = _start_of_week(start, week_start_day)
    # Ensure first boundary isn't before the start too far; keep it, it's ok for grid.
    out: List[date] = []
    while cur <= end:
        out.append(cur)
        cur += timedelta(days=7)
    return out


def _iter_month_starts(start: date, end: date) -> List[date]:
    cur = date(start.year, start.month, 1)
    out: List[date] = []
    while cur <= end:
        out.append(cur)
        # next month
        if cur.month == 12:
            cur = date(cur.year + 1, 1, 1)
        else:
            cur = date(cur.year, cur.month + 1, 1)
    return out




# ---------------------------------------------------------------------------
# Timeline helpers (auto granularity + labeled header band)
# ---------------------------------------------------------------------------

def _months_span_inclusive(start: date, end: date) -> int:
    """Count distinct calendar months touched by [start, end]."""
    return (end.year - start.year) * 12 + (end.month - start.month) + 1


def choose_timeline_mode(start: date, end: date) -> str:
    """
    Auto-select timeline labeling mode based on overall range.

    Rules (per product spec):
      - < 4 months: weeks + months
      - 4-12 months: months
      - > 12 months and <= 24 months: quarters
      - > 24 months: quarters + years
    """
    months = _months_span_inclusive(start, end)
    if months < 4:
        return "weeks"
    if months <= 12:
        return "months"
    if months <= 24:
        return "quarters"
    return "quarters_years"


def _quarter_start(d: date) -> date:
    qm = ((d.month - 1) // 3) * 3 + 1
    return date(d.year, qm, 1)


def _iter_quarter_starts(start: date, end: date) -> List[date]:
    cur = _quarter_start(start)
    out: List[date] = []
    while cur <= end:
        out.append(cur)
        # next quarter (+3 months)
        m = cur.month + 3
        y = cur.year
        if m > 12:
            m -= 12
            y += 1
        cur = date(y, m, 1)
    return out


def _iter_year_starts(start: date, end: date) -> List[date]:
    cur = date(start.year, 1, 1)
    out: List[date] = []
    while cur <= end:
        out.append(cur)
        cur = date(cur.year + 1, 1, 1)
    return out


def _segments_from_boundaries(
    boundaries: List[date],
    start: date,
    end_exclusive: date,
) -> List[Tuple[date, date]]:
    """Build contiguous segments between boundaries, clamped to [start, end_exclusive)."""
    out: List[Tuple[date, date]] = []
    for i, b in enumerate(boundaries):
        seg_start = max(b, start)
        next_b = boundaries[i + 1] if i + 1 < len(boundaries) else end_exclusive
        seg_end = min(next_b, end_exclusive)
        if seg_end > seg_start:
            out.append((seg_start, seg_end))
    return out


def _build_month_segments(start: date, end: date) -> List[Tuple[date, date, str]]:
    end_excl = end + timedelta(days=1)
    boundaries = _iter_month_starts(start, end)
    segs = _segments_from_boundaries(boundaries, start, end_excl)

    out: List[Tuple[date, date, str]] = []
    prev_year: Optional[int] = None
    for idx, (s, e) in enumerate(segs):
        include_year = idx == 0 or (prev_year is not None and s.year != prev_year) or s.month == 1
        label = s.strftime("%b %Y") if include_year else s.strftime("%b")
        out.append((s, e, label))
        prev_year = s.year
    return out


def _build_week_segments(start: date, end: date, week_start_day: str) -> List[Tuple[date, date, str]]:
    end_excl = end + timedelta(days=1)
    boundaries = _iter_week_starts(start, end, week_start_day)
    segs = _segments_from_boundaries(boundaries, start, end_excl)

    out: List[Tuple[date, date, str]] = []
    prev_m_y: Optional[Tuple[int, int]] = None
    for idx, (s, e) in enumerate(segs):
        my = (s.month, s.year)
        include_month = idx == 0 or (prev_m_y is not None and my != prev_m_y)
        label = s.strftime("%d %b") if include_month else s.strftime("%d")
        out.append((s, e, label))
        prev_m_y = my
    return out


def _build_quarter_segments(start: date, end: date, *, include_year: bool) -> List[Tuple[date, date, str]]:
    end_excl = end + timedelta(days=1)
    boundaries = _iter_quarter_starts(start, end)
    segs = _segments_from_boundaries(boundaries, start, end_excl)

    out: List[Tuple[date, date, str]] = []
    prev_year: Optional[int] = None
    for idx, (s, e) in enumerate(segs):
        q = ((s.month - 1) // 3) + 1
        if include_year:
            include = idx == 0 or q == 1 or (prev_year is not None and s.year != prev_year)
            label = f"Q{q} {s.year}" if include else f"Q{q}"
        else:
            label = f"Q{q}"
        out.append((s, e, label))
        prev_year = s.year
    return out


def _build_year_segments(start: date, end: date) -> List[Tuple[date, date, str]]:
    end_excl = end + timedelta(days=1)
    boundaries = _iter_year_starts(start, end)
    segs = _segments_from_boundaries(boundaries, start, end_excl)

    out: List[Tuple[date, date, str]] = []
    for s, e in segs:
        out.append((s, e, str(s.year)))
    return out


def _draw_timeline_rows(
    ax,
    *,
    x0: float,
    x1: float,
    y_top: float,
    row_h: float,
    rows: List[Tuple[str, List[Tuple[date, date, str]]]],
    preview: bool,
) -> None:
    """Draw a compact, executive-friendly timeline band above the swimlanes."""

    alt_a = "#FFFFFF"
    alt_b = "#F7F7F7"
    border = "#DADADA"
    text = "#333333"

    fs_map = {
        "months": 9,
        "weeks": 8,
        "quarters": 10,
        "years": 11,
    }
    if preview:
        fs_map = {k: max(v - 1, 7) for k, v in fs_map.items()}

    # Outer frame
    ax.hlines(y_top, x0, x1, colors=border, linewidth=0.9, zorder=2)
    ax.hlines(0.0, x0, x1, colors=border, linewidth=0.9, zorder=2)

    for r_idx, (kind, segs) in enumerate(rows):
        ry0 = y_top + r_idx * row_h
        # row separator
        ax.hlines(ry0, x0, x1, colors=border, linewidth=0.8, zorder=2)

        for i, (ds, de, label) in enumerate(segs):
            xs = date_to_x(ds)
            xe = date_to_x(de)
            face = alt_a if (i % 2 == 0) else alt_b

            ax.add_patch(
                Rectangle(
                    (xs, ry0),
                    xe - xs,
                    row_h,
                    facecolor=face,
                    edgecolor=border,
                    linewidth=0.8,
                    zorder=2,
                )
            )
            ax.text(
                (xs + xe) / 2.0,
                ry0 + row_h * 0.52,
                label,
                ha="center",
                va="center",
                fontsize=fs_map.get(kind, 9),
                color=text,
                zorder=3,
            )

def compute_bands_and_rows(
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    group_gap: float = 0.35,
    min_rows_per_workstream: int = 1,
) -> Tuple[List[WorkstreamBand], Dict[Tuple[str, int], LayoutRow], float]:
    """
    Returns:
      - bands: one per workstream group
      - row_map: (workstream, sublane) -> LayoutRow (y0/y1)
      - total_height in y units
    """
    # Determine max sublane per workstream
    lane_counts: Dict[str, int] = {}
    for t in tasks:
        if t.sublane is None:
            continue
        lane_counts[t.workstream] = max(lane_counts.get(t.workstream, 0), int(t.sublane) + 1)

    # Ensure at least 1 row per declared workstream
    for ws in workstreams:
        lane_counts.setdefault(ws.workstream, min_rows_per_workstream)

    y = 0.0
    bands: List[WorkstreamBand] = []
    row_map: Dict[Tuple[str, int], LayoutRow] = {}

    for ws in workstreams:
        n = lane_counts.get(ws.workstream, min_rows_per_workstream)
        band_y0 = y
        for lane in range(n):
            row_y0 = y
            row_y1 = y + 1.0
            row_map[(ws.workstream, lane)] = LayoutRow(workstream=ws.workstream, sublane=lane, y0=row_y0, y1=row_y1)
            y = row_y1
        band_y1 = y
        bands.append(WorkstreamBand(workstream=ws.workstream, y0=band_y0, y1=band_y1, color=(ws.color or "")))
        y += group_gap

    total_height = max(y - group_gap, 0.0)  # remove trailing gap
    # Fix last band's y1 if we removed trailing gap
    if bands:
        last = bands[-1]
        bands[-1] = WorkstreamBand(workstream=last.workstream, y0=last.y0, y1=total_height, color=last.color)

    return bands, row_map, total_height


def _pick_workstream_colors(workstreams: List[Workstream]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    palette = DEFAULT_PALETTE
    i = 0
    for ws in workstreams:
        if ws.color:
            out[ws.workstream] = ws.color
        else:
            out[ws.workstream] = palette[i % len(palette)]
            i += 1
    return out


def render_roadmap(
    settings: Settings,
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    include_out_of_range: bool = False,
    preview: bool = False,
    preview_dpi: int = 150,
) -> Tuple[plt.Figure, Dict[str, List[str]], Dict[str, List[str]]]:
    """
    Builds the matplotlib figure. Returns (fig, warnings, hidden_task_ids).

    warnings: dict categories -> messages
    hidden_task_ids: dict categories -> list of ids (e.g., out_of_range_hidden)
    """
    font_family = resolve_font_family(settings.font_family)
    matplotlib.rcParams["font.family"] = font_family

    # Normalize workstream colors
    ws_color_map = _pick_workstream_colors(workstreams)
    workstreams_norm = [ws.model_copy(update={"color": ws_color_map[ws.workstream]}) for ws in workstreams]

    # Filter/clamp tasks by overall range
    overall_start = settings.overall_start_date
    overall_end = settings.overall_end_date

    warnings: Dict[str, List[str]] = {"clamped": [], "out_of_range": []}
    hidden: Dict[str, List[str]] = {"out_of_range_hidden": []}

    visible_tasks: List[Task] = []
    for t in tasks:
        if t.end_date < overall_start or t.start_date > overall_end:
            warnings["out_of_range"].append(f"{t.id}: '{t.title}' is outside the overall date range.")
            if include_out_of_range:
                visible_tasks.append(t)
            else:
                hidden["out_of_range_hidden"].append(t.id)
            continue

        # Clamp for rendering, but keep original model fields for export logic elsewhere.
        clamped_start = max(t.start_date, overall_start)
        clamped_end = min(t.end_date, overall_end)
        if clamped_start != t.start_date or clamped_end != t.end_date:
            warnings["clamped"].append(f"{t.id}: '{t.title}' is partially outside range; clamped in the chart.")
            visible_tasks.append(t.model_copy(update={"start_date": clamped_start, "end_date": clamped_end}))
        else:
            visible_tasks.append(t)

    # Page size (landscape)
    if settings.page_size == "A4":
        fig_w, fig_h = 11.69, 8.27
    else:  # A3
        fig_w, fig_h = 16.54, 11.69

    fig = plt.figure(figsize=(fig_w, fig_h), dpi=preview_dpi if preview else settings.output_dpi)
    # Grids: header row + chart row; chart has labels+main columns
    gs = fig.add_gridspec(
        nrows=2,
        ncols=2,
        height_ratios=[0.13, 0.87],
        width_ratios=[0.23, 0.77],
        wspace=0.02,
        hspace=0.05,
    )
    ax_header = fig.add_subplot(gs[0, :])
    ax_labels = fig.add_subplot(gs[1, 0])
    ax_main = fig.add_subplot(gs[1, 1], sharey=ax_labels)

    ax_header.axis("off")

    # Header text
    ax_header.text(
        0.0, 0.78, settings.chart_title,
        fontsize=18 if not preview else 16,
        fontweight="bold",
        ha="left",
        va="center",
        transform=ax_header.transAxes,
    )
    if settings.chart_subtitle:
        ax_header.text(
            0.0, 0.42, settings.chart_subtitle,
            fontsize=11 if not preview else 10,
            ha="left",
            va="center",
            transform=ax_header.transAxes,
        )
    if settings.confidentiality_label:
        ax_header.text(
            1.0, 0.78, settings.confidentiality_label,
            fontsize=10 if not preview else 9,
            ha="right",
            va="center",
            transform=ax_header.transAxes,
        )

    # Date range (helps executives orient immediately)
    date_range_text = f"{overall_start.strftime('%d %b %Y')} – {overall_end.strftime('%d %b %Y')}"
    ax_header.text(
        0.0,
        0.12,
        date_range_text,
        fontsize=10 if not preview else 9,
        ha="left",
        va="center",
        color="#333333",
        transform=ax_header.transAxes,
    )

    # Subtle divider under header
    ax_header.hlines(0.0, 0.0, 1.0, transform=ax_header.transAxes, colors="#E6E6E6", linewidth=1.0)

    # Timeline axis formatting
    x0 = date_to_x(overall_start)
    x1 = date_to_x(overall_end + timedelta(days=1))
    ax_main.set_xlim(x0, x1)

    # Rows/bands
    bands, row_map, total_height = compute_bands_and_rows(workstreams_norm, visible_tasks)

    # Timeline header band height (adds a dedicated header area above swimlanes)
    timeline_mode = choose_timeline_mode(overall_start, overall_end)
    timeline_rows = 2 if timeline_mode in ("weeks", "quarters_years") else 1
    timeline_row_h = 0.65
    timeline_height = timeline_rows * timeline_row_h

    ax_main.set_ylim(total_height, -timeline_height)
    ax_labels.set_ylim(total_height, -timeline_height)

    # Hide axes frames/ticks
    for ax in (ax_main, ax_labels):
        ax.spines[:].set_visible(False)
        ax.tick_params(left=False, labelleft=False, bottom=False, labelbottom=False)
        ax.set_yticks([])

    ax_labels.set_xlim(0, 1)
    ax_labels.set_xticks([])
    ax_labels.set_facecolor("#F6F8FB")
    ax_main.set_facecolor("white")

    # Vertical divider between labels and chart
    ax_labels.vlines(1.0, -timeline_height, total_height, colors="#E0E0E0", linewidth=1.0)

    # ---------------------------------------------------------------------
    # Timeline header band (auto-selected based on overall range)
    #   < 4 months  -> weeks + months
    #   4-12 months -> months
    #   13-24 months-> quarters
    #   > 24 months -> years + quarters
    # ---------------------------------------------------------------------

    # Background fills for the timeline band area
    ax_labels.add_patch(
        Rectangle((0.0, -timeline_height), 1.0, timeline_height, facecolor="#F6F8FB", edgecolor="none", zorder=0)
    )
    ax_main.add_patch(
        Rectangle((x0, -timeline_height), x1 - x0, timeline_height, facecolor="#FFFFFF", edgecolor="none", zorder=0)
    )

    # Separator line between timeline and swimlanes
    ax_main.hlines(0.0, x0, x1, colors="#D0D0D0", linewidth=1.2, zorder=2)
    ax_labels.hlines(0.0, 0.0, 1.0, colors="#D0D0D0", linewidth=1.2, zorder=2)

    # We draw our own timeline labels, so hide x-axis ticks/labels.
    ax_main.set_xticks([])
    ax_main.tick_params(axis="x", top=False, bottom=False, labeltop=False, labelbottom=False)

    grid_color = "#E6E6E6"
    major_color = "#D0D0D0"
    year_color = "#C8C8C8"

    if timeline_mode == "weeks":
        # Week gridlines + stronger month boundaries
        week_starts = _iter_week_starts(overall_start, overall_end, settings.week_start_day)
        month_starts = _iter_month_starts(overall_start, overall_end)

        for d in week_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=grid_color, linewidth=0.8, zorder=0)
        for d in month_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=major_color, linewidth=1.2, zorder=1)

        rows = [
            ("months", _build_month_segments(overall_start, overall_end)),
            ("weeks", _build_week_segments(overall_start, overall_end, settings.week_start_day)),
        ]
        _draw_timeline_rows(
            ax_main,
            x0=x0,
            x1=x1,
            y_top=-timeline_height,
            row_h=timeline_row_h,
            rows=rows,
            preview=preview,
        )

    elif timeline_mode == "months":
        month_starts = _iter_month_starts(overall_start, overall_end)
        year_starts = _iter_year_starts(overall_start, overall_end)

        for d in month_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=grid_color, linewidth=0.9, zorder=0)
        for d in year_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=year_color, linewidth=1.4, zorder=1)

        rows = [("months", _build_month_segments(overall_start, overall_end))]
        _draw_timeline_rows(
            ax_main,
            x0=x0,
            x1=x1,
            y_top=-timeline_height,
            row_h=timeline_row_h,
            rows=rows,
            preview=preview,
        )

    elif timeline_mode == "quarters":
        quarter_starts = _iter_quarter_starts(overall_start, overall_end)
        year_starts = _iter_year_starts(overall_start, overall_end)

        for d in quarter_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=grid_color, linewidth=1.0, zorder=0)
        for d in year_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=year_color, linewidth=1.6, zorder=1)

        rows = [("quarters", _build_quarter_segments(overall_start, overall_end, include_year=True))]
        _draw_timeline_rows(
            ax_main,
            x0=x0,
            x1=x1,
            y_top=-timeline_height,
            row_h=timeline_row_h,
            rows=rows,
            preview=preview,
        )

    else:  # quarters_years
        quarter_starts = _iter_quarter_starts(overall_start, overall_end)
        year_starts = _iter_year_starts(overall_start, overall_end)

        for d in quarter_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=grid_color, linewidth=1.0, zorder=0)
        for d in year_starts:
            ax_main.axvline(x=date_to_x(d), ymin=0, ymax=1, color=year_color, linewidth=1.6, zorder=1)

        rows = [
            ("years", _build_year_segments(overall_start, overall_end)),
            ("quarters", _build_quarter_segments(overall_start, overall_end, include_year=False)),
        ]
        _draw_timeline_rows(
            ax_main,
            x0=x0,
            x1=x1,
            y_top=-timeline_height,
            row_h=timeline_row_h,
            rows=rows,
            preview=preview,
        )



# Workstream bands and row separators
    sep_color = "#D0D0D0"
    alt_fill = "#FAFAFA"
    for i, band in enumerate(bands):
        # Light alternating background on main axis
        if i % 2 == 0:
            ax_main.add_patch(Rectangle((x0, band.y0), x1 - x0, band.y1 - band.y0, facecolor=alt_fill, edgecolor="none", zorder=0))
        # separators
        ax_main.hlines(band.y0, x0, x1, colors=sep_color, linewidth=1.0, zorder=1)
        ax_main.hlines(band.y1, x0, x1, colors=sep_color, linewidth=1.0, zorder=1)

        # Label axis: workstream name centered in band + color accent
        ws_color = ws_color_map.get(band.workstream, "#1F77B4")
        ax_labels.add_patch(Rectangle((0.02, band.y0), 0.025, band.y1 - band.y0, facecolor=ws_color, edgecolor="none", zorder=2))
        ax_labels.text(
            0.06, (band.y0 + band.y1) / 2.0,
            band.workstream,
            ha="left",
            va="center",
            fontsize=10 if not preview else 9,
            fontweight="bold",
            color="#222222",
        )

    # Row lines (sublanes)
    for (ws, lane), row in row_map.items():
        ax_main.hlines(row.y0, x0, x1, colors="#EFEFEF", linewidth=0.6, zorder=1)

    # Today line
    if settings.show_today_line:
        tz = ZoneInfo(settings.timezone)
        today = settings.today_line_date
        if today is None:
            now = datetime.now(tz=tz)
            today = now.date()
        if overall_start <= today <= overall_end:
            ax_main.axvline(x=date_to_x(today), color="#111111", linewidth=1.2, linestyle="--", zorder=3)
            ax_main.text(
                date_to_x(today) + (x1 - x0) * 0.002,
                -timeline_height + 0.06,
                "Today",
                fontsize=8 if not preview else 7,
                va="bottom",
                ha="left",
                color="#111111",
                zorder=4,
            )

    # Draw tasks
    row_padding_y = 0.12
    block_height = 1.0 - 2 * row_padding_y
    text_color_default = "#1A1A1A"
    border_default = "#3A3A3A"

    # Status styling (high-contrast but still exec-friendly)
    #
    # Problem we are solving: border thickness + alpha alone can be hard to
    # distinguish at a glance. We add a thin status stripe on the left edge of
    # every block plus a clearer border style.
    #
    # The status stripe width is defined in pixels (converted to data units) so
    # it stays visible regardless of the date range.

    def _lighten_hex(hex_color: str, amount: float) -> str:
        """Blend a color with white. amount in [0, 1]."""
        import matplotlib.colors as mcolors

        try:
            r, g, b = mcolors.to_rgb(hex_color)
        except Exception:
            return hex_color
        r = r + (1.0 - r) * amount
        g = g + (1.0 - g) * amount
        b = b + (1.0 - b) * amount
        return mcolors.to_hex((r, g, b))

    def _px_to_data_dx(ax, x: float, y: float, px: float) -> float:
        p = ax.transData.transform((x, y))
        x2, _ = ax.transData.inverted().transform((p[0] + px, p[1]))
        return float(x2 - x)

    STATUS_STYLE = {
        # Stripe uses status-specific color; face keeps workstream color.
        "planned": {
            "stripe": "#6B7280",  # neutral gray
            "edge": border_default,
            "lw": 1.0,
            "ls": "solid",
            "lighten": 0.0,
        },
        "in_progress": {
            "stripe": "#2563EB",  # blue
            "edge": "#2563EB",
            "lw": 1.8,
            "ls": (0, (4, 2)),  # dashed
            "lighten": 0.0,
        },
        "done": {
            "stripe": "#16A34A",  # green
            "edge": "#6B7280",
            "lw": 1.0,
            "ls": "solid",
            "lighten": 0.70,
        },
        "risk": {
            "stripe": "#DC2626",  # red
            "edge": "#DC2626",
            "lw": 2.2,
            "ls": "solid",
            "lighten": 0.0,
        },
    }

    def fit_text(title: str, desc: Optional[str], width_px: float, height_px: float) -> Tuple[str, int]:
        """
        Returns (text, fontsize).
        Simple heuristic: choose largest fontsize that avoids truncating the title.
        """
        base = 9 if not preview else 8
        min_fs = 7 if not preview else 6

        title = (title or "").strip()
        desc = (desc or "").strip() if desc else ""

        best_text = title
        best_fs = min_fs

        for fs in range(base, min_fs - 1, -1):
            # Approximate characters per line; 0.55 is a reasonable average glyph width factor.
            max_cpl = max(int(width_px / (fs * 0.55)), 8)
            max_lines = max(int(height_px / (fs * 1.25)), 1)

            title_max_lines = min(2, max_lines) if desc else max_lines
            title_lines = _wrap_and_truncate(title, max_cpl, title_max_lines)

            remaining_lines = max_lines - len(title_lines)
            desc_lines: List[str] = []
            if desc and remaining_lines > 0:
                desc_lines = _wrap_and_truncate(desc, max_cpl, remaining_lines)

            composed = "\n".join(title_lines + desc_lines)

            # Prefer the first fontsize where title isn't truncated
            if not _is_truncated(title_lines, title, max_cpl):
                best_text = composed
                best_fs = fs
                break

            # Keep best fallback (smallest truncation)
            best_text = composed
            best_fs = fs

        return best_text, best_fs

    def _wrap_and_truncate(text: str, width: int, max_lines: int) -> List[str]:
        import textwrap

        if max_lines <= 0:
            return []
        lines = textwrap.wrap(text, width=width) or [""]
        if len(lines) <= max_lines:
            return lines
        kept = lines[:max_lines]
        # Truncate last line with ellipsis
        kept[-1] = _ellipsis(kept[-1], width)
        return kept

    def _ellipsis(line: str, width: int) -> str:
        if width <= 1:
            return "…"
        line = line.rstrip()
        if len(line) <= width:
            return line
        if width <= 2:
            return line[: width - 1] + "…"
        return line[: width - 1].rstrip() + "…"

    def _is_truncated(lines: List[str], original: str, width: int) -> bool:
        # Rough detection: if last line ends with ellipsis and original doesn't already
        if not lines:
            return False
        return lines[-1].endswith("…") and not original.endswith("…")

    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()

    for t in visible_tasks:
        lane = int(t.sublane or 0)
        row = row_map.get((t.workstream, lane))
        if row is None:
            # In case tasks reference a workstream missing in declarations; render at bottom.
            continue

        y0 = row.y0 + row_padding_y
        y1 = y0 + block_height

        ws_color = ws_color_map.get(t.workstream, "#1F77B4")
        face = (t.color_override or ws_color)
        status = (t.status or "planned").lower() if t.status else "planned"
        style = STATUS_STYLE.get(status, STATUS_STYLE["planned"])
        edge = style["edge"]
        lw = float(style["lw"])
        ls = style["ls"]
        stripe_color = style["stripe"]

        face2 = face
        if float(style.get("lighten", 0.0)) > 0:
            face2 = _lighten_hex(face, float(style["lighten"]))
        alpha = 0.98

        if t.type == "milestone":
            # Diamond centered on start_date
            x = date_to_x(t.start_date)
            # Size in data units: ~0.6 day wide
            half_w = 0.35
            cy = (y0 + y1) / 2.0
            half_h = (y1 - y0) * 0.45
            poly = Polygon(
                [(x, cy + half_h), (x + half_w, cy), (x, cy - half_h), (x - half_w, cy)],
                closed=True,
                facecolor=face2,
                edgecolor=edge,
                linewidth=lw,
                linestyle=ls,
                alpha=alpha,
                zorder=5,
            )
            ax_main.add_patch(poly)

            # Label to the right
            # Compute available width until end of overall range
            x_text0 = x + half_w + 0.2
            x_text1 = min(x_text0 + 14, x1)  # up to ~2 weeks of label real estate
            p0 = ax_main.transData.transform((x_text0, y0))
            p1 = ax_main.transData.transform((x_text1, y1))
            wpx, hpx = max(p1[0] - p0[0], 20), max(p1[1] - p0[1], 14)
            text, fs = fit_text(t.title, None, wpx, hpx)
            ax_main.text(
                x_text0,
                cy,
                text,
                fontsize=fs,
                ha="left",
                va="center",
                color=text_color_default,
                zorder=6,
            )
            continue

        # Block
        bx0, bx1 = block_span_inclusive(t.start_date, t.end_date)
        width = bx1 - bx0
        # Minimum visible width ~0.8 day for very short items
        if width < 0.8:
            bx1 = bx0 + 0.8
            width = bx1 - bx0

        patch = FancyBboxPatch(
            (bx0, y0),
            width,
            (y1 - y0),
            boxstyle="round,pad=0.02,rounding_size=0.08",
            linewidth=lw,
            edgecolor=edge,
            facecolor=face2,
            linestyle=ls,
            alpha=alpha,
            zorder=5,
        )
        ax_main.add_patch(patch)

        # Status stripe (pixel-sized so it stays visible at any zoom/range)
        stripe_px = 8 if not preview else 6
        stripe_dx = _px_to_data_dx(ax_main, bx0, (y0 + y1) / 2.0, stripe_px)
        stripe_dx = max(min(stripe_dx, width * 0.35), 0.0)
        if stripe_dx > 0:
            ax_main.add_patch(
                Rectangle(
                    (bx0, y0),
                    stripe_dx,
                    (y1 - y0),
                    facecolor=stripe_color,
                    edgecolor="none",
                    alpha=0.95,
                    zorder=6,
                )
            )

        # Text fitting based on pixel box
        p0 = ax_main.transData.transform((bx0, y0))
        p1 = ax_main.transData.transform((bx1, y1))
        wpx = max(p1[0] - p0[0], 10)
        hpx = max(p1[1] - p0[1], 10)
        text, fs = fit_text(t.title, t.description, wpx, hpx)

        # Padding from the left edge, after the status stripe.
        pad_dx = _px_to_data_dx(ax_main, bx0, (y0 + y1) / 2.0, 4)
        text_x = bx0 + stripe_dx + max(pad_dx, 0.05)
        ax_main.text(
            text_x,
            y0 + (y1 - y0) * 0.5,
            text,
            fontsize=fs,
            ha="left",
            va="center",
            color=text_color_default if status != "done" else "#4A4A4A",
            zorder=6,
        )

    # Legend (only if it adds value)
    statuses_present = {((t.status or "planned").lower()) for t in visible_tasks}
    if len(statuses_present) > 1:
        order = {"planned": 0, "in_progress": 1, "risk": 2, "done": 3}
        statuses_sorted = sorted(statuses_present, key=lambda s: order.get(s, 99))

        ax_main.text(
            0.0,
            -0.06,
            "Legend:",
            transform=ax_main.transAxes,
            ha="left",
            va="center",
            fontsize=9 if not preview else 8,
        )

        x_cursor = 0.10
        sample_w = 0.030
        sample_h = 0.035
        stripe_frac = 0.28

        label_map = {
            "planned": "Planned",
            "in_progress": "In progress",
            "done": "Done",
            "risk": "Risk",
        }

        for s in statuses_sorted:
            style = STATUS_STYLE.get(s, STATUS_STYLE["planned"])
            edge = style["edge"]
            lw = float(style["lw"])
            ls = style["ls"]
            stripe_color = style["stripe"]

            # Mini block sample
            ax_main.add_patch(
                Rectangle(
                    (x_cursor, -0.085),
                    sample_w,
                    sample_h,
                    transform=ax_main.transAxes,
                    facecolor="#FFFFFF",
                    edgecolor=edge,
                    linewidth=lw,
                    linestyle=ls,
                    clip_on=False,
                )
            )
            ax_main.add_patch(
                Rectangle(
                    (x_cursor, -0.085),
                    sample_w * stripe_frac,
                    sample_h,
                    transform=ax_main.transAxes,
                    facecolor=stripe_color,
                    edgecolor="none",
                    clip_on=False,
                )
            )

            ax_main.text(
                x_cursor + sample_w + 0.01,
                -0.067,
                label_map.get(s, s),
                transform=ax_main.transAxes,
                ha="left",
                va="center",
                fontsize=8 if not preview else 7,
            )
            x_cursor += 0.19

    return fig, warnings, hidden