"""World-class Gantt chart renderer using matplotlib.

Creates a modern, visually stunning swim lane Gantt chart with:
- Clean typography and spacing
- Subtle gradients and shadows
- Category-based color coding with status indicators
- Smart date axis with automatic granularity
- Today marker line
- Professional header and legend
"""
import io
import math
from collections import defaultdict
from datetime import date, timedelta
from typing import List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.font_manager import FontProperties
import numpy as np

from models import WorkItem, ChartConfig, STATUS_COLORS, STATUS_LABELS


# ── Layout constants ─────────────────────────────────────────────────────────

ROW_HEIGHT = 0.45
ROW_GAP = 0.12
LANE_PAD = 0.25
BAR_HEIGHT = 0.38
HEADER_HEIGHT_RATIO = 0.08
LEGEND_HEIGHT_RATIO = 0.05
SIDEBAR_RATIO = 0.14


from utils import (
    lighten_color as _lighten_color,
    darken_color as _darken_color,
    text_color_for_bg as _text_color_for_bg,
    resolve_font as _resolve_font,
    compute_logo_geometry,
    place_logo,
)


def _assign_sublanes(items: List[WorkItem]) -> List[Tuple[WorkItem, int]]:
    """Assign items to sublanes within a category to avoid overlaps."""
    sorted_items = sorted(items, key=lambda it: (it.start_date, it.end_date, it.title))
    sublanes = []  # each entry is list of (item, end_date)
    result = []
    for item in sorted_items:
        placed = False
        for lane_idx, lane_items in enumerate(sublanes):
            last_end = lane_items[-1].end_date
            if item.start_date > last_end:
                lane_items.append(item)
                result.append((item, lane_idx))
                placed = True
                break
        if not placed:
            sublanes.append([item])
            result.append((item, len(sublanes) - 1))
    return result, len(sublanes)


def _compute_layout(items: List[WorkItem]) -> dict:
    """Compute full layout: categories → sublanes → y positions."""
    cats = defaultdict(list)
    for item in items:
        cats[item.category].append(item)

    # Preserve order of first appearance
    cat_order = []
    seen = set()
    for item in items:
        if item.category not in seen:
            cat_order.append(item.category)
            seen.add(item.category)

    layout = {}
    y_cursor = 0.0
    for cat in cat_order:
        cat_items = cats[cat]
        assignments, num_sublanes = _assign_sublanes(cat_items)
        lane_height = num_sublanes * (ROW_HEIGHT + ROW_GAP) + LANE_PAD * 2
        layout[cat] = {
            "y_start": y_cursor,
            "height": lane_height,
            "num_sublanes": num_sublanes,
            "assignments": assignments,
        }
        y_cursor += lane_height

    return layout, cat_order, y_cursor


def _auto_granularity(start: date, end: date) -> str:
    span = (end - start).days
    if span <= 90:
        return "weeks"
    elif span <= 365:
        return "months"
    elif span <= 730:
        return "quarters"
    else:
        return "years"


def _render_gantt_figure(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    figsize: Optional[Tuple[float, float]] = None,
):
    """Build and return a matplotlib Figure for the Gantt chart."""
    if not items:
        raise ValueError(
            "No work items to render. Load a spreadsheet or click 'Load Sample Data' to get started."
        )

    font_name = _resolve_font(config.font_family)
    plt.rcParams.update({
        "font.family": font_name,
        "font.size": 10,
        "axes.unicode_minus": False,
    })

    # Get all categories in order
    categories = []
    seen = set()
    for it in items:
        if it.category not in seen:
            categories.append(it.category)
            seen.add(it.category)

    # Compute layout
    layout, cat_order, total_height = _compute_layout(items)

    # Date range
    all_starts = [it.start_date for it in items]
    all_ends = [it.end_date for it in items]
    chart_start = config.start_date or min(all_starts) - timedelta(days=7)
    chart_end = config.end_date or max(all_ends) + timedelta(days=7)
    date_span = (chart_end - chart_start).days

    # Figure sizing
    if figsize is None:
        width = 20
        height = max(6, min(30, total_height + 2.5))
        figsize = (width, height)

    fig = plt.figure(figsize=figsize, dpi=dpi, facecolor=config.background_color)

    # Logo geometry computed up-front so the date-range badge can leave room for it
    logo_geom = compute_logo_geometry(fig, getattr(config, "logo_bytes", None))

    # Grid layout: header | main chart area
    header_h = 0.06
    legend_h = 0.04 if config.show_legend else 0.0
    chart_h = 1.0 - header_h - legend_h - 0.02

    # Header axis
    ax_header = fig.add_axes([0.0, 1.0 - header_h, 1.0, header_h])
    ax_header.set_xlim(0, 1)
    ax_header.set_ylim(0, 1)
    ax_header.axis("off")

    # Title
    ax_header.text(
        0.02, 0.65, config.title,
        fontsize=18, fontweight="bold", color="#0F172A",
        fontfamily=font_name, va="center",
    )
    if config.subtitle:
        ax_header.text(
            0.02, 0.2, config.subtitle,
            fontsize=11, color="#475569",
            fontfamily=font_name, va="center",
        )

    # Date range badge — shifted left to make room for the logo, if present
    date_range_x = 0.98 if not logo_geom else max(0.5, logo_geom["left"] - 0.02)
    date_range_text = f"{chart_start.strftime('%b %d, %Y')}  →  {chart_end.strftime('%b %d, %Y')}"
    ax_header.text(
        date_range_x, 0.5, date_range_text,
        fontsize=9, color="#64748B",
        fontfamily=font_name, va="center", ha="right",
    )

    # Separator line
    ax_header.axhline(y=0.0, xmin=0.01, xmax=0.99, color="#E2E8F0", linewidth=1.5)

    # Main chart
    sidebar_width = 0.13
    chart_left = sidebar_width + 0.01
    chart_width = 0.98 - chart_left

    ax_main = fig.add_axes([chart_left, legend_h + 0.01, chart_width, chart_h])
    ax_sidebar = fig.add_axes([0.005, legend_h + 0.01, sidebar_width, chart_h])

    # Configure main axes
    ax_main.set_xlim(chart_start, chart_end)
    ax_main.set_ylim(total_height, 0)
    ax_main.set_facecolor("#FAFBFC")

    ax_sidebar.set_xlim(0, 1)
    ax_sidebar.set_ylim(total_height, 0)
    ax_sidebar.axis("off")
    ax_sidebar.set_facecolor(config.background_color)

    # ── Draw timeline header ─────────────────────────────────────────────
    granularity = _auto_granularity(chart_start, chart_end)

    ax_main.xaxis_date()
    if granularity == "weeks":
        ax_main.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
        ax_main.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
        ax_main.xaxis.set_minor_locator(mdates.DayLocator())
    elif granularity == "months":
        ax_main.xaxis.set_major_locator(mdates.MonthLocator())
        ax_main.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    elif granularity == "quarters":
        ax_main.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
        ax_main.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    else:
        ax_main.xaxis.set_major_locator(mdates.YearLocator())
        ax_main.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

    ax_main.tick_params(axis="x", labelsize=8, colors="#475569", length=3, pad=5)
    ax_main.tick_params(axis="y", left=False, labelleft=False)

    # Subtle grid
    ax_main.grid(axis="x", color="#E2E8F0", linewidth=0.5, linestyle="-", alpha=0.7)
    ax_main.grid(axis="y", visible=False)

    # Clean up spines
    for spine in ax_main.spines.values():
        spine.set_visible(False)
    ax_main.spines["top"].set_visible(True)
    ax_main.spines["top"].set_color("#E2E8F0")
    ax_main.spines["top"].set_linewidth(1)

    # ── Draw swim lanes ──────────────────────────────────────────────────
    for i, cat in enumerate(cat_order):
        info = layout[cat]
        y_start = info["y_start"]
        height = info["height"]
        color = config.get_category_color(cat, categories)
        bg_color = _lighten_color(color, 0.92)

        # Lane background
        ax_main.axhspan(y_start, y_start + height, facecolor=bg_color, edgecolor="none", zorder=0)

        # Lane separator
        if i > 0:
            ax_main.axhline(y=y_start, color="#E2E8F0", linewidth=1, zorder=1)

        # Sidebar label
        label_y = y_start + height / 2
        # Category color badge
        badge_rect = FancyBboxPatch(
            (0.05, label_y - 0.2), 0.04, 0.4,
            boxstyle="round,pad=0.02",
            facecolor=color, edgecolor="none",
            transform=ax_sidebar.transData, zorder=5,
        )
        ax_sidebar.add_patch(badge_rect)

        ax_sidebar.text(
            0.15, label_y,
            cat, fontsize=10, fontweight="bold",
            color="#1E293B", va="center", ha="left",
            fontfamily=font_name,
        )

    # ── Draw task bars ───────────────────────────────────────────────────
    for cat in cat_order:
        info = layout[cat]
        y_start = info["y_start"]
        color = config.get_category_color(cat, categories)

        for item, sublane in info["assignments"]:
            # Determine color
            bar_color = item.color_override if item.color_override else color
            border_color = _darken_color(bar_color, 0.25)
            text_color = _text_color_for_bg(bar_color)

            # Y position
            y_pos = y_start + LANE_PAD + sublane * (ROW_HEIGHT + ROW_GAP)

            # Status indicator
            status_color = STATUS_COLORS.get(item.status, STATUS_COLORS["planned"])

            if item.is_milestone:
                # Draw diamond milestone
                mid_date = mdates.date2num(item.start_date)
                diamond_size = BAR_HEIGHT * 0.6
                diamond_y = y_pos + BAR_HEIGHT / 2
                diamond = plt.Polygon(
                    [
                        [mid_date, diamond_y - diamond_size],
                        [mid_date + 1.5, diamond_y],
                        [mid_date, diamond_y + diamond_size],
                        [mid_date - 1.5, diamond_y],
                    ],
                    closed=True,
                    facecolor=bar_color,
                    edgecolor=border_color,
                    linewidth=1.5,
                    zorder=10,
                )
                ax_main.add_patch(diamond)

                # Milestone label
                ax_main.text(
                    mid_date + 2.5, diamond_y,
                    item.title, fontsize=7.5, color="#334155",
                    va="center", ha="left", fontfamily=font_name,
                    fontweight="medium", zorder=11,
                )
            else:
                # Draw task bar
                start_num = mdates.date2num(max(item.start_date, chart_start))
                end_num = mdates.date2num(min(item.end_date, chart_end))
                bar_width = end_num - start_num

                if bar_width <= 0:
                    continue

                # Main bar (rounded rectangle)
                bar = FancyBboxPatch(
                    (start_num, y_pos),
                    bar_width, BAR_HEIGHT,
                    boxstyle="round,pad=0.02,rounding_size=0.05",
                    facecolor=bar_color,
                    edgecolor=border_color,
                    linewidth=1.0,
                    zorder=5,
                )
                ax_main.add_patch(bar)

                # Status stripe (thin left border)
                if config.show_status:
                    stripe_width = max(1.5, bar_width * 0.015)
                    stripe = FancyBboxPatch(
                        (start_num, y_pos),
                        stripe_width, BAR_HEIGHT,
                        boxstyle="round,pad=0,rounding_size=0.05",
                        facecolor=status_color,
                        edgecolor="none",
                        zorder=6,
                    )
                    ax_main.add_patch(stripe)

                # Task text
                text_x = start_num + bar_width * 0.03 + (stripe_width if config.show_status else 0)
                text_y = y_pos + BAR_HEIGHT / 2

                # Decide font size based on bar width
                char_width_est = 0.6  # approximate character width in date units
                available_chars = int(bar_width / char_width_est)

                if available_chars >= len(item.title):
                    display_text = item.title
                    font_size = 8.5
                elif available_chars >= 10:
                    display_text = item.title[:available_chars - 2] + "…"
                    font_size = 8
                elif available_chars >= 5:
                    display_text = item.title[:available_chars - 1] + "…"
                    font_size = 7
                else:
                    display_text = ""
                    font_size = 7

                if display_text:
                    ax_main.text(
                        text_x, text_y,
                        display_text,
                        fontsize=font_size, color=text_color,
                        va="center", ha="left",
                        fontfamily=font_name, fontweight="medium",
                        zorder=7, clip_on=True,
                    )

                # Label badge (if exists)
                if item.label:
                    ax_main.text(
                        end_num + 1, text_y,
                        item.label,
                        fontsize=6.5, color="#64748B",
                        va="center", ha="left",
                        fontfamily=font_name,
                        zorder=7,
                        style="italic",
                    )

    # ── Today line ───────────────────────────────────────────────────────
    if config.show_today_line:
        today = config.today_date or date.today()
        if chart_start <= today <= chart_end:
            today_num = mdates.date2num(today)
            ax_main.axvline(
                x=today_num, color="#EF4444", linewidth=1.5,
                linestyle="--", zorder=15, alpha=0.8,
            )
            ax_main.text(
                today_num, -0.15, "Today",
                fontsize=7, color="#EF4444", ha="center", va="bottom",
                fontfamily=font_name, fontweight="bold", zorder=15,
            )

    # ── Legend ────────────────────────────────────────────────────────────
    if config.show_legend:
        ax_legend = fig.add_axes([0.01, 0.002, 0.98, legend_h])
        ax_legend.set_xlim(0, 1)
        ax_legend.set_ylim(0, 1)
        ax_legend.axis("off")

        # Category legend
        x_pos = 0.01
        ax_legend.text(x_pos, 0.7, "Categories:", fontsize=8, fontweight="bold",
                       color="#64748B", va="center", fontfamily=font_name)
        x_pos += 0.07

        for cat in cat_order:
            color = config.get_category_color(cat, categories)
            badge = FancyBboxPatch(
                (x_pos, 0.35), 0.012, 0.5,
                boxstyle="round,pad=0.002",
                facecolor=color, edgecolor="none",
                transform=ax_legend.transData, zorder=5,
            )
            ax_legend.add_patch(badge)
            ax_legend.text(x_pos + 0.018, 0.6, cat, fontsize=7.5,
                           color="#334155", va="center", fontfamily=font_name)
            x_pos += 0.018 + len(cat) * 0.0065 + 0.02

        # Status legend (right side)
        if config.show_status:
            x_pos = 0.75
            ax_legend.text(x_pos, 0.7, "Status:", fontsize=8, fontweight="bold",
                           color="#64748B", va="center", fontfamily=font_name)
            x_pos += 0.05

            for status_key, status_label in STATUS_LABELS.items():
                s_color = STATUS_COLORS[status_key]
                badge = FancyBboxPatch(
                    (x_pos, 0.35), 0.012, 0.5,
                    boxstyle="round,pad=0.002",
                    facecolor=s_color, edgecolor="none",
                    transform=ax_legend.transData, zorder=5,
                )
                ax_legend.add_patch(badge)
                ax_legend.text(x_pos + 0.018, 0.6, status_label, fontsize=7.5,
                               color="#334155", va="center", fontfamily=font_name)
                x_pos += 0.018 + len(status_label) * 0.006 + 0.02

    place_logo(fig, logo_geom)

    return fig


def render_gantt(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    figsize: Optional[Tuple[float, float]] = None,
) -> bytes:
    """Render a world-class Gantt chart and return PNG bytes."""
    fig = _render_gantt_figure(items, config, dpi, figsize)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none",
                pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_gantt_pdf(items: List[WorkItem], config: ChartConfig, dpi: int = 300) -> bytes:
    """Render Gantt chart as vector PDF bytes.

    Uses the same rendering as render_gantt but outputs PDF format.
    """
    if not items:
        raise ValueError("No work items to render")

    # Temporarily render as PNG then re-render to PDF using same logic
    # We do this by calling _render_gantt_figure and saving as PDF
    fig = _render_gantt_figure(items, config, dpi)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()
