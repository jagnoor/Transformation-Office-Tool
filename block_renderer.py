"""Block Diagram renderer — space-filling timeline layout.

This creates the "On Demand Pay 2019" style visualization where work items
are packed into rows filling a single slide/page. Unlike a Gantt chart with
swim lanes, this optimizes for space utilization to show how much parallel
work is happening across the organization.

Layout algorithm:
1. Items are sorted by start date
2. Greedy row-packing: each item is placed in the first row where it fits
3. Rows fill vertically to use the full available space
4. Items are sized horizontally proportional to their duration
5. Items are sized vertically to fill available space
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
from matplotlib.patches import FancyBboxPatch, Rectangle, FancyArrowPatch
import matplotlib.dates as mdates
import numpy as np

from models import WorkItem, ChartConfig, STATUS_COLORS


def _resolve_font(preferred: str = "Arial") -> str:
    import matplotlib.font_manager as fm
    available = {f.name for f in fm.fontManager.ttflist}
    for candidate in [preferred, "Arial", "Helvetica Neue", "Helvetica", "Calibri", "DejaVu Sans"]:
        if candidate in available:
            return candidate
    return "DejaVu Sans"


def _lighten_color(hex_color: str, factor: float = 0.85) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"


def _darken_color(hex_color: str, factor: float = 0.2) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return f"#{r:02x}{g:02x}{b:02x}"


def _text_color_for_bg(hex_color: str) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#FFFFFF" if luminance < 0.55 else "#1E293B"


def _wrap_text(text: str, max_chars_per_line: int) -> str:
    """Word-wrap text to fit within a character limit per line."""
    if max_chars_per_line <= 0:
        return text
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        if current_line and len(current_line) + 1 + len(word) > max_chars_per_line:
            lines.append(current_line)
            current_line = word
        else:
            current_line = current_line + " " + word if current_line else word
    if current_line:
        lines.append(current_line)
    return "\n".join(lines)


# ── Row-packing layout algorithm ─────────────────────────────────────────────

def _pack_rows(items: List[WorkItem], chart_start: date, chart_end: date) -> List[List[WorkItem]]:
    """Pack items into rows using greedy best-fit algorithm.

    Returns list of rows, where each row is a list of non-overlapping items.
    """
    total_days = (chart_end - chart_start).days
    if total_days <= 0:
        return [[]]

    # Sort by start date, then by longest duration first (harder to place)
    sorted_items = sorted(items, key=lambda it: (it.start_date, -it.duration_days, it.title))

    rows = []  # Each row tracks occupied intervals as list of (start_frac, end_frac)
    row_items = []  # Parallel list of item lists per row

    for item in sorted_items:
        start_frac = max(0, (item.start_date - chart_start).days / total_days)
        end_frac = min(1, (item.end_date - chart_start).days / total_days)
        if end_frac <= start_frac:
            end_frac = start_frac + 1 / total_days  # minimum width

        placed = False
        for row_idx, intervals in enumerate(rows):
            # Check if item fits in this row (no overlap)
            fits = True
            for (s, e) in intervals:
                if start_frac < e and end_frac > s:  # overlap
                    fits = False
                    break
            if fits:
                intervals.append((start_frac, end_frac))
                row_items[row_idx].append(item)
                placed = True
                break

        if not placed:
            rows.append([(start_frac, end_frac)])
            row_items.append([item])

    return row_items


def _compute_time_columns(chart_start: date, chart_end: date) -> list:
    """Compute time column boundaries for the timeline header."""
    span_days = (chart_end - chart_start).days

    if span_days <= 120:
        # Monthly columns
        columns = []
        current = chart_start.replace(day=1)
        while current <= chart_end:
            next_month = (current.replace(day=28) + timedelta(days=4)).replace(day=1)
            col_start = max(current, chart_start)
            col_end = min(next_month - timedelta(days=1), chart_end)
            label = current.strftime("%b %Y")
            columns.append((col_start, col_end, label))
            current = next_month
        return columns, "months"

    elif span_days <= 548:
        # Monthly
        columns = []
        current = chart_start.replace(day=1)
        while current <= chart_end:
            next_month = (current.replace(day=28) + timedelta(days=4)).replace(day=1)
            col_start = max(current, chart_start)
            col_end = min(next_month - timedelta(days=1), chart_end)
            label = current.strftime("%b")
            columns.append((col_start, col_end, label))
            current = next_month
        return columns, "months"

    else:
        # Quarterly
        columns = []
        current = chart_start.replace(day=1)
        q_month = ((current.month - 1) // 3) * 3 + 1
        current = current.replace(month=q_month, day=1)
        while current <= chart_end:
            q_end_month = current.month + 2
            q_end_year = current.year
            if q_end_month > 12:
                q_end_month -= 12
                q_end_year += 1
            next_q = date(q_end_year, q_end_month, 1) + timedelta(days=31)
            next_q = next_q.replace(day=1)
            q_end = next_q - timedelta(days=1)
            col_start = max(current, chart_start)
            col_end = min(q_end, chart_end)
            q_num = (current.month - 1) // 3 + 1
            label = f"Q{q_num} {current.year}"
            columns.append((col_start, col_end, label))
            current = next_q
        return columns, "quarters"


# ── Main renderer ────────────────────────────────────────────────────────────

def render_block_diagram(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    slide_aspect: str = "16:9",
) -> bytes:
    """Render a space-filling block diagram and return PNG bytes."""
    if not items:
        raise ValueError("No work items to render")

    fig = _render_block_figure(items, config, dpi, slide_aspect)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none",
                pad_inches=0.1)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_block_pdf(items: List[WorkItem], config: ChartConfig, dpi: int = 300, slide_aspect: str = "16:9") -> bytes:
    """Render block diagram as vector PDF."""
    if not items:
        raise ValueError("No work items to render")

    fig = _render_block_figure(items, config, dpi, slide_aspect)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), pad_inches=0.1)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _render_block_figure(items, config, dpi, slide_aspect="16:9"):
    """Build and return a matplotlib Figure for the block diagram."""
    font_name = _resolve_font(config.font_family)
    plt.rcParams.update({
        "font.family": font_name,
        "font.size": 10,
        "axes.unicode_minus": False,
    })

    # Get categories
    categories = []
    seen = set()
    for it in items:
        if it.category not in seen:
            categories.append(it.category)
            seen.add(it.category)

    # Date range
    chart_start = config.start_date or min(it.start_date for it in items)
    chart_end = config.end_date or max(it.end_date for it in items)
    chart_start = chart_start - timedelta(days=3)
    chart_end = chart_end + timedelta(days=3)
    total_days = (chart_end - chart_start).days

    # Figure size
    if slide_aspect == "16:9":
        fig_w, fig_h = 16, 9
    elif slide_aspect == "4:3":
        fig_w, fig_h = 13.33, 10
    else:
        fig_w, fig_h = 16, 9

    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi, facecolor=config.background_color)

    # Layout proportions
    header_top = 0.93
    timeline_top = 0.89
    timeline_bot = 0.85
    content_top = timeline_bot - 0.005
    content_bot = 0.06
    legend_bot = 0.01

    # ── Header ───────────────────────────────────────────────────────────
    ax_header = fig.add_axes([0.02, header_top, 0.96, 0.06])
    ax_header.set_xlim(0, 1)
    ax_header.set_ylim(0, 1)
    ax_header.axis("off")

    ax_header.text(0.0, 0.65, config.title, fontsize=22, fontweight="bold",
                   color="#0F172A", fontfamily=font_name, va="center")
    if config.subtitle:
        ax_header.text(0.0, 0.1, config.subtitle, fontsize=12, color="#64748B",
                       fontfamily=font_name, va="center")

    date_text = f"{chart_start.strftime('%b %Y')} — {chart_end.strftime('%b %Y')}"
    ax_header.text(1.0, 0.5, date_text, fontsize=10, color="#94A3B8",
                   fontfamily=font_name, va="center", ha="right")

    # ── Timeline header band ─────────────────────────────────────────────
    ax_timeline = fig.add_axes([0.02, timeline_bot, 0.96, timeline_top - timeline_bot])
    ax_timeline.set_xlim(0, 1)
    ax_timeline.set_ylim(0, 1)
    ax_timeline.axis("off")

    timeline_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.02",
        facecolor="#1E293B", edgecolor="none", zorder=0,
    )
    ax_timeline.add_patch(timeline_bg)

    time_columns, granularity = _compute_time_columns(chart_start, chart_end)

    for i, (col_start, col_end, label) in enumerate(time_columns):
        x_start = (col_start - chart_start).days / total_days
        x_end = (col_end - chart_start).days / total_days
        x_mid = (x_start + x_end) / 2

        if i % 2 == 0:
            col_bg = FancyBboxPatch(
                (x_start, 0), x_end - x_start, 1,
                boxstyle="square,pad=0", facecolor="#334155", edgecolor="none", zorder=1,
            )
            ax_timeline.add_patch(col_bg)

        ax_timeline.text(x_mid, 0.5, label, fontsize=9, color="#F1F5F9",
                         ha="center", va="center", fontfamily=font_name,
                         fontweight="medium", zorder=2)

        if i > 0:
            ax_timeline.axvline(x=x_start, color="#475569", linewidth=0.5, zorder=2)

    # ── Content area — pack blocks ───────────────────────────────────────
    ax_content = fig.add_axes([0.02, content_bot, 0.96, content_top - content_bot])
    ax_content.set_xlim(0, 1)
    ax_content.set_ylim(0, 1)
    ax_content.axis("off")
    ax_content.set_facecolor("#F8FAFC")

    # Background
    content_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.005",
        facecolor="#F8FAFC", edgecolor="#E2E8F0", linewidth=0.5, zorder=0,
    )
    ax_content.add_patch(content_bg)

    # Vertical grid lines matching timeline
    for i, (col_start, col_end, label) in enumerate(time_columns):
        x = (col_start - chart_start).days / total_days
        if i > 0:
            ax_content.axvline(x=x, color="#E2E8F0", linewidth=0.5, zorder=1, alpha=0.5)

    # Pack items into rows
    packed_rows = _pack_rows(items, chart_start, chart_end)
    num_rows = len(packed_rows)
    if num_rows == 0:
        num_rows = 1

    # Row height — cap at a reasonable size so blocks aren't absurdly tall
    row_gap = 0.008
    total_gap = row_gap * (num_rows + 1)
    natural_row_height = (1.0 - total_gap) / num_rows

    # Cap row height: for the reference "block diagram" look, blocks should
    # be wider than tall. With few rows, limit height so content is readable.
    # Max row height of ~0.18 gives a nice dense look even with 2-3 rows.
    max_row_height = 0.22
    row_height = min(natural_row_height, max_row_height)

    # Center the rows vertically if we're not using full space
    used_height = num_rows * row_height + (num_rows + 1) * row_gap
    y_offset = (1.0 - used_height) / 2  # center vertically

    block_pad_x = 0.003
    block_pad_y = 0.003

    # Draw blocks
    for row_idx, row_items in enumerate(packed_rows):
        y_top = 1.0 - y_offset - row_gap - row_idx * (row_height + row_gap)
        y_bot = y_top - row_height

        for item in row_items:
            x_start = max(0, (item.start_date - chart_start).days / total_days)
            x_end = min(1, (item.end_date - chart_start).days / total_days)
            if x_end <= x_start:
                x_end = x_start + 1 / total_days

            x_s = x_start + block_pad_x
            x_e = x_end - block_pad_x
            if x_e <= x_s:
                x_e = x_s + block_pad_x

            bw = x_e - x_s
            bh = row_height - block_pad_y * 2

            # Color
            base_color = config.get_category_color(item.category, categories)
            bar_color = item.color_override if item.color_override else base_color
            border_color = _darken_color(bar_color, 0.15)
            text_color = _text_color_for_bg(bar_color)

            if item.is_milestone:
                mid_x = (x_s + x_e) / 2
                mid_y = (y_top + y_bot) / 2
                size_x = min(0.012, bw / 2)
                size_y = bh * 0.35
                diamond = plt.Polygon([
                    [mid_x, mid_y + size_y], [mid_x + size_x, mid_y],
                    [mid_x, mid_y - size_y], [mid_x - size_x, mid_y],
                ], closed=True, facecolor=bar_color, edgecolor=border_color,
                    linewidth=1.5, zorder=5)
                ax_content.add_patch(diamond)
                ax_content.text(mid_x + size_x + 0.005, mid_y, item.title,
                                fontsize=7, color="#334155", va="center", ha="left",
                                fontfamily=font_name, fontweight="medium", zorder=6)
                continue

            # Main block rectangle
            block_rect = FancyBboxPatch(
                (x_s, y_bot + block_pad_y), bw, bh,
                boxstyle="round,pad=0,rounding_size=0.005",
                facecolor=bar_color, edgecolor=border_color,
                linewidth=0.8, zorder=3,
            )
            ax_content.add_patch(block_rect)

            # Status indicator (thin top bar)
            status_color = STATUS_COLORS.get(item.status, "#94A3B8")
            if config.show_status and item.status != "planned":
                stripe_h = min(0.006, bh * 0.08)
                stripe = FancyBboxPatch(
                    (x_s, y_top - block_pad_y - stripe_h), bw, stripe_h,
                    boxstyle="round,pad=0,rounding_size=0.002",
                    facecolor=status_color, edgecolor="none", zorder=4,
                )
                ax_content.add_patch(stripe)

            # ── Text inside block ────────────────────────────────────────
            # Scale font sizes relative to figure dimensions
            # At 16" wide, 0.01 in x ≈ 0.16". A char at fontsize 9 ≈ 0.07"
            fig_w_inches = fig.get_figwidth()
            content_w_inches = fig_w_inches * 0.96  # axes width fraction
            block_w_inches = bw * content_w_inches
            content_h_inches = fig.get_figheight() * (content_top - content_bot)
            block_h_inches = bh * content_h_inches

            # Approximate characters that fit
            # At fontsize 9: ~7 chars per inch. At fontsize 8: ~8. At fontsize 7: ~9.
            if block_w_inches >= 2.5:
                title_fs = 10
            elif block_w_inches >= 1.5:
                title_fs = 9
            elif block_w_inches >= 0.8:
                title_fs = 8
            elif block_w_inches >= 0.4:
                title_fs = 7
            else:
                title_fs = 6

            chars_per_inch = title_fs * 0.9
            max_chars_per_line = max(3, int(block_w_inches * chars_per_inch))

            # Lines that fit vertically (approx 0.18" per line at fontsize 9)
            line_h_inches = title_fs * 0.018
            max_lines = max(1, int(block_h_inches / line_h_inches * 0.75))

            # Margin inside block
            text_margin_x = bw * 0.04
            text_margin_y = bh * 0.08
            text_x = x_s + text_margin_x
            text_y = y_top - block_pad_y - text_margin_y

            # Don't render text if block is too narrow
            if block_w_inches < 0.3:
                continue

            lines_used = 0

            # Label (e.g., "Delivery 1", "D1")
            if item.label and max_lines >= 2:
                ax_content.text(
                    text_x, text_y, item.label,
                    fontsize=title_fs - 0.5, color=text_color,
                    va="top", ha="left", fontfamily=font_name,
                    fontweight="bold", style="italic",
                    zorder=5, alpha=0.85,
                )
                text_y -= line_h_inches / content_h_inches * 1.1
                lines_used += 1

            # Title (bold)
            title_text = item.title
            wrapped_title = _wrap_text(title_text, max_chars_per_line)
            title_lines = wrapped_title.split("\n")
            available_title_lines = max(1, min(len(title_lines), max_lines - lines_used))
            if available_title_lines < len(title_lines):
                title_lines = title_lines[:available_title_lines]
                last = title_lines[-1]
                if len(last) > 2:
                    title_lines[-1] = last[:-1] + "…"
            display_title = "\n".join(title_lines)

            ax_content.text(
                text_x, text_y, display_title,
                fontsize=title_fs, color=text_color,
                va="top", ha="left", fontfamily=font_name,
                fontweight="bold", zorder=5,
                linespacing=1.2,
            )
            lines_used += len(title_lines)
            text_y -= len(title_lines) * line_h_inches / content_h_inches * 1.15

            # Description (smaller, lighter)
            remaining_lines = max_lines - lines_used - 1
            if item.description and remaining_lines >= 1 and block_w_inches >= 1.0:
                desc_fs = max(5.5, title_fs - 1.5)
                desc_chars = max(3, int(block_w_inches * desc_fs * 0.9))
                desc_wrapped = _wrap_text(item.description, desc_chars)
                desc_lines = desc_wrapped.split("\n")[:remaining_lines]

                # Add bullet points
                desc_display = "\n".join("• " + l for l in desc_lines)

                desc_alpha = 0.75 if text_color == "#FFFFFF" else 0.55
                ax_content.text(
                    text_x, text_y, desc_display,
                    fontsize=desc_fs, color=text_color,
                    va="top", ha="left", fontfamily=font_name,
                    zorder=5, alpha=desc_alpha, linespacing=1.15,
                )

    # ── Today line ───────────────────────────────────────────────────────
    if config.show_today_line:
        today = config.today_date or date.today()
        if chart_start <= today <= chart_end:
            today_x = (today - chart_start).days / total_days
            ax_content.axvline(x=today_x, color="#EF4444", linewidth=2,
                               linestyle="--", zorder=15, alpha=0.7)
            ax_timeline.axvline(x=today_x, color="#EF4444", linewidth=2,
                                linestyle="--", zorder=15, alpha=0.7)

    # ── Legend / Footer ──────────────────────────────────────────────────
    if config.show_legend:
        ax_legend = fig.add_axes([0.02, legend_bot, 0.96, content_bot - legend_bot - 0.005])
        ax_legend.set_xlim(0, 1)
        ax_legend.set_ylim(0, 1)
        ax_legend.axis("off")

        x_pos = 0.0
        for cat in categories:
            color = config.get_category_color(cat, categories)
            swatch = FancyBboxPatch(
                (x_pos, 0.15), 0.015, 0.7,
                boxstyle="round,pad=0.001,rounding_size=0.05",
                facecolor=color, edgecolor="none", zorder=5,
            )
            ax_legend.add_patch(swatch)
            ax_legend.text(x_pos + 0.022, 0.5, cat, fontsize=8.5, color="#334155",
                           va="center", ha="left", fontfamily=font_name, fontweight="medium")
            x_pos += 0.022 + len(cat) * 0.007 + 0.025

        if config.show_status:
            x_pos = max(x_pos + 0.02, 0.7)
            for status_key, label in [("in_progress", "In Progress"), ("done", "Done"), ("at_risk", "At Risk")]:
                s_color = STATUS_COLORS[status_key]
                swatch = FancyBboxPatch(
                    (x_pos, 0.15), 0.015, 0.7,
                    boxstyle="round,pad=0.001,rounding_size=0.05",
                    facecolor=s_color, edgecolor="none", zorder=5,
                )
                ax_legend.add_patch(swatch)
                ax_legend.text(x_pos + 0.022, 0.5, label, fontsize=8, color="#64748B",
                               va="center", ha="left", fontfamily=font_name)
                x_pos += 0.022 + len(label) * 0.006 + 0.02

    return fig
