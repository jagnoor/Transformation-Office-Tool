"""Sequencing Block Diagram renderer — Tetris-style space-filling layout.

Inspired by Tetris: blocks drop into position and fill the entire available
rectangle with no gaps. The algorithm uses a heightmap (skyline) approach
where each block claims a proportional share of remaining vertical space
in its time range.

Layout:
1. Items sorted by delivery sequence (label) then start date
2. Heightmap tracks the "floor" at each day across the timeline
3. Each block drops to its floor level and takes a proportional share
   of remaining vertical space — guaranteeing full coverage
4. Sequencing flows top-to-bottom, then left-to-right

Visual style:
- Bold category colors with 3D beveled edges (Tetris-like)
- Two-row timeline header (quarters + months)
- Left-side vertical legend
- Prominent bullet points inside blocks
- Checkmark icons on completed items
"""
import io
import math
import re
from collections import defaultdict
from datetime import date, timedelta
from typing import List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle
import matplotlib.dates as mdates
import numpy as np

from models import WorkItem, ChartConfig, STATUS_COLORS


# ── Color utilities ──────────────────────────────────────────────────────────

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


# ── Description parsing ──────────────────────────────────────────────────────

def _parse_bullets(description: str) -> List[str]:
    """Parse description text into individual bullet points."""
    if not description or not description.strip():
        return []

    if "\n" in description:
        parts = description.split("\n")
    elif ";" in description:
        parts = description.split(";")
    else:
        parts = [description]

    bullets = []
    for part in parts:
        part = part.strip()
        if part and part[0] in "-*+":
            part = part[1:].strip()
        if part and len(part) > 1:
            i = 0
            while i < len(part) and part[i].isdigit():
                i += 1
            if i > 0 and i < len(part) and part[i] in ".)":
                part = part[i + 1:].strip()
        if part:
            bullets.append(part)

    return bullets[:8]


def _extract_sequence_number(label: str) -> int:
    """Extract numeric sequence from label like 'Delivery 3' -> 3."""
    if not label:
        return 999
    match = re.search(r'(\d+)', label)
    return int(match.group(1)) if match else 999


# ── Two-row timeline computation ─────────────────────────────────────────────

def _compute_two_row_timeline(chart_start: date, chart_end: date):
    """Compute quarter and month column boundaries for a two-row header."""
    month_columns = []
    current = chart_start.replace(day=1)
    while current <= chart_end:
        next_month = (current.replace(day=28) + timedelta(days=4)).replace(day=1)
        col_start = max(current, chart_start)
        col_end = min(next_month - timedelta(days=1), chart_end)
        label = current.strftime("%b")
        month_columns.append((col_start, col_end, label))
        current = next_month

    quarter_columns = []
    if not month_columns:
        return quarter_columns, month_columns

    first_date = month_columns[0][0]
    last_date = month_columns[-1][1]

    q_month = ((first_date.month - 1) // 3) * 3 + 1
    current = date(first_date.year, q_month, 1)

    while current <= last_date:
        q_num = (current.month - 1) // 3 + 1
        q_end_month = current.month + 2
        q_end_year = current.year
        if q_end_month > 12:
            q_end_month -= 12
            q_end_year += 1
        next_q = date(q_end_year, q_end_month, 1)
        next_q = (next_q.replace(day=28) + timedelta(days=4)).replace(day=1)
        q_end = next_q - timedelta(days=1)

        col_start = max(current, chart_start)
        col_end = min(q_end, chart_end)
        label = f"Q{q_num} {current.year}"
        quarter_columns.append((col_start, col_end, label))
        current = next_q

    return quarter_columns, month_columns


# ── Tetris layout algorithm — heightmap space-filling ────────────────────────

def _layout_sequencing(
    items: List[WorkItem],
    chart_start: date,
    chart_end: date,
    content_w_inches: float,
    content_h_inches: float,
    base_line_h: float,
    font_size: float,
) -> List[dict]:
    """Tetris-style space-filling layout using a heightmap.

    Each block drops to its current floor level and claims a proportional
    share of the remaining vertical space. This guarantees the entire
    rectangle is filled with no gaps.

    Returns list of dicts with keys:
        item, x_start, x_end, y_top, y_bot, height
    Coordinates in normalized axes units (0-1). y=1 is top, y=0 is bottom.
    """
    total_days = (chart_end - chart_start).days
    if total_days <= 0:
        return []

    # Sort by sequence number (Delivery 1, 2, 3...) then by start date
    sorted_items = sorted(
        items,
        key=lambda it: (_extract_sequence_number(it.label), it.start_date, it.title)
    )

    gap = 0.004  # tiny gap between blocks for grid effect

    # Heightmap: tracks how much vertical space is used at each day
    # Value = fraction of height consumed (0.0 = nothing placed, 1.0 = full)
    heightmap = np.zeros(total_days + 1, dtype=float)

    # Pre-compute: for each day, count how many items will occupy it
    # This lets us divide space proportionally
    day_item_count = np.zeros(total_days + 1, dtype=float)
    item_day_ranges = []

    for item in sorted_items:
        d_start = max(0, (item.start_date - chart_start).days)
        d_end = min(total_days, (item.end_date - chart_start).days)
        if d_end <= d_start:
            d_end = d_start + 1
        item_day_ranges.append((d_start, d_end))
        day_item_count[d_start:d_end] += 1

    # Ensure no zeros (avoid division by zero)
    day_item_count = np.maximum(day_item_count, 1)

    layouts = []
    for idx, item in enumerate(sorted_items):
        d_start, d_end = item_day_ranges[idx]

        x_start = d_start / total_days
        x_end = d_end / total_days

        x_s = x_start + gap
        x_e = x_end - gap
        if x_e <= x_s:
            x_e = x_s + gap

        # Find the current floor (max heightmap value in this range)
        floor = float(np.max(heightmap[d_start:d_end]))

        # Count remaining items (including this one) that still need space
        # at each day in this range. Use the maximum overlap count.
        remaining_at_days = day_item_count[d_start:d_end].copy()
        # The share for this block: proportional to 1/remaining
        # Use the max remaining count to ensure uniform height across the block
        max_remaining = float(np.max(remaining_at_days))
        share = (1.0 - floor) / max(1, max_remaining)

        # Enforce a minimum height so text is readable
        min_h = 0.04
        height = max(share, min_h)

        # Don't exceed remaining space
        height = min(height, 1.0 - floor)

        y_top = 1.0 - floor - gap
        y_bot = 1.0 - floor - height + gap
        if y_bot >= y_top:
            y_bot = y_top - min_h

        # Update heightmap — mark this space as consumed
        heightmap[d_start:d_end] += height

        # Decrement remaining item count for these days
        day_item_count[d_start:d_end] -= 1
        day_item_count = np.maximum(day_item_count, 0)

        layouts.append({
            "item": item,
            "x_start": x_s,
            "x_end": x_e,
            "y_top": y_top,
            "y_bot": y_bot,
            "height": y_top - y_bot,
        })

    return layouts


# ── Main renderer ────────────────────────────────────────────────────────────

def render_sequencing_diagram(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    slide_aspect: str = "16:9",
) -> bytes:
    """Render a Tetris-style sequencing diagram and return PNG bytes."""
    if not items:
        raise ValueError("No work items to render")

    fig = _render_sequencing_figure(items, config, dpi, slide_aspect)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none",
                pad_inches=0.1)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_sequencing_pdf(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 300,
    slide_aspect: str = "16:9",
) -> bytes:
    """Render sequencing diagram as vector PDF."""
    if not items:
        raise ValueError("No work items to render")

    fig = _render_sequencing_figure(items, config, dpi, slide_aspect)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), pad_inches=0.1)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _render_sequencing_figure(items, config, dpi, slide_aspect="16:9"):
    """Build and return a matplotlib Figure for the Tetris sequencing diagram."""
    font_name = _resolve_font(config.font_family)
    plt.rcParams.update({
        "font.family": font_name,
        "font.size": 10,
        "axes.unicode_minus": False,
    })

    # Get categories (preserve order)
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
        fig_w, fig_h = 20, 11
    elif slide_aspect == "4:3":
        fig_w, fig_h = 16, 12
    else:
        fig_w, fig_h = 20, 11

    # Layout proportions
    header_top = 0.94
    quarter_top = 0.91
    quarter_bot = 0.885
    month_top = quarter_bot
    month_bot = 0.86
    content_top = month_bot - 0.005
    content_bot = 0.04
    legend_left_x = 0.02
    legend_width = 0.08
    content_left_x = legend_left_x + legend_width + 0.01
    content_right_x = 0.98
    content_x_width = content_right_x - content_left_x

    content_h_frac = content_top - content_bot
    content_w_inches = fig_w * content_x_width
    content_h_inches = fig_h * content_h_frac

    base_font_size = 8.5
    base_line_h_inches = base_font_size * 0.016
    base_line_h = base_line_h_inches / content_h_inches

    layouts = _layout_sequencing(
        items, chart_start, chart_end,
        content_w_inches, content_h_inches,
        base_line_h, base_font_size,
    )

    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi, facecolor=config.background_color)

    # ── Header ───────────────────────────────────────────────────────────
    ax_header = fig.add_axes([0.02, header_top, 0.96, 0.055])
    ax_header.set_xlim(0, 1)
    ax_header.set_ylim(0, 1)
    ax_header.axis("off")

    ax_header.text(0.0, 0.65, config.title, fontsize=24, fontweight="bold",
                   color="#0F172A", fontfamily=font_name, va="center")
    if config.subtitle:
        ax_header.text(0.0, 0.1, config.subtitle, fontsize=13, color="#64748B",
                       fontfamily=font_name, va="center")

    date_text = f"{chart_start.strftime('%b %Y')} \u2014 {chart_end.strftime('%b %Y')}"
    ax_header.text(1.0, 0.5, date_text, fontsize=11, color="#94A3B8",
                   fontfamily=font_name, va="center", ha="right")

    # ── Two-row timeline header ──────────────────────────────────────────
    quarter_columns, month_columns = _compute_two_row_timeline(chart_start, chart_end)

    # Quarter row
    ax_quarter = fig.add_axes([content_left_x, quarter_bot, content_x_width, quarter_top - quarter_bot])
    ax_quarter.set_xlim(0, 1)
    ax_quarter.set_ylim(0, 1)
    ax_quarter.axis("off")

    q_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="square,pad=0",
        facecolor="#1E293B", edgecolor="none", zorder=0,
    )
    ax_quarter.add_patch(q_bg)

    for i, (col_start, col_end, label) in enumerate(quarter_columns):
        x_start = (col_start - chart_start).days / total_days
        x_end = (col_end - chart_start).days / total_days
        x_mid = (x_start + x_end) / 2

        if i % 2 == 0:
            col_bg = FancyBboxPatch(
                (x_start, 0), x_end - x_start, 1,
                boxstyle="square,pad=0", facecolor="#334155", edgecolor="none", zorder=1,
            )
            ax_quarter.add_patch(col_bg)

        ax_quarter.text(x_mid, 0.5, label, fontsize=10, color="#F1F5F9",
                        ha="center", va="center", fontfamily=font_name,
                        fontweight="bold", zorder=2)

        if i > 0:
            ax_quarter.axvline(x=x_start, color="#475569", linewidth=0.5, zorder=2)

    # Month row
    ax_month = fig.add_axes([content_left_x, month_bot, content_x_width, month_top - month_bot])
    ax_month.set_xlim(0, 1)
    ax_month.set_ylim(0, 1)
    ax_month.axis("off")

    m_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="square,pad=0",
        facecolor="#475569", edgecolor="none", zorder=0,
    )
    ax_month.add_patch(m_bg)

    for i, (col_start, col_end, label) in enumerate(month_columns):
        x_start = (col_start - chart_start).days / total_days
        x_end = (col_end - chart_start).days / total_days
        x_mid = (x_start + x_end) / 2

        if i % 2 == 0:
            col_bg = FancyBboxPatch(
                (x_start, 0), x_end - x_start, 1,
                boxstyle="square,pad=0", facecolor="#546378", edgecolor="none", zorder=1,
            )
            ax_month.add_patch(col_bg)

        ax_month.text(x_mid, 0.5, label, fontsize=8.5, color="#E2E8F0",
                      ha="center", va="center", fontfamily=font_name,
                      fontweight="medium", zorder=2)

        if i > 0:
            ax_month.axvline(x=x_start, color="#64748B", linewidth=0.5, zorder=2)

    # ── Left-side vertical legend ────────────────────────────────────────
    if config.show_legend:
        ax_legend = fig.add_axes([legend_left_x, content_bot, legend_width, content_top - content_bot])
        ax_legend.set_xlim(0, 1)
        ax_legend.set_ylim(0, 1)
        ax_legend.axis("off")

        legend_bg = FancyBboxPatch(
            (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.02",
            facecolor="#F1F5F9", edgecolor="#E2E8F0", linewidth=0.5, zorder=0,
        )
        ax_legend.add_patch(legend_bg)

        ax_legend.text(0.5, 0.97, "LEGEND", fontsize=7.5, color="#64748B",
                       ha="center", va="top", fontfamily=font_name,
                       fontweight="bold", zorder=2)

        num_cats = len(categories)
        if num_cats > 0:
            spacing = min(0.08, 0.85 / num_cats)
            y_pos = 0.92

            for cat in categories:
                color = config.get_category_color(cat, categories)

                swatch = FancyBboxPatch(
                    (0.08, y_pos - 0.025), 0.18, 0.035,
                    boxstyle="round,pad=0.001,rounding_size=0.01",
                    facecolor=color, edgecolor="none", zorder=5,
                )
                ax_legend.add_patch(swatch)

                ax_legend.text(0.35, y_pos - 0.008, cat, fontsize=7,
                               color="#334155", va="center", ha="left",
                               fontfamily=font_name, fontweight="medium",
                               zorder=5, clip_on=True)

                y_pos -= spacing

    # ── Content area ─────────────────────────────────────────────────────
    ax_content = fig.add_axes([content_left_x, content_bot, content_x_width, content_top - content_bot])
    ax_content.set_xlim(0, 1)
    ax_content.set_ylim(0, 1)
    ax_content.axis("off")
    ax_content.set_facecolor("#F0F2F5")

    # Content background — light gray to make blocks pop
    content_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="square,pad=0",
        facecolor="#F0F2F5", edgecolor="#CBD5E1", linewidth=1.0, zorder=0,
    )
    ax_content.add_patch(content_bg)

    # Subtle vertical grid lines from month boundaries
    for i, (col_start, col_end, label) in enumerate(month_columns):
        x = (col_start - chart_start).days / total_days
        if i > 0:
            ax_content.axvline(x=x, color="#D1D5DB", linewidth=0.3, zorder=1, alpha=0.5)

    # ── Draw Tetris blocks ───────────────────────────────────────────────
    for layout in layouts:
        item = layout["item"]
        x_s = layout["x_start"]
        x_e = layout["x_end"]
        y_top = layout["y_top"]
        y_bot = layout["y_bot"]
        bw = x_e - x_s
        bh = y_top - y_bot

        if bh <= 0 or bw <= 0:
            continue

        # Color — use bold saturated colors (Tetris-style)
        base_color = config.get_category_color(item.category, categories)
        bar_color = item.color_override if item.color_override else base_color
        highlight_color = _lighten_color(bar_color, 0.3)
        shadow_color = _darken_color(bar_color, 0.25)
        text_color = _text_color_for_bg(bar_color)

        if item.is_milestone:
            mid_x = (x_s + x_e) / 2
            mid_y = (y_top + y_bot) / 2
            size_x = min(0.012, bw / 2)
            size_y = bh * 0.35
            diamond = plt.Polygon([
                [mid_x, mid_y + size_y], [mid_x + size_x, mid_y],
                [mid_x, mid_y - size_y], [mid_x - size_x, mid_y],
            ], closed=True, facecolor=bar_color, edgecolor=shadow_color,
                linewidth=1.5, zorder=5)
            ax_content.add_patch(diamond)
            ax_content.text(mid_x + size_x + 0.005, mid_y, item.title,
                            fontsize=7.5, color="#334155", va="center", ha="left",
                            fontfamily=font_name, fontweight="medium", zorder=6)
            continue

        # ── Tetris block: main fill ──────────────────────────────────
        block_rect = FancyBboxPatch(
            (x_s, y_bot), bw, bh,
            boxstyle="round,pad=0,rounding_size=0.003",
            facecolor=bar_color, edgecolor=shadow_color,
            linewidth=1.2, zorder=3,
        )
        ax_content.add_patch(block_rect)

        # ── Tetris 3D bevel: top highlight edge ─────────────────────
        bevel = 0.003
        top_edge = plt.Polygon([
            [x_s, y_top], [x_e, y_top],
            [x_e - bevel, y_top - bevel], [x_s + bevel, y_top - bevel],
        ], closed=True, facecolor=highlight_color, edgecolor="none",
            zorder=4, alpha=0.6)
        ax_content.add_patch(top_edge)

        # ── Tetris 3D bevel: left highlight edge ────────────────────
        left_edge = plt.Polygon([
            [x_s, y_top], [x_s, y_bot],
            [x_s + bevel, y_bot + bevel], [x_s + bevel, y_top - bevel],
        ], closed=True, facecolor=highlight_color, edgecolor="none",
            zorder=4, alpha=0.4)
        ax_content.add_patch(left_edge)

        # ── Tetris 3D bevel: bottom shadow edge ────────────────────
        bot_edge = plt.Polygon([
            [x_s, y_bot], [x_e, y_bot],
            [x_e - bevel, y_bot + bevel], [x_s + bevel, y_bot + bevel],
        ], closed=True, facecolor=shadow_color, edgecolor="none",
            zorder=4, alpha=0.4)
        ax_content.add_patch(bot_edge)

        # ── Tetris 3D bevel: right shadow edge ─────────────────────
        right_edge = plt.Polygon([
            [x_e, y_top], [x_e, y_bot],
            [x_e - bevel, y_bot + bevel], [x_e - bevel, y_top - bevel],
        ], closed=True, facecolor=shadow_color, edgecolor="none",
            zorder=4, alpha=0.3)
        ax_content.add_patch(right_edge)

        # ── Text inside block ────────────────────────────────────────
        fig_w_inches = fig.get_figwidth()
        block_w_inches = bw * fig_w_inches * content_x_width
        block_h_inches = bh * fig.get_figheight() * content_h_frac

        # Font sizing based on block dimensions
        if block_h_inches >= 1.5 and block_w_inches >= 2.5:
            title_fs = 9.5
        elif block_h_inches >= 0.8 and block_w_inches >= 1.5:
            title_fs = 8.5
        elif block_h_inches >= 0.4 and block_w_inches >= 0.8:
            title_fs = 7.5
        elif block_w_inches >= 0.5:
            title_fs = 6.5
        else:
            title_fs = 5.5

        bullet_fs = max(5, title_fs - 1.5)

        chars_per_inch = title_fs * 0.85
        max_chars = max(5, int(block_w_inches * chars_per_inch))

        line_h = base_line_h_inches / (fig.get_figheight() * content_h_frac)

        text_margin_x = bw * 0.05
        text_margin_y = bh * 0.08
        text_x = x_s + text_margin_x + bevel
        text_y = y_top - text_margin_y - bevel

        # Skip text for tiny blocks
        if block_w_inches < 0.35 or block_h_inches < 0.15:
            continue

        # Estimate how many text lines fit
        available_h = bh - text_margin_y * 2 - bevel * 2
        max_text_lines = max(1, int(available_h / (line_h * 1.15)))

        lines_used = 0

        # Checkmark for completed items
        if item.status == "done":
            check_x = x_e - text_margin_x - bevel
            check_y = y_top - text_margin_y - bevel
            ax_content.text(check_x, check_y, "\u2713", fontsize=title_fs + 2,
                            color="#FFFFFF" if text_color == "#FFFFFF" else "#10B981",
                            va="top", ha="right",
                            fontfamily=font_name, fontweight="bold", zorder=6,
                            alpha=0.9)

        # Label (e.g., "Delivery 1")
        if item.label and max_text_lines >= 2:
            ax_content.text(
                text_x, text_y, item.label,
                fontsize=max(5, title_fs - 0.5), color=text_color,
                va="top", ha="left", fontfamily=font_name,
                fontweight="bold", style="italic",
                zorder=5, alpha=0.9,
            )
            text_y -= line_h * 1.15
            lines_used += 1

        # Title (bold)
        if lines_used < max_text_lines:
            wrapped_title = _wrap_text(item.title, max_chars)
            title_lines = wrapped_title.split("\n")
            avail = max(1, min(len(title_lines), max_text_lines - lines_used))
            if avail < len(title_lines):
                title_lines = title_lines[:avail]
                last = title_lines[-1]
                if len(last) > 2:
                    title_lines[-1] = last[:-1] + "\u2026"
            display_title = "\n".join(title_lines)

            ax_content.text(
                text_x, text_y, display_title,
                fontsize=title_fs, color=text_color,
                va="top", ha="left", fontfamily=font_name,
                fontweight="bold", zorder=5,
                linespacing=1.15,
            )
            lines_used += len(title_lines)
            text_y -= len(title_lines) * line_h * 1.15

        # Bullet points from description
        bullets = _parse_bullets(item.description)
        remaining_lines = max_text_lines - lines_used
        if bullets and remaining_lines >= 1 and block_w_inches >= 0.6:
            bullet_chars = max(5, int(block_w_inches * bullet_fs * 0.85))
            for bullet in bullets[:remaining_lines]:
                if len(bullet) > bullet_chars:
                    bullet = bullet[:bullet_chars - 1] + "\u2026"
                bullet_text = f"\u2022 {bullet}"

                ax_content.text(
                    text_x, text_y, bullet_text,
                    fontsize=bullet_fs, color=text_color,
                    va="top", ha="left", fontfamily=font_name,
                    zorder=5, alpha=0.85, linespacing=1.1,
                )
                text_y -= line_h * 1.05

    # ── Today line ───────────────────────────────────────────────────────
    if config.show_today_line:
        today = config.today_date or date.today()
        if chart_start <= today <= chart_end:
            today_x = (today - chart_start).days / total_days
            ax_content.axvline(x=today_x, color="#EF4444", linewidth=2,
                               linestyle="--", zorder=15, alpha=0.7)
            ax_month.axvline(x=today_x, color="#EF4444", linewidth=2,
                             linestyle="--", zorder=15, alpha=0.7)
            ax_quarter.axvline(x=today_x, color="#EF4444", linewidth=2,
                               linestyle="--", zorder=15, alpha=0.7)

    return fig
