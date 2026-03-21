"""Block Diagram renderer — space-filling mosaic timeline layout.

This creates the "On Demand Pay 2019" style visualization where work items
are packed into a mosaic filling a single slide/page. Unlike a Gantt chart with
swim lanes, this optimizes for space utilization to show how much parallel
work is happening across the organization.

Layout algorithm:
1. Items are sorted by start date and packed into rows (greedy best-fit)
2. Horizontal gaps are filled: blocks expand to cover all horizontal space
3. Vertical gaps are filled: blocks expand downward into unoccupied rows
4. Result is a dense mosaic with no visible gaps between blocks
"""
import io
import math
from dataclasses import dataclass
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


@dataclass
class BlockRect:
    """Final layout geometry for a block in the mosaic."""
    item: WorkItem
    x_left: float     # 0..1 in content area
    x_right: float    # 0..1 in content area
    row_top: int       # starting row index (0 = topmost)
    row_bottom: int    # ending row index (inclusive)


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

def _pack_rows(items: List[WorkItem], chart_start: date, chart_end: date) -> List[List[Tuple[WorkItem, float, float]]]:
    """Pack items into rows using greedy best-fit algorithm.

    Returns list of rows, where each row is a list of (item, x_start, x_end) tuples.
    """
    total_days = (chart_end - chart_start).days
    if total_days <= 0:
        return [[]]

    # Sort by start date, then by longest duration first (harder to place)
    sorted_items = sorted(items, key=lambda it: (it.start_date, -it.duration_days, it.title))

    rows = []       # Each row tracks occupied intervals as list of (start_frac, end_frac)
    row_items = []   # Parallel list of (item, start_frac, end_frac) tuples per row

    for item in sorted_items:
        start_frac = max(0, (item.start_date - chart_start).days / total_days)
        end_frac = min(1, (item.end_date - chart_start).days / total_days)
        if end_frac <= start_frac:
            end_frac = start_frac + 1 / total_days  # minimum width

        placed = False
        for row_idx, intervals in enumerate(rows):
            fits = True
            for (s, e) in intervals:
                if start_frac < e and end_frac > s:
                    fits = False
                    break
            if fits:
                intervals.append((start_frac, end_frac))
                row_items[row_idx].append((item, start_frac, end_frac))
                placed = True
                break

        if not placed:
            rows.append([(start_frac, end_frac)])
            row_items.append([(item, start_frac, end_frac)])

    return row_items


def _fill_horizontal_gaps(rows: List[List[Tuple[WorkItem, float, float]]]) -> List[List[Tuple[WorkItem, float, float]]]:
    """Expand blocks horizontally to fill all gaps in each row.

    First item extends to x=0, last item extends to x=1, interior gaps
    are split at the midpoint between adjacent blocks.
    """
    result = []
    for row in rows:
        if not row:
            result.append(row)
            continue

        # Sort by x_start
        sorted_row = sorted(row, key=lambda t: t[1])

        if len(sorted_row) == 1:
            item, _, _ = sorted_row[0]
            result.append([(item, 0.0, 1.0)])
            continue

        # Compute new boundaries
        new_row = []
        for i, (item, x_start, x_end) in enumerate(sorted_row):
            if i == 0:
                new_left = 0.0
            else:
                # Midpoint between this item's original start and previous item's original end
                prev_end = sorted_row[i - 1][2]
                new_left = (prev_end + x_start) / 2

            if i == len(sorted_row) - 1:
                new_right = 1.0
            else:
                next_start = sorted_row[i + 1][1]
                new_right = (x_end + next_start) / 2

            new_row.append((item, new_left, new_right))

        result.append(new_row)
    return result


def _fill_vertical_gaps(rows: List[List[Tuple[WorkItem, float, float]]], grid_cols: int = 200) -> List[BlockRect]:
    """Use a 2D occupancy grid to allow blocks to span multiple rows vertically.

    After initial placement, blocks expand downward into unoccupied cells,
    creating the variable-height mosaic effect.
    """
    num_rows = len(rows)
    if num_rows == 0:
        return []

    # Build occupancy grid (num_rows x grid_cols), initially all False
    grid = [[False] * grid_cols for _ in range(num_rows)]

    # Create BlockRect for each item and mark grid cells
    blocks = []
    for row_idx, row in enumerate(rows):
        for item, x_left, x_right in row:
            col_start = max(0, int(x_left * grid_cols))
            col_end = min(grid_cols, int(x_right * grid_cols))
            if col_end <= col_start:
                col_end = col_start + 1

            # Mark cells as occupied
            for c in range(col_start, col_end):
                grid[row_idx][c] = True

            blocks.append(BlockRect(
                item=item,
                x_left=x_left,
                x_right=x_right,
                row_top=row_idx,
                row_bottom=row_idx,
            ))

    # Expand blocks downward into unoccupied cells
    # Process from bottom-to-top so lower blocks get priority first,
    # then upper blocks can expand into remaining space
    for block in sorted(blocks, key=lambda b: (-b.row_top, -b.x_left)):
        col_start = max(0, int(block.x_left * grid_cols))
        col_end = min(grid_cols, int(block.x_right * grid_cols))
        if col_end <= col_start:
            col_end = col_start + 1

        # Try to expand downward
        while block.row_bottom + 1 < num_rows:
            next_row = block.row_bottom + 1
            # Check if ALL cells in the column range are unoccupied in the next row
            can_expand = True
            for c in range(col_start, col_end):
                if grid[next_row][c]:
                    can_expand = False
                    break

            if can_expand:
                # Mark the new cells as occupied
                for c in range(col_start, col_end):
                    grid[next_row][c] = True
                block.row_bottom = next_row
            else:
                break

    return blocks


def _compute_time_columns(chart_start: date, chart_end: date) -> list:
    """Compute time column boundaries for the timeline header."""
    span_days = (chart_end - chart_start).days

    if span_days <= 120:
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

    # ── Content area — mosaic blocks ─────────────────────────────────────
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

    # Vertical grid lines matching timeline (behind blocks)
    for i, (col_start, col_end, label) in enumerate(time_columns):
        x = (col_start - chart_start).days / total_days
        if i > 0:
            ax_content.axvline(x=x, color="#E2E8F0", linewidth=0.5, zorder=1, alpha=0.3)

    # ── Mosaic layout pipeline ───────────────────────────────────────────
    packed_rows = _pack_rows(items, chart_start, chart_end)
    packed_rows = _fill_horizontal_gaps(packed_rows)
    blocks = _fill_vertical_gaps(packed_rows)
    num_rows = max(1, len(packed_rows))

    # Row height — fill all vertical space, no gaps
    row_height = 1.0 / num_rows

    # Minimal padding for thin border effect between blocks
    block_pad_x = 0.0015
    block_pad_y = 0.0015

    # Draw blocks
    for block in blocks:
        item = block.item
        x_s = block.x_left + block_pad_x
        x_e = block.x_right - block_pad_x
        if x_e <= x_s:
            x_e = x_s + block_pad_x

        # y coordinates: row_top=0 is at top of content area
        y_top = 1.0 - block.row_top * row_height - block_pad_y
        y_bot = 1.0 - (block.row_bottom + 1) * row_height + block_pad_y
        if y_bot >= y_top:
            y_bot = y_top - block_pad_y

        bw = x_e - x_s
        bh = y_top - y_bot

        # Color
        base_color = config.get_category_color(item.category, categories)
        bar_color = item.color_override if item.color_override else base_color
        border_color = _darken_color(bar_color, 0.15)
        text_color = _text_color_for_bg(bar_color)

        # All items rendered as blocks (including milestones) for mosaic consistency
        block_rect = FancyBboxPatch(
            (x_s, y_bot), bw, bh,
            boxstyle="round,pad=0,rounding_size=0.003",
            facecolor=bar_color, edgecolor=border_color,
            linewidth=0.8, zorder=3,
        )
        ax_content.add_patch(block_rect)

        # Status indicator (thin top bar)
        status_color = STATUS_COLORS.get(item.status, "#94A3B8")
        if config.show_status and item.status != "planned":
            stripe_h = min(0.005, bh * 0.06)
            stripe = FancyBboxPatch(
                (x_s, y_top - stripe_h), bw, stripe_h,
                boxstyle="round,pad=0,rounding_size=0.002",
                facecolor=status_color, edgecolor="none", zorder=4,
            )
            ax_content.add_patch(stripe)

        # Milestone diamond overlay (small diamond icon in top-right corner)
        if item.is_milestone:
            diamond_size = min(0.008, bw * 0.15, bh * 0.15)
            diamond_x = x_e - diamond_size * 2
            diamond_y = y_top - diamond_size * 2
            diamond = plt.Polygon([
                [diamond_x, diamond_y + diamond_size],
                [diamond_x + diamond_size, diamond_y],
                [diamond_x, diamond_y - diamond_size],
                [diamond_x - diamond_size, diamond_y],
            ], closed=True, facecolor="#FFFFFF", edgecolor=border_color,
                linewidth=1, zorder=5, alpha=0.8)
            ax_content.add_patch(diamond)

        # ── Text inside block ────────────────────────────────────────
        fig_w_inches = fig.get_figwidth()
        content_w_inches = fig_w_inches * 0.96
        block_w_inches = bw * content_w_inches
        content_h_inches = fig.get_figheight() * (content_top - content_bot)
        block_h_inches = bh * content_h_inches

        # Font size based on block width
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

        line_h_inches = title_fs * 0.018
        max_lines = max(1, int(block_h_inches / line_h_inches * 0.75))

        # Margin inside block
        text_margin_x = bw * 0.04
        text_margin_y = bh * 0.06
        text_x = x_s + text_margin_x
        text_y = y_top - text_margin_y

        # Don't render text if block is too narrow
        if block_w_inches < 0.3:
            continue

        lines_used = 0

        # Label (e.g., "Delivery 1")
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
                title_lines[-1] = last[:-1] + "\u2026"
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

            desc_display = "\n".join("\u2022 " + l for l in desc_lines)

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
