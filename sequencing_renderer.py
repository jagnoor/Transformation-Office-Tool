"""Sequencing Block Diagram renderer — timeline-based overlapping layout.

This creates the "On Demand Pay 2019" style visualization where work items
are positioned on a calendar timeline and allowed to overlap vertically,
making it clear to executive leadership where capacity constraints exist.

Key differences from block_renderer.py:
- Gravity-stacking layout: blocks overlap vertically when they share time
- Variable block heights based on content (title + bullet points)
- Two-row timeline header (quarters + months)
- Left-side vertical legend
- Prominent bullet points inside blocks from description field
- Checkmark icons on completed items
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
from matplotlib.patches import FancyBboxPatch, Rectangle
import matplotlib.dates as mdates
import numpy as np

from models import WorkItem, ChartConfig, STATUS_COLORS


# ── Color utilities (same as block_renderer.py) ──────────────────────────────

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
    """Parse description text into individual bullet points.

    Split on newlines first, then semicolons. Never split on commas
    (would break sentences). Strip leading bullet chars (- * +).
    """
    if not description or not description.strip():
        return []

    # Try newlines first
    if "\n" in description:
        parts = description.split("\n")
    elif ";" in description:
        parts = description.split(";")
    else:
        parts = [description]

    bullets = []
    for part in parts:
        part = part.strip()
        # Strip leading bullet markers
        if part and part[0] in "-*+":
            part = part[1:].strip()
        # Strip leading numbering like "1." or "1)"
        if part and len(part) > 1:
            i = 0
            while i < len(part) and part[i].isdigit():
                i += 1
            if i > 0 and i < len(part) and part[i] in ".)" :
                part = part[i + 1:].strip()
        if part:
            bullets.append(part)

    # Cap at 8 bullet points
    return bullets[:8]


# ── Two-row timeline computation ─────────────────────────────────────────────

def _compute_two_row_timeline(chart_start: date, chart_end: date):
    """Compute quarter and month column boundaries for a two-row header.

    Returns (quarter_columns, month_columns) where each is a list of
    (col_start_date, col_end_date, label) tuples.
    """
    # Month columns
    month_columns = []
    current = chart_start.replace(day=1)
    while current <= chart_end:
        next_month = (current.replace(day=28) + timedelta(days=4)).replace(day=1)
        col_start = max(current, chart_start)
        col_end = min(next_month - timedelta(days=1), chart_end)
        label = current.strftime("%b")
        month_columns.append((col_start, col_end, label))
        current = next_month

    # Quarter columns
    quarter_columns = []
    if not month_columns:
        return quarter_columns, month_columns

    first_date = month_columns[0][0]
    last_date = month_columns[-1][1]

    q_month = ((first_date.month - 1) // 3) * 3 + 1
    current = date(first_date.year, q_month, 1)

    while current <= last_date:
        q_num = (current.month - 1) // 3 + 1
        # End of quarter
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


# ── Layout algorithm — gravity stacking ──────────────────────────────────────

def _calc_block_height(
    item: WorkItem,
    block_w_inches: float,
    base_line_h: float,
    font_size: float,
) -> float:
    """Calculate block height in normalized axes units based on content."""
    lines = 0

    # Label line (e.g., "Delivery 1")
    if item.label:
        lines += 1

    # Title lines (estimate wrapping)
    chars_per_inch = font_size * 0.85
    max_chars = max(5, int(block_w_inches * chars_per_inch))
    title_words = item.title.split()
    title_lines = 1
    current_len = 0
    for word in title_words:
        if current_len > 0 and current_len + 1 + len(word) > max_chars:
            title_lines += 1
            current_len = len(word)
        else:
            current_len += (1 if current_len > 0 else 0) + len(word)
    lines += min(title_lines, 3)  # cap title at 3 lines

    # Bullet points
    bullets = _parse_bullets(item.description)
    lines += len(bullets)

    # Minimum 2 lines worth of height
    lines = max(2, lines)

    # Add padding (top + bottom margins ~0.6 lines)
    return (lines + 0.8) * base_line_h


def _layout_sequencing(
    items: List[WorkItem],
    chart_start: date,
    chart_end: date,
    content_w_inches: float,
    content_h_inches: float,
    base_line_h: float,
    font_size: float,
) -> List[dict]:
    """Compute layout positions using gravity-stacking algorithm.

    Returns list of dicts with keys:
        item, x_start, x_end, y_top, y_bot, height
    All coordinates are in normalized axes units (0-1).
    """
    total_days = (chart_end - chart_start).days
    if total_days <= 0:
        return []

    # Sort by start date, then longest duration first
    sorted_items = sorted(items, key=lambda it: (it.start_date, -it.duration_days, it.title))

    gap_y = 0.008  # vertical gap between blocks
    gap_x = 0.003  # horizontal padding

    placed = []  # list of (x_s, x_e, y_top, y_bot) for collision detection

    layouts = []
    for item in sorted_items:
        x_start = max(0, (item.start_date - chart_start).days / total_days)
        x_end = min(1, (item.end_date - chart_start).days / total_days)
        if x_end <= x_start:
            x_end = x_start + 1 / total_days

        x_s = x_start + gap_x
        x_e = x_end - gap_x
        if x_e <= x_s:
            x_e = x_s + gap_x

        # Compute block width in inches for height calculation
        block_w_inches = (x_e - x_s) * content_w_inches

        # Compute height
        height = _calc_block_height(item, block_w_inches, base_line_h, font_size)

        # Find y position — gravity stacking from top
        y_top = 1.0 - gap_y

        # Collect all placed blocks that overlap in x
        overlapping = []
        for (px_s, px_e, py_top, py_bot) in placed:
            if x_s < px_e and x_e > px_s:  # x-overlap
                overlapping.append((py_top, py_bot))

        # Sort overlapping blocks by y_top descending (highest first)
        overlapping.sort(key=lambda o: -o[0])

        # Try placing at y_top = 1.0 - gap, then check collisions
        for (oy_top, oy_bot) in overlapping:
            candidate_bot = y_top - height
            # Check if there's a collision
            if y_top > oy_bot and candidate_bot < oy_top:
                # Collision — move below this block
                y_top = oy_bot - gap_y

        y_bot = y_top - height

        placed.append((x_s, x_e, y_top, y_bot))
        layouts.append({
            "item": item,
            "x_start": x_s,
            "x_end": x_e,
            "y_top": y_top,
            "y_bot": y_bot,
            "height": height,
        })

    return layouts


# ── Main renderer ────────────────────────────────────────────────────────────

def render_sequencing_diagram(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    slide_aspect: str = "16:9",
) -> bytes:
    """Render a sequencing block diagram and return PNG bytes."""
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
    """Build and return a matplotlib Figure for the sequencing diagram."""
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

    # Compute layout first to determine if we need more height
    content_h_frac = content_top - content_bot
    content_w_inches = fig_w * content_x_width
    content_h_inches = fig_h * content_h_frac

    # Font size for blocks
    base_font_size = 8.5
    base_line_h_inches = base_font_size * 0.016
    base_line_h = base_line_h_inches / content_h_inches

    layouts = _layout_sequencing(
        items, chart_start, chart_end,
        content_w_inches, content_h_inches,
        base_line_h, base_font_size,
    )

    # Check if blocks overflow and adjust figure height
    if layouts:
        min_y = min(l["y_bot"] for l in layouts)
        if min_y < -0.05:
            # Scale up figure height to fit all content
            scale = content_h_frac / (content_h_frac + abs(min_y) + 0.05)
            fig_h = fig_h / scale
            # Recompute with new height
            content_h_inches = fig_h * content_h_frac
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

        # Legend background
        legend_bg = FancyBboxPatch(
            (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.02",
            facecolor="#F1F5F9", edgecolor="#E2E8F0", linewidth=0.5, zorder=0,
        )
        ax_legend.add_patch(legend_bg)

        # Legend title
        ax_legend.text(0.5, 0.97, "LEGEND", fontsize=7.5, color="#64748B",
                       ha="center", va="top", fontfamily=font_name,
                       fontweight="bold", zorder=2)

        # Category swatches stacked vertically
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
    ax_content.set_facecolor("#FFFFFF")

    # Content background
    content_bg = FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="square,pad=0",
        facecolor="#FFFFFF", edgecolor="#E2E8F0", linewidth=0.5, zorder=0,
    )
    ax_content.add_patch(content_bg)

    # Vertical grid lines from month boundaries
    for i, (col_start, col_end, label) in enumerate(month_columns):
        x = (col_start - chart_start).days / total_days
        if i > 0:
            ax_content.axvline(x=x, color="#E2E8F0", linewidth=0.5, zorder=1, alpha=0.6)

    # ── Draw blocks ──────────────────────────────────────────────────────
    for layout in layouts:
        item = layout["item"]
        x_s = layout["x_start"]
        x_e = layout["x_end"]
        y_top = layout["y_top"]
        y_bot = layout["y_bot"]
        bw = x_e - x_s
        bh = y_top - y_bot

        # Color
        base_color = config.get_category_color(item.category, categories)
        bar_color = item.color_override if item.color_override else base_color
        fill_color = _lighten_color(bar_color, 0.55)
        border_color = _darken_color(bar_color, 0.1)
        text_color = _text_color_for_bg(fill_color)

        if item.is_milestone:
            # Milestone marker
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
                            fontsize=7.5, color="#334155", va="center", ha="left",
                            fontfamily=font_name, fontweight="medium", zorder=6)
            continue

        # Block rectangle with category-tinted fill
        block_rect = FancyBboxPatch(
            (x_s, y_bot), bw, bh,
            boxstyle="round,pad=0,rounding_size=0.004",
            facecolor=fill_color, edgecolor=border_color,
            linewidth=1.0, zorder=3,
        )
        ax_content.add_patch(block_rect)

        # Left color accent bar
        accent_w = bw * 0.02
        accent_rect = FancyBboxPatch(
            (x_s, y_bot), max(accent_w, 0.003), bh,
            boxstyle="square,pad=0",
            facecolor=bar_color, edgecolor="none",
            zorder=4,
        )
        ax_content.add_patch(accent_rect)

        # ── Text inside block ────────────────────────────────────────
        fig_w_inches = fig.get_figwidth()
        block_w_inches = bw * fig_w_inches * content_x_width
        block_h_inches = bh * fig.get_figheight() * (content_top - content_bot)

        # Font sizing based on block width
        if block_w_inches >= 3.0:
            title_fs = 9.5
        elif block_w_inches >= 2.0:
            title_fs = 8.5
        elif block_w_inches >= 1.2:
            title_fs = 7.5
        elif block_w_inches >= 0.6:
            title_fs = 7
        else:
            title_fs = 6

        bullet_fs = max(5.5, title_fs - 1.5)

        chars_per_inch = title_fs * 0.85
        max_chars = max(5, int(block_w_inches * chars_per_inch))

        line_h = base_line_h_inches / (fig.get_figheight() * (content_top - content_bot))

        text_margin_x = bw * 0.06
        text_margin_y = bh * 0.06
        text_x = x_s + text_margin_x + max(accent_w, 0.003)
        text_y = y_top - text_margin_y

        if block_w_inches < 0.4:
            continue

        # Checkmark for completed items
        if item.status == "done":
            check_x = x_e - bw * 0.06
            check_y = y_top - text_margin_y
            ax_content.text(check_x, check_y, "\u2713", fontsize=title_fs + 3,
                            color="#10B981", va="top", ha="right",
                            fontfamily=font_name, fontweight="bold", zorder=6)

        # Label (e.g., "Delivery 1")
        if item.label:
            ax_content.text(
                text_x, text_y, item.label,
                fontsize=title_fs - 0.5, color="#475569",
                va="top", ha="left", fontfamily=font_name,
                fontweight="bold", style="italic",
                zorder=5,
            )
            text_y -= line_h * 1.2

        # Title (bold)
        wrapped_title = _wrap_text(item.title, max_chars)
        title_lines = wrapped_title.split("\n")[:3]
        display_title = "\n".join(title_lines)

        ax_content.text(
            text_x, text_y, display_title,
            fontsize=title_fs, color="#1E293B",
            va="top", ha="left", fontfamily=font_name,
            fontweight="bold", zorder=5,
            linespacing=1.2,
        )
        text_y -= len(title_lines) * line_h * 1.2

        # Bullet points from description
        bullets = _parse_bullets(item.description)
        if bullets and block_w_inches >= 0.8:
            bullet_chars = max(5, int(block_w_inches * bullet_fs * 0.85))
            for bullet in bullets:
                # Wrap long bullets
                if len(bullet) > bullet_chars:
                    bullet = bullet[:bullet_chars - 1] + "\u2026"
                bullet_text = f"\u2022 {bullet}"

                ax_content.text(
                    text_x, text_y, bullet_text,
                    fontsize=bullet_fs, color="#334155",
                    va="top", ha="left", fontfamily=font_name,
                    zorder=5, linespacing=1.15,
                )
                text_y -= line_h * 1.1

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
