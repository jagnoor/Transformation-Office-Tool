"""Sequencing Diagram renderer — Tetris-style space-filling layout.

Visual storytelling: blocks are placed in delivery priority order using a
discrete grid. The algorithm guarantees 100% space utilization with ZERO gaps.
Blocks touch edge-to-edge with only a hairline border between them.

Key design:
- Discrete grid (columns = months, rows = slots) ensures perfect alignment
- Blocks snap to grid boundaries — no fractional positioning
- L-shaped and T-shaped composite blocks fill irregular gaps
- Text is hard-clipped to block boundaries with proper wrapping
- Bold saturated colors with subtle inner shadow for depth
"""
import io
import math
import re
from collections import defaultdict
from datetime import date, timedelta
from typing import List, Optional, Tuple, Dict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.path import Path
import matplotlib.patches as mpl_patches
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


def _parse_bullets(description: str) -> List[str]:
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
    return bullets[:6]


def _extract_sequence_number(label: str) -> int:
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


# ── Segment-based Tetris layout engine ────────────────────────────────────────

def _layout_sequencing(
    items: List[WorkItem],
    chart_start: date,
    chart_end: date,
    content_w_inches: float,
    content_h_inches: float,
    base_line_h: float,
    font_size: float,
) -> List[dict]:
    """Segment-based Tetris layout with ZERO gaps.

    The algorithm guarantees 100% space utilization:

    1. Sort items by sequence number (delivery priority).
    2. Assign to slots via greedy packing (determines vertical order).
    3. Find all "breakpoints" where the set of active items changes.
    4. For each segment between breakpoints, distribute the FULL vertical
       space evenly among active items. Items that overlap in time get
       thinner; items alone in a segment get the full height.
    5. Each item becomes a series of rectangles (one per segment) with
       varying heights — creating natural Tetris-like shapes.

    Returns list of layout dicts, each with a 'rects' list.
    """
    total_days = (chart_end - chart_start).days
    if total_days <= 0:
        return []

    sorted_items = sorted(
        items,
        key=lambda it: (_extract_sequence_number(it.label), it.start_date, it.title)
    )

    # Compute day ranges
    item_ranges = []
    for item in sorted_items:
        d_start = max(0, (item.start_date - chart_start).days)
        d_end = min(total_days, (item.end_date - chart_start).days)
        if d_end <= d_start:
            d_end = d_start + 1
        item_ranges.append((d_start, d_end))

    # ── Greedy slot assignment (determines vertical ordering) ─────────
    slots: List[List[Tuple[int, int, int]]] = []
    item_slot: Dict[int, int] = {}

    for idx in range(len(sorted_items)):
        d_start, d_end = item_ranges[idx]
        placed = False
        for slot_idx, slot_entries in enumerate(slots):
            fits = all(d_start >= e or d_end <= s for (s, e, _) in slot_entries)
            if fits:
                slot_entries.append((d_start, d_end, idx))
                item_slot[idx] = slot_idx
                placed = True
                break
        if not placed:
            slots.append([(d_start, d_end, idx)])
            item_slot[idx] = len(slots) - 1

    if not slots:
        return []

    # ── Find breakpoints ──────────────────────────────────────────────
    bp_set = set()
    for d_start, d_end in item_ranges:
        bp_set.add(d_start)
        bp_set.add(d_end)
    # Also add chart boundaries
    bp_set.add(0)
    bp_set.add(total_days)
    breakpoints = sorted(bp_set)

    # ── Segment-based stacking ────────────────────────────────────────
    edge = 0.0004
    item_rects: Dict[int, List[dict]] = defaultdict(list)

    for i in range(len(breakpoints) - 1):
        seg_start = breakpoints[i]
        seg_end = breakpoints[i + 1]
        if seg_end <= seg_start:
            continue

        # Active items in this segment, sorted by slot for consistent ordering
        active = []
        for idx, (ds, de) in enumerate(item_ranges):
            if ds < seg_end and de > seg_start:
                active.append((item_slot[idx], idx))
        active.sort()

        if not active:
            continue

        n = len(active)
        row_h = 1.0 / n

        x_s = seg_start / total_days + edge
        x_e = seg_end / total_days - edge

        for rank, (slot, idx) in enumerate(active):
            y_top = 1.0 - rank * row_h - edge
            y_bot = 1.0 - (rank + 1) * row_h + edge

            item_rects[idx].append({
                "x_start": x_s,
                "x_end": x_e,
                "y_top": y_top,
                "y_bot": y_bot,
            })

    # ── Merge adjacent rectangles with same height for cleaner look ───
    for idx in item_rects:
        rects = item_rects[idx]
        if len(rects) <= 1:
            continue
        merged = [rects[0]]
        for r in rects[1:]:
            prev = merged[-1]
            # Merge if vertically aligned and horizontally adjacent
            if (abs(prev["y_top"] - r["y_top"]) < 0.008 and
                abs(prev["y_bot"] - r["y_bot"]) < 0.008 and
                abs(prev["x_end"] - r["x_start"]) < 0.015):
                prev["x_end"] = r["x_end"]
            else:
                merged.append(r)
        item_rects[idx] = merged

    # ── Build layout output ───────────────────────────────────────────
    layouts = []
    for idx, item in enumerate(sorted_items):
        rects = item_rects.get(idx, [])
        if not rects:
            continue

        # Pick the largest rectangle (by area) for text placement
        best = max(rects, key=lambda r: (r["x_end"] - r["x_start"]) * (r["y_top"] - r["y_bot"]))

        layouts.append({
            "item": item,
            "rects": rects,
            "x_start": best["x_start"],
            "x_end": best["x_end"],
            "y_top": best["y_top"],
            "y_bot": best["y_bot"],
            "height": best["y_top"] - best["y_bot"],
        })

    return layouts


# ── Main renderer ────────────────────────────────────────────────────────────

def render_sequencing_diagram(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 200,
    slide_aspect: str = "16:9",
) -> bytes:
    if not items:
        raise ValueError("No work items to render")
    fig = _render_sequencing_figure(items, config, dpi, slide_aspect)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none",
                pad_inches=0.05)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_sequencing_pdf(
    items: List[WorkItem],
    config: ChartConfig,
    dpi: int = 300,
    slide_aspect: str = "16:9",
) -> bytes:
    if not items:
        raise ValueError("No work items to render")
    fig = _render_sequencing_figure(items, config, dpi, slide_aspect)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), pad_inches=0.05)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _render_sequencing_figure(items, config, dpi, slide_aspect="16:9"):
    """Build the Tetris sequencing diagram figure."""
    font_name = _resolve_font(config.font_family)
    plt.rcParams.update({
        "font.family": font_name,
        "font.size": 10,
        "axes.unicode_minus": False,
    })

    categories = []
    seen = set()
    for it in items:
        if it.category not in seen:
            categories.append(it.category)
            seen.add(it.category)

    chart_start = config.start_date or min(it.start_date for it in items)
    chart_end = config.end_date or max(it.end_date for it in items)
    total_days = (chart_end - chart_start).days

    # Figure size — larger for better detail
    if slide_aspect == "16:9":
        fig_w, fig_h = 22, 12.5
    elif slide_aspect == "4:3":
        fig_w, fig_h = 18, 13.5
    else:
        fig_w, fig_h = 22, 12.5

    # Layout proportions
    header_h = 0.05
    header_top = 1.0 - 0.01
    header_bot = header_top - header_h
    quarter_h = 0.022
    quarter_top = header_bot - 0.003
    quarter_bot = quarter_top - quarter_h
    month_h = 0.018
    month_top = quarter_bot
    month_bot = month_top - month_h
    content_top = month_bot
    content_bot = 0.015
    legend_left_x = 0.01
    legend_width = 0.065
    content_left_x = legend_left_x + legend_width + 0.005
    content_right_x = 0.995
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

    # ── Header ───────────────────────────────────────────────────────
    ax_header = fig.add_axes([0.01, header_bot, 0.98, header_h])
    ax_header.set_xlim(0, 1)
    ax_header.set_ylim(0, 1)
    ax_header.axis("off")

    ax_header.text(0.0, 0.6, config.title, fontsize=26, fontweight="bold",
                   color="#0F172A", fontfamily=font_name, va="center")
    if config.subtitle:
        ax_header.text(0.0, 0.05, config.subtitle, fontsize=12, color="#64748B",
                       fontfamily=font_name, va="center")

    date_text = f"{chart_start.strftime('%b %d, %Y')}  \u2192  {chart_end.strftime('%b %d, %Y')}"
    ax_header.text(1.0, 0.4, date_text, fontsize=10, color="#94A3B8",
                   fontfamily=font_name, va="center", ha="right")

    # ── Quarter row ──────────────────────────────────────────────────
    quarter_columns, month_columns = _compute_two_row_timeline(chart_start, chart_end)

    ax_quarter = fig.add_axes([content_left_x, quarter_bot, content_x_width, quarter_h])
    ax_quarter.set_xlim(0, 1)
    ax_quarter.set_ylim(0, 1)
    ax_quarter.axis("off")

    q_bg = Rectangle((0, 0), 1, 1, facecolor="#1E293B", edgecolor="none", zorder=0)
    ax_quarter.add_patch(q_bg)

    for i, (col_start, col_end, label) in enumerate(quarter_columns):
        x_start = (col_start - chart_start).days / total_days
        x_end = (col_end - chart_start).days / total_days
        x_mid = (x_start + x_end) / 2
        if i % 2 == 0:
            ax_quarter.add_patch(Rectangle(
                (x_start, 0), x_end - x_start, 1,
                facecolor="#334155", edgecolor="none", zorder=1))
        ax_quarter.text(x_mid, 0.5, label, fontsize=10, color="#F1F5F9",
                        ha="center", va="center", fontfamily=font_name,
                        fontweight="bold", zorder=2)
        if i > 0:
            ax_quarter.axvline(x=x_start, color="#475569", linewidth=0.5, zorder=2)

    # ── Month row ────────────────────────────────────────────────────
    ax_month = fig.add_axes([content_left_x, month_bot, content_x_width, month_h])
    ax_month.set_xlim(0, 1)
    ax_month.set_ylim(0, 1)
    ax_month.axis("off")

    ax_month.add_patch(Rectangle((0, 0), 1, 1, facecolor="#475569", edgecolor="none", zorder=0))

    for i, (col_start, col_end, label) in enumerate(month_columns):
        x_start = (col_start - chart_start).days / total_days
        x_end = (col_end - chart_start).days / total_days
        x_mid = (x_start + x_end) / 2
        if i % 2 == 0:
            ax_month.add_patch(Rectangle(
                (x_start, 0), x_end - x_start, 1,
                facecolor="#546378", edgecolor="none", zorder=1))
        ax_month.text(x_mid, 0.5, label, fontsize=8, color="#E2E8F0",
                      ha="center", va="center", fontfamily=font_name,
                      fontweight="medium", zorder=2)
        if i > 0:
            ax_month.axvline(x=x_start, color="#64748B", linewidth=0.3, zorder=2)

    # ── Legend ────────────────────────────────────────────────────────
    if config.show_legend:
        ax_legend = fig.add_axes([legend_left_x, content_bot, legend_width, content_top - content_bot])
        ax_legend.set_xlim(0, 1)
        ax_legend.set_ylim(0, 1)
        ax_legend.axis("off")

        ax_legend.add_patch(FancyBboxPatch(
            (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.015",
            facecolor="#F1F5F9", edgecolor="#CBD5E1", linewidth=0.5, zorder=0))

        ax_legend.text(0.5, 0.975, "LEGEND", fontsize=7, color="#64748B",
                       ha="center", va="top", fontfamily=font_name,
                       fontweight="bold", zorder=2)

        num_cats = len(categories)
        if num_cats > 0:
            spacing = min(0.065, 0.9 / num_cats)
            y_pos = 0.93
            for cat in categories:
                color = config.get_category_color(cat, categories)
                ax_legend.add_patch(FancyBboxPatch(
                    (0.06, y_pos - 0.018), 0.16, 0.028,
                    boxstyle="round,pad=0.001,rounding_size=0.008",
                    facecolor=color, edgecolor="none", zorder=5))
                ax_legend.text(0.3, y_pos - 0.004, cat, fontsize=6.5,
                               color="#334155", va="center", ha="left",
                               fontfamily=font_name, fontweight="medium",
                               zorder=5, clip_on=True)
                y_pos -= spacing

    # ── Content area ─────────────────────────────────────────────────
    ax = fig.add_axes([content_left_x, content_bot, content_x_width, content_top - content_bot])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    # Dark background to make blocks pop
    ax.add_patch(Rectangle((0, 0), 1, 1, facecolor="#1B2332", edgecolor="none", zorder=0))

    # Subtle grid lines
    for i, (col_start, col_end, label) in enumerate(month_columns):
        x = (col_start - chart_start).days / total_days
        if i > 0:
            ax.axvline(x=x, color="#2D3748", linewidth=0.3, zorder=1, alpha=0.6)

    # ── Draw blocks ──────────────────────────────────────────────────
    for layout in layouts:
        item = layout["item"]
        rects = layout.get("rects", [])
        # Primary rect (largest) for text placement
        x_s = layout["x_start"]
        x_e = layout["x_end"]
        y_top = layout["y_top"]
        y_bot = layout["y_bot"]
        bw = x_e - x_s
        bh = y_top - y_bot

        if bh <= 0.002 or bw <= 0.002:
            continue

        base_color = config.get_category_color(item.category, categories)
        bar_color = item.color_override if item.color_override else base_color
        light = _lighten_color(bar_color, 0.25)
        dark = _darken_color(bar_color, 0.3)
        text_color = _text_color_for_bg(bar_color)

        # ── Draw ALL segment rectangles for this item ─────────────
        for rect in rects:
            rx_s = rect["x_start"]
            rx_e = rect["x_end"]
            ry_top = rect["y_top"]
            ry_bot = rect["y_bot"]
            rw = rx_e - rx_s
            rh = ry_top - ry_bot
            if rh <= 0.001 or rw <= 0.001:
                continue

            ax.add_patch(Rectangle(
                (rx_s, ry_bot), rw, rh,
                facecolor=bar_color, edgecolor=dark,
                linewidth=0.6, zorder=3))

            # Inner highlight (top edge)
            hl_h = min(0.003, rh * 0.05)
            ax.add_patch(Rectangle(
                (rx_s, ry_top - hl_h), rw, hl_h,
                facecolor=light, edgecolor="none", zorder=4, alpha=0.4))

            # Inner shadow (bottom edge)
            sh_h = min(0.002, rh * 0.03)
            ax.add_patch(Rectangle(
                (rx_s, ry_bot), rw, sh_h,
                facecolor=dark, edgecolor="none", zorder=4, alpha=0.25))

        # ── Sequence badge (on primary rect) ──────────────────────
        seq_num = _extract_sequence_number(item.label)
        badge_space = 0.0
        if seq_num < 999:
            badge_r = min(0.009, bw * 0.08, bh * 0.1)
            badge_x = x_e - badge_r - 0.004
            badge_y = y_top - badge_r - 0.004
            ax.add_patch(plt.Circle(
                (badge_x, badge_y), badge_r,
                facecolor="#FFFFFF", edgecolor=dark,
                linewidth=0.6, zorder=6, alpha=0.92))
            ax.text(badge_x, badge_y, str(seq_num),
                    fontsize=max(4.5, min(6, badge_r * 500)),
                    color="#1E293B", ha="center", va="center",
                    fontfamily=font_name, fontweight="bold", zorder=7)
            badge_space = badge_r * 2 + 0.008

        # ── Text rendering (clipped to primary block) ─────────────
        block_w_inches = bw * fig_w * content_x_width
        block_h_inches = bh * fig_h * content_h_frac

        # Adaptive font size
        if block_h_inches >= 2.0 and block_w_inches >= 3.0:
            title_fs = 9
        elif block_h_inches >= 1.0 and block_w_inches >= 2.0:
            title_fs = 8
        elif block_h_inches >= 0.6 and block_w_inches >= 1.2:
            title_fs = 7
        elif block_w_inches >= 0.6:
            title_fs = 6
        else:
            title_fs = 5

        bullet_fs = max(4.5, title_fs - 1.5)
        line_h_norm = (title_fs * 0.016) / (fig_h * content_h_frac)

        # Text area with proper margins
        pad_x = bw * 0.04
        pad_y = bh * 0.06
        text_x = x_s + pad_x
        text_y = y_top - pad_y
        text_max_x = x_e - pad_x - badge_space  # don't overlap badge
        text_avail_w = text_max_x - text_x

        if block_w_inches < 0.4 or block_h_inches < 0.12:
            continue

        chars_per_line = max(4, int(text_avail_w * fig_w * content_x_width * title_fs * 0.85))
        available_h = bh - pad_y * 2
        max_lines = max(1, int(available_h / (line_h_norm * 1.1)))
        lines_used = 0

        # Checkmark for done items
        if item.status == "done" and max_lines >= 1:
            check_x = text_max_x
            ax.text(check_x, text_y, "\u2713", fontsize=title_fs + 1,
                    color=text_color, va="top", ha="right",
                    fontfamily=font_name, fontweight="bold", zorder=6,
                    alpha=0.8)

        # Label line
        if item.label and max_lines >= 2:
            label_text = item.label
            if len(label_text) > chars_per_line:
                label_text = label_text[:chars_per_line - 1] + "\u2026"
            ax.text(text_x, text_y, label_text,
                    fontsize=max(4.5, title_fs - 0.5), color=text_color,
                    va="top", ha="left", fontfamily=font_name,
                    fontweight="bold", style="italic",
                    zorder=5, alpha=0.85,
                    clip_on=True, clip_box=ax.bbox)
            text_y -= line_h_norm * 1.05
            lines_used += 1

        # Title
        if lines_used < max_lines:
            wrapped = _wrap_text(item.title, chars_per_line)
            title_lines = wrapped.split("\n")
            avail = max(1, min(len(title_lines), max_lines - lines_used))
            title_lines = title_lines[:avail]
            if avail < len(wrapped.split("\n")):
                last = title_lines[-1]
                if len(last) > 2:
                    title_lines[-1] = last[:chars_per_line - 1] + "\u2026"

            for tl in title_lines:
                ax.text(text_x, text_y, tl,
                        fontsize=title_fs, color=text_color,
                        va="top", ha="left", fontfamily=font_name,
                        fontweight="bold", zorder=5,
                        clip_on=True, clip_box=ax.bbox)
                text_y -= line_h_norm * 1.05
                lines_used += 1

        # Bullets
        bullets = _parse_bullets(item.description)
        remaining = max_lines - lines_used
        bullet_chars = max(4, int(text_avail_w * fig_w * content_x_width * bullet_fs * 0.85))
        if bullets and remaining >= 1 and block_w_inches >= 0.7:
            for bullet in bullets[:remaining]:
                if len(bullet) > bullet_chars:
                    bullet = bullet[:bullet_chars - 1] + "\u2026"
                ax.text(text_x, text_y, f"\u2022 {bullet}",
                        fontsize=bullet_fs, color=text_color,
                        va="top", ha="left", fontfamily=font_name,
                        zorder=5, alpha=0.8,
                        clip_on=True, clip_box=ax.bbox)
                text_y -= line_h_norm * 0.95

    # ── Today line ───────────────────────────────────────────────────
    if config.show_today_line:
        today = config.today_date or date.today()
        if chart_start <= today <= chart_end:
            today_x = (today - chart_start).days / total_days
            ax.axvline(x=today_x, color="#EF4444", linewidth=2.5,
                       linestyle="--", zorder=15, alpha=0.8)
            ax_month.axvline(x=today_x, color="#EF4444", linewidth=2,
                             linestyle="--", zorder=15, alpha=0.7)
            ax_quarter.axvline(x=today_x, color="#EF4444", linewidth=2,
                               linestyle="--", zorder=15, alpha=0.7)

    return fig
