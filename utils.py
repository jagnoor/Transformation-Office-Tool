"""Shared utilities used across renderers and exporters.

Centralizes color manipulation, contrast calculation, and font resolution so
each renderer doesn't reimplement them.
"""
import io
from functools import lru_cache
from typing import Optional, Tuple

import matplotlib.font_manager as fm


# ── Color manipulation ───────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert '#RRGGBB' or 'RRGGBB' to an (R, G, B) integer tuple (0..255)."""
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB integers (0..255) back to '#RRGGBB'."""
    return f"#{max(0, min(255, r)):02X}{max(0, min(255, g)):02X}{max(0, min(255, b)):02X}"


def lighten_color(hex_color: str, factor: float = 0.85) -> str:
    """Return a lighter shade of the given color.

    factor=0 returns the original color; factor=1 returns white.
    """
    r, g, b = hex_to_rgb(hex_color)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return rgb_to_hex(r, g, b)


def darken_color(hex_color: str, factor: float = 0.2) -> str:
    """Return a darker shade of the given color.

    factor=0 returns the original color; factor=1 returns black.
    """
    r, g, b = hex_to_rgb(hex_color)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return rgb_to_hex(r, g, b)


def relative_luminance(hex_color: str) -> float:
    """Return WCAG-style relative luminance (0..1) of a hex color."""
    r, g, b = hex_to_rgb(hex_color)
    # ITU-R BT.601 weighted average — good-enough proxy for WCAG luminance
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def text_color_for_bg(bg_hex: str, threshold: float = 0.5) -> str:
    """Pick white or near-black text color for best contrast on bg_hex."""
    return "#FFFFFF" if relative_luminance(bg_hex) < threshold else "#1E293B"


# ── Font resolution (cached) ─────────────────────────────────────────────────

@lru_cache(maxsize=1)
def _available_fonts() -> frozenset:
    """Cache the system font list — rebuilding it is slow."""
    return frozenset(f.name for f in fm.fontManager.ttflist)


def resolve_font(preferred: str = "Arial") -> str:
    """Resolve to a font that matplotlib can render, with sensible fallbacks."""
    available = _available_fonts()
    for candidate in (preferred, "Arial", "Helvetica", "DejaVu Sans"):
        if candidate in available:
            return candidate
    return "DejaVu Sans"


# ── Logo placement (matplotlib figures) ──────────────────────────────────────

def compute_logo_geometry(
    fig, logo_bytes: Optional[bytes],
    max_height_frac: float = 0.05, margin_frac: float = 0.012,
    max_width_frac: float = 0.18,
) -> Optional[dict]:
    """Compute placement for a logo in the top-right corner of a figure.

    Returns a dict with the opened PIL image and the figure-fraction
    rectangle to draw it in (preserving aspect ratio), or None if
    logo_bytes is empty or not a readable image.
    """
    if not logo_bytes:
        return None
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(logo_bytes))
        img.load()
        img = img.convert("RGBA")
    except Exception:
        return None

    fig_w_in, fig_h_in = fig.get_size_inches()
    img_w_px, img_h_px = img.size
    if img_w_px <= 0 or img_h_px <= 0 or fig_w_in <= 0 or fig_h_in <= 0:
        return None
    aspect = img_w_px / img_h_px

    height_frac = max_height_frac
    width_frac = height_frac * (fig_h_in / fig_w_in) * aspect
    if width_frac > max_width_frac:
        width_frac = max_width_frac
        height_frac = width_frac * (fig_w_in / fig_h_in) / aspect

    left = 1.0 - margin_frac - width_frac
    bottom = 1.0 - margin_frac - height_frac
    return {
        "img": img,
        "left": left,
        "bottom": bottom,
        "width_frac": width_frac,
        "height_frac": height_frac,
    }


def place_logo(fig, geometry: Optional[dict]) -> None:
    """Draw a logo onto a figure using geometry from compute_logo_geometry().

    No-op if geometry is None. Adds the logo as the topmost axes so it
    renders above all other chart elements.
    """
    if not geometry:
        return
    ax_logo = fig.add_axes([
        geometry["left"], geometry["bottom"],
        geometry["width_frac"], geometry["height_frac"],
    ])
    ax_logo.set_zorder(50)
    ax_logo.imshow(geometry["img"])
    ax_logo.axis("off")
