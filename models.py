"""Data models for the Transformation Office Block & Gantt Creator Tool."""
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional


# ── Modern color palettes ────────────────────────────────────────────────────
PALETTES = {
    "Ocean": [
        "#0077B6", "#00B4D8", "#0096C7", "#023E8A", "#48CAE4",
        "#90E0EF", "#CAF0F8", "#03045E", "#ADE8F4", "#0077B6",
    ],
    "Sunset": [
        "#F94144", "#F3722C", "#F8961E", "#F9844A", "#F9C74F",
        "#90BE6D", "#43AA8B", "#4D908E", "#577590", "#277DA1",
    ],
    "Forest": [
        "#2D6A4F", "#40916C", "#52B788", "#74C69D", "#95D5B2",
        "#1B4332", "#B7E4C7", "#D8F3DC", "#081C15", "#3A5A40",
    ],
    "Corporate": [
        "#2563EB", "#7C3AED", "#0891B2", "#059669", "#D97706",
        "#DC2626", "#4F46E5", "#0D9488", "#CA8A04", "#9333EA",
    ],
    "Vibrant": [
        "#6366F1", "#EC4899", "#F59E0B", "#10B981", "#3B82F6",
        "#8B5CF6", "#EF4444", "#14B8A6", "#F97316", "#06B6D4",
    ],
    "Monochrome": [
        "#1E293B", "#334155", "#475569", "#64748B", "#94A3B8",
        "#CBD5E1", "#0F172A", "#374151", "#6B7280", "#9CA3AF",
    ],
}

DEFAULT_PALETTE = "Corporate"

STATUS_COLORS = {
    "planned": "#94A3B8",
    "in_progress": "#3B82F6",
    "done": "#10B981",
    "at_risk": "#EF4444",
}

STATUS_LABELS = {
    "planned": "Planned",
    "in_progress": "In Progress",
    "done": "Done",
    "at_risk": "At Risk",
}


@dataclass
class WorkItem:
    """A single work item (task/deliverable/initiative)."""
    title: str
    start_date: date
    end_date: date
    category: str = "General"
    description: str = ""
    status: str = "planned"
    owner: str = ""
    label: str = ""
    is_milestone: bool = False
    color_override: str = ""

    def __post_init__(self):
        if isinstance(self.start_date, datetime):
            self.start_date = self.start_date.date()
        if isinstance(self.end_date, datetime):
            self.end_date = self.end_date.date()
        if self.start_date == self.end_date:
            self.is_milestone = True
        if not self.status:
            self.status = "planned"
        self.status = str(self.status).strip().lower().replace(" ", "_")

    @property
    def duration_days(self) -> int:
        return max(1, (self.end_date - self.start_date).days)


@dataclass
class ChartConfig:
    """Configuration for chart rendering."""
    title: str = "Project Roadmap"
    subtitle: str = ""
    palette_name: str = DEFAULT_PALETTE
    show_today_line: bool = True
    today_date: Optional[date] = None
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    show_legend: bool = True
    show_status: bool = True
    font_family: str = "Arial"
    background_color: str = "#FFFFFF"
    slide_size: str = "WIDE"  # WIDE (16:9), A4, A3

    @property
    def palette(self):
        return PALETTES.get(self.palette_name, PALETTES[DEFAULT_PALETTE])

    def get_category_color(self, category: str, categories: list) -> str:
        """Get color for a category based on its index in the list."""
        if category in categories:
            idx = categories.index(category)
        else:
            idx = 0
        palette = self.palette
        return palette[idx % len(palette)]
