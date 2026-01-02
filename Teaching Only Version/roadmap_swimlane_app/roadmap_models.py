from __future__ import annotations


# =============================================================================
# Teaching notes: roadmap_models.py (data validation models)
#
# This module defines the Pydantic models:
#   - Settings: global chart settings
#   - Workstream: swimlane definition
#   - Task: a block or milestone on the timeline
#
# Pydantic is used here to:
#   - validate required fields (title, dates, etc.)
#   - normalize input (trim whitespace, map color names to hex, etc.)
#   - provide clear error messages when input is messy
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
from datetime import date
from typing import Literal, Optional

from pydantic import BaseModel, Field, field_validator, model_validator

HEX_COLOR_RE = r"^#?[0-9A-Fa-f]{6}$"


# Friendly, non-technical color names used in Excel/app dropdowns.
#
# The dropdown offers 12 accessible, common colors (plus "Auto").
#
# Backwards-compatibility:
# - Older templates may include Primary/Secondary/Tertiary/Accent/Neutral.
# - Power users may still provide a hex like #1F77B4.

# What the UI/Excel template shows.
COLOR_DROPDOWN_VALUES = [
    "Auto",
    "Blue",
    "Orange",
    "Green",
    "Red",
    "Purple",
    "Brown",
    "Pink",
    "Gray",
    "Olive",
    "Cyan",
    "Sky Blue",
    "Peach",
]

# Canonical mapping for dropdown values.
_DROPDOWN_COLOR_TO_HEX = {
    "auto": None,
    "blue": "#1F77B4",
    "orange": "#FF7F0E",
    "green": "#2CA02C",
    "red": "#D62728",
    "purple": "#9467BD",
    "brown": "#8C564B",
    "pink": "#E377C2",
    "gray": "#7F7F7F",
    "olive": "#BCBD22",
    "cyan": "#17BECF",
    "sky blue": "#AEC7E8",
    "peach": "#FFBB78",
}

# Backwards-compatible synonyms from earlier builds/templates.
_LEGACY_SYNONYMS_TO_HEX = {
    "primary": "#1F77B4",
    "secondary": "#FF7F0E",
    "tertiary": "#2CA02C",
    "accent 1": "#9467BD",
    "accent 2": "#D62728",
    "accent 3": "#17BECF",
    "neutral": "#7F7F7F",
}

# Normalized token -> hex.
COLOR_NAME_TO_HEX = {**_DROPDOWN_COLOR_TO_HEX, **_LEGACY_SYNONYMS_TO_HEX}


def _normalize_color_token(value: str) -> str:
    s = (value or "").strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    return s


class Settings(BaseModel):
    chart_title: str = Field(default="Roadmap")
    chart_subtitle: Optional[str] = Field(default=None)
    confidentiality_label: Optional[str] = Field(default=None)

    overall_start_date: date
    overall_end_date: date

    timezone: str = Field(default="America/Chicago")
    week_start_day: Literal["Mon", "Sun"] = Field(default="Mon")
    time_granularity: Literal["Weekly", "Monthly"] = Field(default="Weekly")

    output_dpi: Literal[150, 300, 600] = Field(default=300)

    show_today_line: bool = Field(default=True)
    today_line_date: Optional[date] = Field(default=None)

    page_size: Literal["A3", "A4"] = Field(default="A3")
    font_family: str = Field(default="Calibri")  # Falls back at render-time if not found.

    @field_validator("chart_title")
    @classmethod
    def _title_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("chart_title is required.")
        return v

    @field_validator("timezone")
    @classmethod
    def _tz_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("timezone is required.")
        return v

    @field_validator("font_family")
    @classmethod
    def _font_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            return "Calibri"
        return v

    @model_validator(mode="after")
    def _date_range_valid(self) -> "Settings":
        if self.overall_end_date < self.overall_start_date:
            raise ValueError("overall_end_date must be on/after overall_start_date.")
        return self


class Workstream(BaseModel):
    workstream: str
    order: Optional[int] = None
    color: Optional[str] = None  # normalized to "#RRGGBB" when present

    @field_validator("workstream")
    @classmethod
    def _ws_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("workstream is required.")
        return v

    @field_validator("color")
    @classmethod
    def _normalize_workstream_color(cls, v: Optional[str]) -> Optional[str]:
        """Accepts either a friendly name (Primary/Secondary/...) or a hex value."""

        if v is None:
            return None
        v = str(v).strip()
        if not v:
            return None

        token = _normalize_color_token(v)
        if token in COLOR_NAME_TO_HEX:
            return COLOR_NAME_TO_HEX[token]

        # Fallback: hex
        import re

        if not re.match(HEX_COLOR_RE, v):
            allowed = ", ".join(COLOR_DROPDOWN_VALUES)
            raise ValueError(f"color must be one of: {allowed} (or a hex like #1F77B4).")
        if not v.startswith("#"):
            v = "#" + v
        return v.upper()


TaskStatus = Literal["planned", "in_progress", "done", "risk"]
TaskType = Literal["block", "milestone"]


class Task(BaseModel):
    id: str
    workstream: str
    title: str
    description: Optional[str] = None
    start_date: date
    end_date: date

    status: Optional[TaskStatus] = None
    owner: Optional[str] = None
    color_override: Optional[str] = None
    type: TaskType = "block"
    hyperlink: Optional[str] = None

    # Computed during scheduling:
    sublane: Optional[int] = None

    @field_validator("id")
    @classmethod
    def _id_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("id is required (blank IDs should have been auto-generated).")
        return v

    @field_validator("workstream")
    @classmethod
    def _task_ws_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("workstream is required.")
        return v

    @field_validator("title")
    @classmethod
    def _title_not_empty(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("title is required.")
        return v

    @field_validator("description")
    @classmethod
    def _desc_strip(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        v = str(v).strip()
        return v or None

    @field_validator("owner")
    @classmethod
    def _owner_strip(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        v = str(v).strip()
        return v or None

    @field_validator("color_override")
    @classmethod
    def _normalize_color_override(cls, v: Optional[str]) -> Optional[str]:
        """Optional. Accept friendly names or hex; if blank -> None."""
        if v is None:
            return None
        v = str(v).strip()
        if not v:
            return None

        token = _normalize_color_token(v)
        if token in COLOR_NAME_TO_HEX:
            # 'auto' maps to None, which is valid for an optional override.
            return COLOR_NAME_TO_HEX[token]

        import re

        if not re.match(HEX_COLOR_RE, v):
            allowed = ", ".join(COLOR_DROPDOWN_VALUES)
            raise ValueError(f"color_override must be one of: {allowed} (or a hex like #1F77B4).")
        if not v.startswith("#"):
            v = "#" + v
        return v.upper()

    @model_validator(mode="after")
    def _dates_valid(self) -> "Task":
        if self.end_date < self.start_date:
            raise ValueError("end_date must be on/after start_date.")
        return self
