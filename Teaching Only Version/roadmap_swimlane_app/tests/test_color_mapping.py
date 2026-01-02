# =============================================================================
# Teaching notes: tests/test_color_mapping.py
#
# Tests the color dropdown name mapping (English names → hex codes).
# Also covers backwards compatibility with older template values.
# =============================================================================

from datetime import date

from roadmap_models import Task, Workstream


def test_workstream_color_friendly_names_map_to_hex() -> None:
    ws = Workstream(workstream="WS", order=1, color="Blue")
    assert ws.color == "#1F77B4"

    ws2 = Workstream(workstream="WS2", order=2, color="Sky Blue")
    assert ws2.color == "#AEC7E8"


def test_workstream_color_auto_becomes_none() -> None:
    ws = Workstream(workstream="WS", order=1, color="Auto")
    assert ws.color is None


def test_task_color_override_accepts_friendly_names() -> None:
    t = Task(
        id="T1",
        workstream="WS",
        title="T",
        start_date=date(2025, 1, 1),
        end_date=date(2025, 1, 1),
        color_override="Orange",
    )
    assert t.color_override == "#FF7F0E"


def test_legacy_color_names_still_work() -> None:
    # Backwards compatibility for older templates
    ws = Workstream(workstream="WS", order=1, color="Primary")
    assert ws.color == "#1F77B4"
