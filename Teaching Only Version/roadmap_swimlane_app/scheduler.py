from __future__ import annotations


# =============================================================================
# Teaching notes: scheduler.py (overlap stacking)
#
# This module solves the key visual problem:
# "What if two tasks in the same workstream overlap in time?"
#
# Answer:
#   - We create sublanes inside each workstream.
#   - We assign each task to the first sublane that is "free".
#
# This is a classic greedy interval-partitioning algorithm.
# Deterministic = same input always results in the same layout.
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
from collections.abc import Iterable
from datetime import date
from typing import Dict, List, Tuple

from roadmap_models import Task


def assign_sublanes(tasks: Iterable[Task], *, touching_counts_as_overlap: bool = True) -> List[Task]:
    """
    Deterministic interval partitioning per workstream.

    Rules:
    - Intervals are treated as inclusive on both ends for overlap purposes.
    - If touching_counts_as_overlap is True (default), then two tasks where
      previous.end_date == next.start_date are considered overlapping and will
      be placed on different sublanes.
      Availability condition: last_end_date < next.start_date
    - If touching_counts_as_overlap is False, then touching endpoints are allowed
      in the same sublane.
      Availability condition: last_end_date <= next.start_date

    The algorithm is greedy: each task is assigned to the first available sublane
    (lowest index) to keep output stable and predictable.
    """
    tasks_sorted = sorted(tasks, key=lambda t: (t.start_date, t.end_date, t.id))
    lane_end_dates: List[date] = []
    out: List[Task] = []

    for t in tasks_sorted:
        assigned_lane = None
        for lane_idx, last_end in enumerate(lane_end_dates):
            if touching_counts_as_overlap:
                ok = last_end < t.start_date
            else:
                ok = last_end <= t.start_date
            if ok:
                assigned_lane = lane_idx
                lane_end_dates[lane_idx] = t.end_date
                break

        if assigned_lane is None:
            assigned_lane = len(lane_end_dates)
            lane_end_dates.append(t.end_date)

        out.append(t.model_copy(update={"sublane": assigned_lane}))

    return out


def schedule_by_workstream(
    tasks: Iterable[Task],
    *,
    touching_counts_as_overlap: bool = True,
) -> Dict[str, List[Task]]:
    """
    Groups tasks by workstream and assigns sublanes within each workstream.

    Returns a dict workstream -> list[Task] with sublane populated.
    """
    grouped: Dict[str, List[Task]] = {}
    for t in tasks:
        grouped.setdefault(t.workstream, []).append(t)

    scheduled: Dict[str, List[Task]] = {}
    for ws, ws_tasks in grouped.items():
        scheduled[ws] = assign_sublanes(ws_tasks, touching_counts_as_overlap=touching_counts_as_overlap)

    return scheduled


def validate_no_overlaps_per_lane(tasks: Iterable[Task], *, touching_counts_as_overlap: bool = True) -> Tuple[bool, str]:
    """
    Utility for tests/debug: confirms no overlaps exist within any (workstream, sublane).

    Returns (ok, message).
    """
    by_lane: Dict[tuple[str, int], List[Task]] = {}
    for t in tasks:
        if t.sublane is None:
            return False, f"Task {t.id} has no sublane assigned."
        by_lane.setdefault((t.workstream, int(t.sublane)), []).append(t)

    for (ws, lane), lane_tasks in by_lane.items():
        lane_tasks_sorted = sorted(lane_tasks, key=lambda t: (t.start_date, t.end_date, t.id))
        prev: Task | None = None
        for cur in lane_tasks_sorted:
            if prev is None:
                prev = cur
                continue
            if touching_counts_as_overlap:
                overlap = prev.end_date >= cur.start_date
            else:
                overlap = prev.end_date > cur.start_date
            if overlap:
                return False, f"Overlap detected in workstream={ws}, sublane={lane}: {prev.id} vs {cur.id}"
            prev = cur

    return True, "ok"
