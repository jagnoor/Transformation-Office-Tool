from datetime import date

from roadmap_models import Task
from scheduler import assign_sublanes, validate_no_overlaps_per_lane


def test_interval_partitioning_no_overlaps_touching_is_overlap():
    tasks = [
        Task(id="A", workstream="WS", title="A", start_date=date(2025, 1, 1), end_date=date(2025, 1, 10)),
        Task(id="B", workstream="WS", title="B", start_date=date(2025, 1, 5), end_date=date(2025, 1, 6)),
        Task(id="C", workstream="WS", title="C", start_date=date(2025, 1, 11), end_date=date(2025, 1, 12)),
        Task(id="D", workstream="WS", title="D", start_date=date(2025, 1, 10), end_date=date(2025, 1, 10)),
    ]
    scheduled = assign_sublanes(tasks, touching_counts_as_overlap=True)

    ok, msg = validate_no_overlaps_per_lane(scheduled, touching_counts_as_overlap=True)
    assert ok, msg

    # Deterministic expectations:
    # A goes lane 0
    # B overlaps A -> lane 1
    # D touches A at 1/10 -> considered overlap -> lane 1 (since lane 0 still "busy" until 1/10)
    # C starts after A ends -> can reuse lane 0
    by_id = {t.id: t for t in scheduled}
    assert by_id["A"].sublane == 0
    assert by_id["B"].sublane == 1
    assert by_id["D"].sublane == 1
    assert by_id["C"].sublane == 0


def test_interval_partitioning_touching_allowed():
    tasks = [
        Task(id="A", workstream="WS", title="A", start_date=date(2025, 1, 1), end_date=date(2025, 1, 10)),
        Task(id="D", workstream="WS", title="D", start_date=date(2025, 1, 10), end_date=date(2025, 1, 10)),
    ]
    scheduled = assign_sublanes(tasks, touching_counts_as_overlap=False)

    ok, msg = validate_no_overlaps_per_lane(scheduled, touching_counts_as_overlap=False)
    assert ok, msg

    by_id = {t.id: t for t in scheduled}
    assert by_id["A"].sublane == 0
    assert by_id["D"].sublane == 0  # can share lane when touching is allowed
