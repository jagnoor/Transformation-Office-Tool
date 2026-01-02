from __future__ import annotations


# =============================================================================
# Teaching notes: tests/test_export_smoke.py
#
# A small export test to make sure render/export functions produce non-empty
# outputs with correct file signatures (PNG header, PDF header, etc.).
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
from datetime import date

from roadmap_models import Settings, Task, Workstream
from scheduler import schedule_by_workstream
from export import export_pdf_bytes, export_png_bytes, export_pptx_bytes, preview_png_bytes


def test_exports_produce_bytes() -> None:
    settings = Settings(
        chart_title="Smoke Test Roadmap",
        chart_subtitle="",
        confidentiality_label="",
        overall_start_date=date(2025, 1, 1),
        overall_end_date=date(2025, 3, 31),
        timezone="America/Chicago",
        week_start_day="Mon",
        time_granularity="Weekly",
        output_dpi=300,
        show_today_line=False,
        today_line_date=None,
        page_size="A4",
        font_family="DejaVu Sans",
    )

    workstreams = [
        Workstream(workstream="A", order=1, color="#1F77B4"),
        Workstream(workstream="B", order=2, color="#FF7F0E"),
    ]

    tasks = [
        Task(
            id="T1",
            workstream="A",
            title="Overlapping 1",
            description="",
            start_date=date(2025, 1, 5),
            end_date=date(2025, 1, 20),
            status="in_progress",
            type="block",
        ),
        Task(
            id="T2",
            workstream="A",
            title="Overlapping 2",
            description="",
            start_date=date(2025, 1, 10),
            end_date=date(2025, 1, 15),
            status="planned",
            type="block",
        ),
        Task(
            id="M1",
            workstream="B",
            title="Milestone",
            description=None,
            start_date=date(2025, 2, 1),
            end_date=date(2025, 2, 1),
            status="planned",
            type="milestone",
        ),
    ]

    scheduled = schedule_by_workstream(tasks, touching_counts_as_overlap=True)
    by_id = {t.id: t for ws_tasks in scheduled.values() for t in ws_tasks}
    tasks = [by_id[t.id] for t in tasks]

    # Preview
    preview = preview_png_bytes(settings, workstreams, tasks, include_out_of_range=False, dpi=140)
    assert isinstance(preview, (bytes, bytearray))
    assert preview[:8] == b"\x89PNG\r\n\x1a\n"
    assert len(preview) > 10_000

    # High-res PNG
    png = export_png_bytes(settings, workstreams, tasks, include_out_of_range=False, dpi=300)
    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    assert len(png) > 25_000

    # Vector PDF
    pdf = export_pdf_bytes(settings, workstreams, tasks, include_out_of_range=False)
    assert pdf[:4] == b"%PDF"
    assert len(pdf) > 10_000

    # Editable PPTX
    pptx = export_pptx_bytes(settings, workstreams, tasks, include_out_of_range=False)
    assert pptx[:2] == b"PK"  # zip container
    assert len(pptx) > 20_000

    # Can be opened by python-pptx (basic validity)
    from io import BytesIO
    from pptx import Presentation

    prs = Presentation(BytesIO(pptx))
    assert len(prs.slides) == 1
