from __future__ import annotations


# =============================================================================
# Teaching notes: scripts/smoke_test.py
#
# This script is a quick "does everything basically work?" check.
# It generates random roadmaps, round-trips them through Excel, and exports
# PNG/PDF/PPTX multiple times.
#
# If you change rendering or parsing logic, run this script to catch regressions.
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
import sys
import random
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import List

import pandas as pd

# Allow running this file directly via: python scripts/smoke_test.py
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_io import TASK_COLUMNS, WORKSTREAM_COLUMNS, read_roadmap_excel, write_roadmap_excel_bytes
from export import export_pdf_bytes, export_png_bytes, export_pptx_bytes, preview_png_bytes
from roadmap_models import Settings, Task, Workstream
from scheduler import schedule_by_workstream


@dataclass
class SmokeResult:
    iteration: int
    workstreams: int
    tasks: int
    preview_png_bytes: int
    png_bytes: int
    pdf_bytes: int
    pptx_bytes: int


def _random_hex() -> str:
    return "#" + "".join(random.choice("0123456789ABCDEF") for _ in range(6))


def build_random_case(iteration: int) -> tuple[Settings, List[Workstream], List[Task]]:
    random.seed(1000 + iteration)

    overall_start = date(2026, 1, 1)
    overall_end = date(2026, 12, 31)

    settings = Settings(
        chart_title=f"Smoke Roadmap {iteration}",
        chart_subtitle="Generated data",
        confidentiality_label="Dayforce Confidential",
        overall_start_date=overall_start,
        overall_end_date=overall_end,
        timezone="America/Chicago",
        week_start_day="Mon",
        time_granularity="Weekly",
        output_dpi=300,
        show_today_line=False,
        page_size="A3",
        font_family="DejaVu Sans",
    )

    ws_names = ["Product", "Marketing", "Integration", "GTM"]
    workstreams = [
        Workstream(workstream=n, order=i + 1, color=_random_hex()) for i, n in enumerate(ws_names)
    ]

    tasks: List[Task] = []
    task_id = 1

    # Generate overlapping blocks in each workstream
    for ws in ws_names:
        base = overall_start + timedelta(days=random.randint(0, 60))
        for _ in range(10):
            start = base + timedelta(days=random.randint(0, 240))
            dur = random.randint(3, 35)
            end = min(start + timedelta(days=dur), overall_end)
            status = random.choice(["planned", "in_progress", "done", "risk"])
            ttype = random.choice(["block", "milestone", "block", "block"])  # bias toward blocks
            if ttype == "milestone":
                end = start

            tasks.append(
                Task(
                    id=f"T{task_id:03d}",
                    workstream=ws,
                    title=f"{ws} task {task_id}",
                    description="Some descriptive text that might wrap depending on space.",
                    start_date=start,
                    end_date=end,
                    status=status,
                    owner=random.choice(["", "Jag", "Team", "PMO"]),
                    type=ttype,
                    hyperlink="https://example.com" if random.random() < 0.15 else None,
                )
            )
            task_id += 1

    # Add a couple out-of-range tasks (to ensure warnings / hiding works)
    tasks.append(
        Task(
            id=f"T{task_id:03d}",
            workstream="Product",
            title="Out of range (before)",
            description=None,
            start_date=overall_start - timedelta(days=40),
            end_date=overall_start - timedelta(days=10),
            status="planned",
            type="block",
        )
    )
    task_id += 1

    tasks.append(
        Task(
            id=f"T{task_id:03d}",
            workstream="GTM",
            title="Out of range (after)",
            description=None,
            start_date=overall_end + timedelta(days=10),
            end_date=overall_end + timedelta(days=20),
            status="planned",
            type="block",
        )
    )

    # Schedule (sublanes)
    scheduled = schedule_by_workstream(tasks, touching_counts_as_overlap=True)
    by_id = {t.id: t for ws_tasks in scheduled.values() for t in ws_tasks}
    tasks = [by_id[t.id] for t in tasks]

    return settings, workstreams, tasks


def main() -> None:
    results: List[SmokeResult] = []

    for i in range(1, 6):
        settings, workstreams, tasks = build_random_case(i)

        # Round-trip through Excel writer/reader to simulate the app flow.
        workstreams_df = pd.DataFrame([w.model_dump() for w in workstreams])
        tasks_df = pd.DataFrame([t.model_dump() for t in tasks])

        for c in WORKSTREAM_COLUMNS:
            if c not in workstreams_df.columns:
                workstreams_df[c] = pd.NA
        workstreams_df = workstreams_df[WORKSTREAM_COLUMNS]

        for c in TASK_COLUMNS:
            if c not in tasks_df.columns:
                tasks_df[c] = pd.NA
        tasks_df = tasks_df[TASK_COLUMNS]

        xlsx = write_roadmap_excel_bytes(settings.model_dump(), workstreams_df, tasks_df)
        payload = read_roadmap_excel(xlsx)

        # Rebuild models from round-tripped payload
        settings2 = Settings(**payload.settings)
        ws2 = [Workstream(**r) for r in payload.workstreams_df.to_dict(orient="records") if str(r.get("workstream") or "").strip()]

        task_models: List[Task] = []
        for r in payload.tasks_df.to_dict(orient="records"):
            if not str(r.get("id") or "").strip():
                continue
            # pandas will represent blanks as NaN (float), which pydantic treats as an invalid string.
            for k, v in list(r.items()):
                if isinstance(v, float) and pd.isna(v):
                    r[k] = None
            task_models.append(Task(**r))

        # Schedule again to ensure deterministic behavior
        scheduled2 = schedule_by_workstream(task_models, touching_counts_as_overlap=True)
        by_id2 = {t.id: t for ws_tasks in scheduled2.values() for t in ws_tasks}
        task_models = [by_id2[t.id] for t in task_models]

        preview = preview_png_bytes(settings2, ws2, task_models, include_out_of_range=False, dpi=140)
        png = export_png_bytes(settings2, ws2, task_models, include_out_of_range=False, dpi=300)
        pdf = export_pdf_bytes(settings2, ws2, task_models, include_out_of_range=False)
        pptx = export_pptx_bytes(settings2, ws2, task_models, include_out_of_range=False)

        assert preview[:8] == b"\x89PNG\r\n\x1a\n"
        assert png[:8] == b"\x89PNG\r\n\x1a\n"
        assert pdf[:4] == b"%PDF"
        assert pptx[:2] == b"PK"

        results.append(
            SmokeResult(
                iteration=i,
                workstreams=len(ws2),
                tasks=len(task_models),
                preview_png_bytes=len(preview),
                png_bytes=len(png),
                pdf_bytes=len(pdf),
                pptx_bytes=len(pptx),
            )
        )

    print("Smoke test results")
    for r in results:
        print(
            f"- Iter {r.iteration}: workstreams={r.workstreams}, tasks={r.tasks}, "
            f"preview_png={r.preview_png_bytes:,} B, png={r.png_bytes:,} B, pdf={r.pdf_bytes:,} B, pptx={r.pptx_bytes:,} B"
        )

    print("OK: all iterations exported preview PNG, high-res PNG, PDF, and PPTX successfully.")


if __name__ == "__main__":
    main()
