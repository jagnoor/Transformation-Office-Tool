from __future__ import annotations

from io import BytesIO
from typing import List, Tuple

from roadmap_models import Settings, Task, Workstream
from renderer import render_roadmap
from pptx_export import export_pptx_bytes as _export_pptx_bytes


def export_pdf_bytes(
    settings: Settings,
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    include_out_of_range: bool = False,
) -> bytes:
    fig, _, _ = render_roadmap(settings, workstreams, tasks, include_out_of_range=include_out_of_range, preview=False)
    bio = BytesIO()
    fig.savefig(bio, format="pdf", facecolor="white")
    # Important: close to avoid memory growth in Streamlit
    import matplotlib.pyplot as plt
    plt.close(fig)
    return bio.getvalue()


def export_png_bytes(
    settings: Settings,
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    include_out_of_range: bool = False,
    dpi: int = 300,
) -> bytes:
    # We let the caller choose high-res DPI independently of settings.output_dpi
    settings2 = settings.model_copy(update={"output_dpi": dpi})
    fig, _, _ = render_roadmap(settings2, workstreams, tasks, include_out_of_range=include_out_of_range, preview=False)
    bio = BytesIO()
    fig.savefig(bio, format="png", dpi=dpi, facecolor="white")
    import matplotlib.pyplot as plt
    plt.close(fig)
    return bio.getvalue()


def preview_png_bytes(
    settings: Settings,
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    include_out_of_range: bool = False,
    dpi: int = 150,
) -> bytes:
    fig, _, _ = render_roadmap(settings, workstreams, tasks, include_out_of_range=include_out_of_range, preview=True, preview_dpi=dpi)
    bio = BytesIO()
    fig.savefig(bio, format="png", dpi=dpi, facecolor="white")
    import matplotlib.pyplot as plt
    plt.close(fig)
    return bio.getvalue()


def export_pptx_bytes(
    settings: Settings,
    workstreams: List[Workstream],
    tasks: List[Task],
    *,
    include_out_of_range: bool = False,
) -> bytes:
    """Export an editable PowerPoint slide as a .pptx byte string."""

    return _export_pptx_bytes(
        settings,
        workstreams,
        tasks,
        include_out_of_range=include_out_of_range,
    )
