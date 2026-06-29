"""Excel input/output — dead simple format for non-technical users."""
import io
from datetime import date, datetime
from typing import Tuple, List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from models import WorkItem, ChartConfig


# ── Reading Excel ────────────────────────────────────────────────────────────

_DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y", "%m/%d/%y", "%Y/%m/%d")


def _parse_date(val) -> date:
    """Parse a date from various formats."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        raise ValueError("Date is required")
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"Cannot parse date '{val}'. Use YYYY-MM-DD, MM/DD/YYYY, or DD/MM/YYYY."
    )


def read_excel(file) -> Tuple[List[WorkItem], ChartConfig, List[str]]:
    """Read work items from an Excel file.

    Expects a single sheet with columns:
      Title (required), Start Date (required), End Date (required),
      Category (optional), Description (optional), Status (optional),
      Owner (optional), Label (optional)

    Returns (items, config) where config has auto-detected settings.
    """
    df = pd.read_excel(file, sheet_name=0, engine="openpyxl")

    # Normalize column names
    col_map = {}
    for col in df.columns:
        normalized = str(col).strip().lower().replace(" ", "_").replace("-", "_")
        col_map[col] = normalized
    df = df.rename(columns=col_map)

    # Find required columns with flexible matching
    title_col = _find_col(df, ["title", "name", "task", "item", "work_item", "deliverable"])
    start_col = _find_col(df, ["start_date", "start", "begin", "from", "begin_date"])
    end_col = _find_col(df, ["end_date", "end", "finish", "to", "due", "due_date", "finish_date"])

    if not title_col:
        raise ValueError("Missing required column: 'Title' (or 'Name', 'Task', 'Item')")
    if not start_col:
        raise ValueError("Missing required column: 'Start Date' (or 'Start', 'Begin')")
    if not end_col:
        raise ValueError("Missing required column: 'End Date' (or 'End', 'Finish', 'Due')")

    # Optional columns
    cat_col = _find_col(df, ["category", "workstream", "team", "group", "stream", "department"])
    desc_col = _find_col(df, ["description", "desc", "details", "notes", "note"])
    status_col = _find_col(df, ["status", "state", "progress"])
    owner_col = _find_col(df, ["owner", "assigned", "assignee", "responsible", "lead"])
    label_col = _find_col(df, ["label", "tag", "milestone_label", "id"])

    items = []
    warnings = []
    for idx, row in df.iterrows():
        title = str(row.get(title_col, "")).strip()
        if not title or title == "nan":
            continue
        try:
            start = _parse_date(row.get(start_col))
            end = _parse_date(row.get(end_col))
        except (ValueError, TypeError) as e:
            warnings.append(f"Row {idx + 2}: {e} — skipping")
            continue

        if end < start:
            start, end = end, start
            warnings.append(f"Row {idx + 2}: Start/End dates swapped for '{title}'")

        category = _safe_str(row.get(cat_col)) if cat_col else "General"
        description = _safe_str(row.get(desc_col)) if desc_col else ""
        status = _safe_str(row.get(status_col)).lower().replace(" ", "_") if status_col else "planned"
        owner = _safe_str(row.get(owner_col)) if owner_col else ""
        label = _safe_str(row.get(label_col)) if label_col else ""

        if not category:
            category = "General"
        if status not in ("planned", "in_progress", "done", "at_risk"):
            status = "planned"

        items.append(WorkItem(
            title=title,
            start_date=start,
            end_date=end,
            category=category,
            description=description,
            status=status,
            owner=owner,
            label=label,
        ))

    # Auto-detect config from data
    config = ChartConfig()
    if items:
        config.start_date = min(it.start_date for it in items)
        config.end_date = max(it.end_date for it in items)

    return items, config, warnings


def _find_col(df, candidates):
    """Find a column matching any of the candidate names."""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _safe_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


# ── Writing Template ─────────────────────────────────────────────────────────

def create_template_bytes() -> bytes:
    """Create a beautiful, user-friendly Excel template."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Roadmap"

    # Styling
    header_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    # Headers
    headers = [
        ("Title", 35),
        ("Start Date", 15),
        ("End Date", 15),
        ("Category", 20),
        ("Description", 40),
        ("Status", 15),
        ("Owner", 20),
        ("Label", 15),
    ]

    for col_idx, (name, width) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Status dropdown validation
    status_dv = DataValidation(
        type="list",
        formula1='"planned,in_progress,done,at_risk"',
        allow_blank=True,
    )
    status_dv.error = "Please select: planned, in_progress, done, or at_risk"
    status_dv.errorTitle = "Invalid Status"
    ws.add_data_validation(status_dv)
    status_dv.add(f"F2:F200")

    # Sample data — a 1-year transformation roadmap with 25 deliveries.
    # Deliveries are prioritized (Delivery 1 = highest priority, starts first).
    # Multiple workstreams run concurrently to create a dense Tetris-style visualization.
    sample_data = [
        # Wave 1: Foundation (Jan–Mar) — 8 concurrent deliveries
        ("Strategic Vision & Roadmap",   "2025-01-06", "2025-03-14", "Strategy",    "Define transformation goals; Stakeholder alignment; Success metrics; Executive sign-off", "done", "Sarah", "Delivery 1"),
        ("Customer Research & Insights", "2025-01-06", "2025-03-28", "Product",     "Recruit participants; Conduct 30 interviews; Synthesize findings; Present to leadership", "done", "Lisa", "Delivery 2"),
        ("Brand Identity Refresh",       "2025-01-06", "2025-04-11", "Marketing",   "New visual identity; Logo redesign; Brand guidelines; Template library; Asset rollout", "done", "Tom", "Delivery 3"),
        ("Technology Assessment",        "2025-01-06", "2025-03-28", "Engineering", "Platform audit; Architecture review; Tool evaluation; Migration plan; Vendor demos", "done", "Mike", "Delivery 4"),
        ("Vendor Selection & Contracts", "2025-01-06", "2025-02-28", "Operations",  "RFP process; Vendor demos; Contract negotiation; Legal review; Final selection", "done", "Karen", "Delivery 5"),
        ("CRM & Pipeline Setup",         "2025-01-06", "2025-03-14", "Sales",       "Salesforce configuration; Pipeline stages; Reporting dashboards; Data migration", "done", "Nina", "Delivery 6"),
        ("Pricing Strategy",             "2025-01-06", "2025-02-14", "Sales",       "Competitive analysis; Pricing tiers; Discount framework; Approval workflows", "done", "Jake", "Delivery 7"),
        ("Hiring Plan Execution",        "2025-01-06", "2025-04-25", "Operations",  "Job postings; Candidate screening; Interview rounds; Offer management; Onboarding 12 roles", "done", "Karen", "Delivery 8"),
        # Wave 2: Build (Mar–Jun) — 7 concurrent deliveries
        ("Product Requirements & Specs", "2025-03-03", "2025-05-09", "Product",     "Define MVP features; Write specifications; Stakeholder review; Final sign-off", "in_progress", "Lisa", "Delivery 9"),
        ("Website Redesign",             "2025-03-03", "2025-06-06", "Marketing",   "Information architecture; UX wireframes; Visual design; Content migration; QA testing", "in_progress", "Sarah", "Delivery 10"),
        ("Core Platform Build",          "2025-03-03", "2025-07-11", "Engineering", "Backend services; Database design; API gateway; Authentication; CI/CD pipeline", "in_progress", "Mike", "Delivery 11"),
        ("Sales Playbook Creation",      "2025-03-17", "2025-05-23", "Sales",       "Pitch decks; Battle cards; Objection handling; ROI calculator; Competitive positioning", "in_progress", "Jake", "Delivery 12"),
        ("UX/UI Design & Prototypes",    "2025-04-07", "2025-06-20", "Product",     "User flows; Wireframes; High-fidelity mockups; Clickable prototype; Usability testing", "in_progress", "Amy", "Delivery 13"),
        ("Payment Integration",          "2025-04-14", "2025-06-27", "Engineering", "Stripe integration; PayPal setup; Webhook handling; Reconciliation; PCI compliance", "in_progress", "Dave", "Delivery 14"),
        ("Office Expansion Setup",       "2025-04-28", "2025-06-20", "Operations",  "Floor plan design; Construction; IT infrastructure; Furniture procurement; Move logistics", "planned", "Karen", "Delivery 15"),
        # Wave 3: Scale (Jun–Sep) — 5 concurrent deliveries
        ("Product Launch Campaign",      "2025-06-02", "2025-08-15", "Marketing",   "Campaign strategy; Social media content; Email sequences; Press releases; Launch event", "planned", "Tom", "Delivery 16"),
        ("Mobile App Development",       "2025-06-16", "2025-09-19", "Engineering", "iOS development; Android development; Push notifications; Offline mode; App Store submission", "planned", "Mike", "Delivery 17"),
        ("Sales Team Training",          "2025-06-16", "2025-08-08", "Sales",       "Product training; Demo certification; Role-play exercises; Assessment; Field readiness", "planned", "Jake", "Delivery 18"),
        ("Beta Testing Program",         "2025-06-30", "2025-09-05", "Product",     "Recruit 50 beta testers; Onboarding; Feedback collection; Bug triage; Iteration cycles", "planned", "Lisa", "Delivery 19"),
        ("Partner Onboarding Program",   "2025-07-14", "2025-09-19", "Operations",  "Partner agreements; Technical integration; Training sessions; Go-live support; SLA setup", "planned", "Rob", "Delivery 20"),
        # Wave 4: Launch & Optimize (Sep–Dec) — 5 concurrent deliveries
        ("Pilot Customer Program",       "2025-09-01", "2025-11-14", "Sales",       "Identify 10 prospects; Negotiate terms; Onboard pilots; Track success metrics; Case studies", "planned", "Nina", "Delivery 21"),
        ("Performance & Load Testing",   "2025-09-01", "2025-10-31", "Engineering", "Load test scripts; Stress testing; Performance optimization; Capacity planning; SLA validation", "planned", "Dave", "Delivery 22"),
        ("Product Documentation",        "2025-09-15", "2025-11-28", "Product",     "User guides; API documentation; Help center articles; Video tutorials; Knowledge base", "planned", "Amy", "Delivery 23"),
        ("Customer Testimonials",        "2025-10-01", "2025-12-05", "Marketing",   "Identify advocates; Schedule interviews; Video production; Post-production; Distribution", "planned", "Sarah", "Delivery 24"),
        ("Security Audit & Compliance",  "2025-10-13", "2025-12-19", "Engineering", "Third-party assessment; Vulnerability remediation; Penetration testing; SOC 2 prep; Certification", "planned", "Mike", "Delivery 25"),
    ]

    data_font = Font(name="Calibri", size=11)
    data_align = Alignment(vertical="center", wrap_text=False)
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for row_idx, row_data in enumerate(sample_data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    # Freeze header row
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = "2563EB"

    # Instructions sheet
    ws_info = wb.create_sheet("Instructions")
    ws_info.sheet_properties.tabColor = "10B981"

    instructions = [
        ("Block & Gantt Creator — Input Guide", ""),
        ("", ""),
        ("Required Columns:", ""),
        ("Title", "Name of the work item (e.g., 'API Gateway Development')"),
        ("Start Date", "When work begins (YYYY-MM-DD format)"),
        ("End Date", "When work ends (YYYY-MM-DD format)"),
        ("", ""),
        ("Optional Columns:", ""),
        ("Category", "Team or workstream (e.g., Platform, Product, Security)"),
        ("Description", "Brief description of the work"),
        ("Status", "planned, in_progress, done, or at_risk"),
        ("Owner", "Person responsible"),
        ("Label", "Short label (e.g., D1, MVP, Phase 2)"),
        ("", ""),
        ("Tips:", ""),
        ("", "• Categories are used to color-code items and group them in swim lanes"),
        ("", "• For milestones, set Start Date = End Date"),
        ("", "• The app auto-detects date ranges — no settings sheet needed"),
        ("", "• Column names are flexible: 'Name' works for 'Title', 'Team' for 'Category', etc."),
        ("", ""),
        ("Block Diagram tips:", ""),
        ("", "• Use the Label column for short identifiers (e.g., 'D1', 'MVP', 'Phase 2')"),
        ("", "• Descriptions are shown inside blocks when there is enough space"),
        ("", "• Overlapping date ranges stack vertically to show parallel work"),
    ]

    title_font = Font(name="Calibri", bold=True, size=16, color="1E293B")
    section_font = Font(name="Calibri", bold=True, size=12, color="2563EB")
    key_font = Font(name="Calibri", bold=True, size=11, color="334155")
    val_font = Font(name="Calibri", size=11, color="64748B")

    for row_idx, (key, val) in enumerate(instructions, 1):
        c1 = ws_info.cell(row=row_idx, column=1, value=key)
        c2 = ws_info.cell(row=row_idx, column=2, value=val)
        if row_idx == 1:
            c1.font = title_font
        elif key and not val:
            c1.font = section_font
        elif key and val:
            c1.font = key_font
            c2.font = val_font
        else:
            c2.font = val_font

    ws_info.column_dimensions["A"].width = 25
    ws_info.column_dimensions["B"].width = 70

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def create_sample_bytes() -> bytes:
    """Create a sample Excel with realistic data (same as template's sample data)."""
    return create_template_bytes()
