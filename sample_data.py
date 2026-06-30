"""Built-in sample datasets for exploring the tool without uploading a file.

Each dataset is a list of raw row tuples in the same shape used by the Excel
template: (title, start, end, category, description, status, owner, label).
Keeping them as plain tuples lets both the Excel template writer and the
in-app "Load Sample Data" flow share a single source of truth.
"""
from datetime import date
from typing import List

from models import WorkItem


# ── Product Launch — software/product roadmap across 6 workstreams ──────────
PRODUCT_LAUNCH_ROWS = [
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

# ── Marketing Campaign — go-to-market campaign across 3 workstreams ─────────
MARKETING_CAMPAIGN_ROWS = [
    ("Campaign Strategy & Briefing",  "2025-01-06", "2025-01-24", "Strategy", "Define objectives; Audience segments; Budget allocation; KPI framework", "done", "Maria", "Phase 1"),
    ("Creative Concept Development",  "2025-01-20", "2025-02-21", "Brand",    "Mood boards; Concept testing; Creative direction; Asset briefs", "done", "Chris", "Phase 1"),
    ("Media Plan & Buying",           "2025-01-20", "2025-02-28", "Digital",  "Channel mix; Budget split; Vendor negotiation; Placements booked", "done", "Jordan", "Phase 1"),
    ("Website Landing Pages",         "2025-02-03", "2025-03-14", "Digital",  "Wireframes; Copywriting; Design; Build; QA", "in_progress", "Priya", "Phase 2"),
    ("Social Content Calendar",       "2025-02-10", "2025-03-21", "Brand",    "Content pillars; 60-post calendar; Influencer outreach; Approvals", "in_progress", "Chris", "Phase 2"),
    ("Email Nurture Sequence",        "2025-02-17", "2025-03-14", "Digital",  "5-email sequence; Segmentation; A/B subject lines; Automation setup", "in_progress", "Jordan", "Phase 2"),
    ("PR & Press Outreach",           "2025-02-24", "2025-04-04", "Brand",    "Press list; Embargoed briefing; Press kit; Pitch outreach", "planned", "Maria", "Phase 2"),
    ("Paid Search & Social Ads",      "2025-03-10", "2025-05-02", "Digital",  "Ad creative; Targeting; Budget pacing; Optimization", "planned", "Jordan", "Phase 3"),
    ("Launch Event Planning",         "2025-03-10", "2025-04-25", "Events",   "Venue booking; Guest list; Run of show; Vendor coordination", "planned", "Priya", "Phase 3"),
    ("Influencer Partnerships",       "2025-03-17", "2025-04-18", "Brand",    "Influencer shortlist; Contracts; Content briefs; Posting schedule", "planned", "Chris", "Phase 3"),
    ("Launch Day",                    "2025-04-25", "2025-04-25", "Events",   "", "planned", "Maria", "Milestone"),
    ("Post-Launch Reporting",         "2025-04-28", "2025-05-16", "Strategy", "Performance dashboard; ROI analysis; Lessons learned; Exec readout", "planned", "Maria", "Phase 4"),
    ("Retargeting & Always-On",       "2025-05-05", "2025-07-11", "Digital",  "Retargeting audiences; Always-on budget; Ongoing optimization", "planned", "Jordan", "Phase 4"),
    ("Customer Story Collection",     "2025-05-12", "2025-06-27", "Brand",    "Identify advocates; Interviews; Case study production", "planned", "Chris", "Phase 4"),
]

# ── Construction Project — facility build-out across 4 trades ───────────────
CONSTRUCTION_PROJECT_ROWS = [
    ("Site Survey & Feasibility",       "2025-01-06", "2025-02-07", "Design",     "Topographic survey; Soil testing; Feasibility report; Budget estimate", "done", "Wei", "Phase 1"),
    ("Architectural Design",            "2025-02-03", "2025-04-11", "Design",     "Schematic design; Design development; Client review; Final drawings", "done", "Marcus", "Phase 1"),
    ("Permitting & Approvals",          "2025-03-17", "2025-06-06", "Permitting", "Building permit application; Zoning review; Inspections scheduling", "done", "Aisha", "Phase 1"),
    ("Site Preparation & Excavation",   "2025-06-09", "2025-07-11", "Construction", "Clearing; Grading; Excavation; Erosion control", "in_progress", "Tom", "Phase 2"),
    ("Foundation Work",                 "2025-07-07", "2025-08-15", "Construction", "Footings; Foundation walls; Waterproofing; Backfill", "in_progress", "Tom", "Phase 2"),
    ("Structural Framing",              "2025-08-11", "2025-10-03", "Construction", "Steel/wood framing; Roof trusses; Sheathing; Inspections", "planned", "Marcus", "Phase 3"),
    ("Roofing",                         "2025-09-29", "2025-10-24", "Construction", "Roof deck; Underlayment; Shingles/membrane; Flashing", "planned", "Tom", "Phase 3"),
    ("Electrical Rough-In",             "2025-10-06", "2025-11-07", "MEP",        "Panel install; Wiring; Rough-in inspection", "planned", "Aisha", "Phase 3"),
    ("Plumbing Rough-In",               "2025-10-06", "2025-11-07", "MEP",        "Supply lines; Drain-waste-vent; Rough-in inspection", "planned", "Wei", "Phase 3"),
    ("HVAC Installation",               "2025-10-20", "2025-11-21", "MEP",        "Ductwork; Equipment install; System testing", "planned", "Marcus", "Phase 3"),
    ("Insulation & Drywall",            "2025-11-10", "2025-12-12", "Finishing",  "Insulation install; Drywall hang; Tape and texture", "planned", "Tom", "Phase 4"),
    ("Interior Finishing",              "2025-12-08", "2026-02-06", "Finishing",  "Flooring; Cabinetry; Paint; Trim; Fixtures", "planned", "Marcus", "Phase 4"),
    ("Final Inspections & Punch List",  "2026-02-02", "2026-02-27", "Permitting", "Final inspection; Punch list walkthrough; Corrections", "planned", "Aisha", "Phase 5"),
    ("Certificate of Occupancy",        "2026-03-02", "2026-03-02", "Permitting", "", "planned", "Aisha", "Milestone"),
]

# ── Hiring Plan — org build-out across 4 workstreams ─────────────────────────
HIRING_PLAN_ROWS = [
    ("Workforce Planning",            "2025-01-06", "2025-01-24", "Leadership",  "Headcount plan; Role definitions; Budget approval; Org chart draft", "done", "Elena", "Phase 1"),
    ("Job Descriptions & Leveling",   "2025-01-20", "2025-02-14", "HR Ops",      "Write JDs for 15 roles; Leveling guide; Compensation bands", "done", "Priyanka", "Phase 1"),
    ("Sourcing Strategy",             "2025-01-27", "2025-02-21", "Recruiting",  "Channel strategy; Agency partnerships; Referral program design", "done", "Sam", "Phase 1"),
    ("Engineering Hiring Wave",       "2025-02-17", "2025-05-09", "Recruiting",  "Sourcing; Screening; Interview loops; Offers — 6 roles", "in_progress", "Sam", "Phase 2"),
    ("Sales Hiring Wave",             "2025-02-24", "2025-04-25", "Recruiting",  "Sourcing; Screening; Interview loops; Offers — 4 roles", "in_progress", "Sam", "Phase 2"),
    ("Interview Panel Training",      "2025-02-24", "2025-03-14", "HR Ops",      "Structured interview training; Bias workshop; Scorecards rollout", "in_progress", "Priyanka", "Phase 2"),
    ("Benefits & Payroll Setup",      "2025-03-03", "2025-04-04", "HR Ops",      "Benefits vendor selection; Payroll system config; Policy docs", "in_progress", "Elena", "Phase 2"),
    ("Onboarding Program Design",     "2025-03-17", "2025-04-18", "Onboarding",  "30-60-90 plan template; Welcome kit; Buddy program", "planned", "Noah", "Phase 2"),
    ("Leadership Hiring",             "2025-03-24", "2025-06-13", "Leadership",  "Exec search; Panel interviews; Reference checks — 2 roles", "planned", "Elena", "Phase 3"),
    ("Operations Hiring Wave",        "2025-04-07", "2025-06-06", "Recruiting",  "Sourcing; Screening; Interview loops; Offers — 3 roles", "planned", "Sam", "Phase 3"),
    ("New Hire Onboarding Cohort 1",  "2025-05-12", "2025-05-23", "Onboarding",  "Orientation week; Systems access; Manager 1:1s; First project", "planned", "Noah", "Phase 3"),
    ("Manager Training Program",      "2025-05-19", "2025-06-20", "HR Ops",      "First-time manager training; Feedback skills; Performance reviews", "planned", "Priyanka", "Phase 3"),
    ("New Hire Onboarding Cohort 2",  "2025-06-23", "2025-07-04", "Onboarding",  "Orientation week; Systems access; Manager 1:1s; First project", "planned", "Noah", "Phase 4"),
    ("90-Day Retention Check-In",     "2025-08-04", "2025-08-15", "HR Ops",      "Pulse survey; Manager feedback loop; Retention risk review", "planned", "Elena", "Phase 4"),
]


SAMPLE_DATASETS = {
    "Product Launch": {
        "rows": PRODUCT_LAUNCH_ROWS,
        "title": "Transformation Roadmap 2025",
        "subtitle": "25 initiatives across 6 workstreams",
        "palette": "Vibrant",
        "blurb": "Software product launch across Strategy, Product, Marketing, Engineering, Sales, and Operations.",
    },
    "Marketing Campaign": {
        "rows": MARKETING_CAMPAIGN_ROWS,
        "title": "Campaign Launch Plan",
        "subtitle": "14 workstreams across Strategy, Brand, Digital, and Events",
        "palette": "Sunset",
        "blurb": "Go-to-market campaign plan from strategy through launch and post-launch reporting.",
    },
    "Construction Project": {
        "rows": CONSTRUCTION_PROJECT_ROWS,
        "title": "Facility Build-Out",
        "subtitle": "14 phases across Design, Permitting, Construction, and MEP",
        "palette": "Forest",
        "blurb": "Ground-up construction project from site survey through certificate of occupancy.",
    },
    "Hiring Plan": {
        "rows": HIRING_PLAN_ROWS,
        "title": "Org Build-Out Plan",
        "subtitle": "14 initiatives across Recruiting, HR Ops, Onboarding, and Leadership",
        "palette": "Ocean",
        "blurb": "Workforce planning and phased hiring waves with onboarding cohorts.",
    },
}


def list_dataset_names() -> List[str]:
    """Return dataset names in a stable, deliberate order."""
    return list(SAMPLE_DATASETS.keys())


def get_dataset_meta(name: str) -> dict:
    """Return the display metadata (title/subtitle/palette/blurb) for a dataset."""
    meta = SAMPLE_DATASETS[name]
    return {k: v for k, v in meta.items() if k != "rows"}


def get_sample_rows(name: str):
    """Return the raw row tuples for a dataset (used by the Excel template writer)."""
    return SAMPLE_DATASETS[name]["rows"]


def get_sample_items(name: str) -> List[WorkItem]:
    """Build a fresh list of WorkItem objects for a dataset.

    Returns new instances each call so callers can safely mutate the result
    without affecting other sessions sharing this module.
    """
    items = []
    for title, start, end, category, description, status, owner, label in get_sample_rows(name):
        items.append(WorkItem(
            title=title,
            start_date=date.fromisoformat(start),
            end_date=date.fromisoformat(end),
            category=category,
            description=description,
            status=status,
            owner=owner,
            label=label,
        ))
    return items
