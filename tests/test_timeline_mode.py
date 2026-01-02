from datetime import date

from renderer import choose_timeline_mode


def test_timeline_mode_weeks_under_4_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2026, 3, 31)) == "weeks"  # 3 months


def test_timeline_mode_months_at_4_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2026, 4, 30)) == "months"  # 4 months


def test_timeline_mode_months_at_12_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2026, 12, 31)) == "months"  # 12 months


def test_timeline_mode_quarters_over_12_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2027, 1, 2)) == "quarters"  # 13 months


def test_timeline_mode_quarters_at_24_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2027, 12, 31)) == "quarters"  # 24 months


def test_timeline_mode_quarters_years_over_24_months():
    assert choose_timeline_mode(date(2026, 1, 1), date(2028, 1, 1)) == "quarters_years"  # 25 months
