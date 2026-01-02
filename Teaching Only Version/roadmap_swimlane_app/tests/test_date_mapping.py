# =============================================================================
# Teaching notes: tests/test_date_mapping.py
#
# Unit tests for date_utils.py.
# Confirms date-to-x mapping and inclusive end-date behavior.
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
from datetime import date

from date_utils import block_span_inclusive, date_to_x


def test_block_span_inclusive_same_day_width_one_day():
    x0, x1 = block_span_inclusive(date(2025, 1, 1), date(2025, 1, 1))
    assert x1 > x0
    # In matplotlib date units, 1 day == 1.0
    assert abs((x1 - x0) - 1.0) < 1e-9


def test_date_to_x_monotonic_increasing():
    d1 = date(2025, 1, 1)
    d2 = date(2025, 1, 2)
    assert date_to_x(d2) > date_to_x(d1)
