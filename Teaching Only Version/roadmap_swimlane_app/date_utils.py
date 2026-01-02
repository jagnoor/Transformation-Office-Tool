from __future__ import annotations


# =============================================================================
# Teaching notes: date_utils.py (date → x coordinate helpers)
#
# Matplotlib represents dates as a float (days since an epoch).
# We use a tiny helper here so the rest of the code can work in "date" objects.
#
# This file is tiny on purpose: it makes date math easy to test.
# =============================================================================

# ---------------------------------------------------------------------------
# Imports (standard library, third-party packages, and local modules)
# ---------------------------------------------------------------------------
from datetime import date, timedelta
from typing import Tuple

# Matplotlib's default epoch is 1970-01-01 (days since epoch).
# This matches matplotlib.dates.get_epoch() default in modern matplotlib.
_MPL_EPOCH = date(1970, 1, 1)


def date_to_x(d: date) -> float:
    """
    Convert a date to matplotlib "date number" units without importing matplotlib.
    Unit: days since 1970-01-01 (float).
    """
    return float((d - _MPL_EPOCH).days)


def block_span_inclusive(start: date, end: date) -> Tuple[float, float]:
    """
    Inclusive end-date semantics:
      - A same-day task spans 1 day of width.
      - Render interval is [start, end + 1 day) in date-number units.
    """
    x0 = date_to_x(start)
    x1 = date_to_x(end + timedelta(days=1))
    return x0, x1
