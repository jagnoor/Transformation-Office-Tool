import sys
from pathlib import Path

# Force a headless backend for matplotlib before any pyplot imports.
import matplotlib

matplotlib.use("Agg")

# Ensure project root is on sys.path for tests
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
