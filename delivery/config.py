from __future__ import annotations

import os
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_ROOT = ROOT / "output"
SERVICE_PORT = int(os.environ.get("THESIS_REVISER_PORT", "8000"))
SERVICE_HOST = os.environ.get("THESIS_REVISER_HOST", "0.0.0.0")

