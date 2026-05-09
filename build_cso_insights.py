"""
build_cso_insights.py
Builds the Olympic Paints CSO Insights dashboard.
Computes Store Buying Frequency and Product Mix from the sales parquet.
Rep Performance is sourced from build_kpi_dashboard.py (single source of truth).

Run:  python build_cso_insights.py
"""

import json
import shutil
import subprocess
import sys
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd

BASE_DIR   = Path(__file__).parent
OUT_DIR    = BASE_DIR / "cso_insights"
OUT_HTML   = OUT_DIR / "index.html"
LOGO_SRC   = Path(r"C:\Users\quint\OneDrive\1.Projects\1.Olympic Paints\3.Resources\9. Brand Assets & Images\Misc Pictures\Olympic Paints Logo Digital.jpg")
PARQUET    = Path(r"C:\Users\quint\OneDrive\1.Projects\1.Olympic Paints\3.Resources\16.Sales and Other data\Sales_Invoices_All.parquet")
REPORT_DATE = datetime.now().strftime("%Y-%m-%d")
REPO_URL   = "https://github.com/FlomaticAuto/olympic-paints-cso-insights.git"
