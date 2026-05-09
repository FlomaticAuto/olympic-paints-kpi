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

# ── Rep Performance — imported from KPI dashboard (single source of truth) ────
sys.path.insert(0, str(BASE_DIR))
from build_kpi_dashboard import REPS, REPORT_WEEK


def compute_store_buying_frequency(df: pd.DataFrame) -> dict:
    """Top 20 stores by order count over last 12 months, with avg days between orders."""
    cutoff = pd.Timestamp(date.today() - timedelta(days=365))
    recent = df[df["trandate"] >= cutoff].copy()

    # One row per delivery (delno) per store — count distinct deliveries as orders
    orders = (
        recent.groupby(["accno", "store_name"])
        .agg(
            order_count=("delno", "nunique"),
            last_order=("trandate", "max"),
            first_order=("trandate", "min"),
        )
        .reset_index()
        .sort_values("order_count", ascending=False)
        .head(20)
    )

    # Avg days between orders = date range / (orders - 1), floor at 1 order
    def avg_days(row):
        if row["order_count"] <= 1:
            return None
        span = (row["last_order"] - row["first_order"]).days
        return round(span / (row["order_count"] - 1), 1)

    orders["avg_days"] = orders.apply(avg_days, axis=1)

    labels = orders["store_name"].tolist()
    counts = orders["order_count"].tolist()

    rows = [
        {
            "store_name": r["store_name"],
            "accno": r["accno"],
            "orders_12m": int(r["order_count"]),
            "last_order": r["last_order"].strftime("%Y-%m-%d"),
            "avg_days": r["avg_days"] if r["avg_days"] else "—",
        }
        for _, r in orders.iterrows()
    ]

    if orders.empty:
        analysis_text = "No store order data available for the last 12 months."
    else:
        analysis_text = (
            f"The top 20 stores by order frequency over the last 12 months are shown below. "
            f"<strong>{orders.iloc[0]['store_name']}</strong> leads with "
            f"<strong>{int(orders.iloc[0]['order_count'])} orders</strong>. "
            f"High-frequency buyers represent your most reliable revenue base — "
            f"stores ordering more than once a month warrant priority service attention."
        )

    return {
        "id": "store-buying-frequency",
        "title": "Store Buying Frequency",
        "summary": "How often each store places orders (last 12 months)",
        "updated": REPORT_DATE,
        "icon": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>',
        "analysis": analysis_text,
        "chart": {
            "type": "bar",
            "options": {"indexAxis": "y"},
            "labels": labels,
            "datasets": [
                {
                    "label": "Orders (last 12m)",
                    "data": counts,
                    "backgroundColor": "#F5C400",
                    "borderRadius": 4,
                }
            ],
        },
        "table": {
            "columns": [
                {"key": "store_name", "label": "Store"},
                {"key": "accno", "label": "Account Ref"},
                {"key": "orders_12m", "label": "Orders (12m)"},
                {"key": "last_order", "label": "Last Order"},
                {"key": "avg_days", "label": "Avg Days Between"},
            ],
            "rows": rows,
        },
    }
