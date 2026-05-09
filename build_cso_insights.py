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


def build_rep_performance() -> dict:
    """Rep MTD performance vs target — sourced from REPS in build_kpi_dashboard.py."""
    reps_with_target = [r for r in REPS if r["target"] is not None]

    labels  = [r["name"] for r in reps_with_target]
    actuals = [round(r["sales"] / 1_000_000, 2) for r in reps_with_target]
    targets = [round(r["target"] / 1_000_000, 2) for r in reps_with_target]

    def status(pct):
        if pct is None:
            return "—"
        if pct >= 0:
            return f'<span style="color:var(--color-success-fg)">+{pct:.1f}%</span>'
        return f'<span style="color:var(--color-danger-fg)">{pct:.1f}%</span>'

    rows = [
        {
            "rep":    r["name"],
            "sales":  f"R {r['sales']:,.0f}",
            "target": f"R {r['target']:,.0f}" if r["target"] else "—",
            "pct":    status(r["pct"]),
        }
        for r in REPS
    ]

    leaders = [r for r in reps_with_target if (r["pct"] or 0) > 0]
    leader_text = (
        f"<strong>{leaders[0]['name']}</strong> leads at "
        f"<strong>+{leaders[0]['pct']:.1f}%</strong> above target"
        if leaders else "No rep is currently above target"
    )

    return {
        "id": "rep-performance",
        "title": "Rep Performance",
        "summary": f"MTD sales vs target — {REPORT_WEEK}",
        "updated": REPORT_DATE,
        "icon": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
        "analysis": (
            f"MTD sales performance for week ending <strong>{REPORT_WEEK}</strong>. "
            f"{leader_text}. "
            f"Reps below target require pipeline review — cross-reference outstanding orders "
            f"and last customer contact dates in Zoho."
        ),
        "chart": {
            "type": "bar",
            "options": {},
            "labels": labels,
            "datasets": [
                {
                    "label": "Actual (R millions)",
                    "data": actuals,
                    "backgroundColor": "#F5C400",
                    "borderRadius": 4,
                },
                {
                    "label": "Target (R millions)",
                    "data": targets,
                    "backgroundColor": "#1A3D6E",
                    "borderRadius": 4,
                },
            ],
        },
        "table": {
            "columns": [
                {"key": "rep",    "label": "Rep"},
                {"key": "sales",  "label": "MTD Sales"},
                {"key": "target", "label": "Target"},
                {"key": "pct",    "label": "% vs Target"},
            ],
            "rows": rows,
        },
    }


# Multi-series chart colours — Olympic Paints design system order
_CHART_COLOURS = ["#F5C400","#1A3D6E","#2D8C7A","#C97A3A","#E87BAD","#9B7DBF","#5C6B7A"]

def compute_product_mix(df: pd.DataFrame) -> dict:
    """Revenue by category_l1 (product group) for current FY, with YoY comparison."""
    current_fy = df["fy"].max()
    prev_fy    = current_fy - 1

    def rev_by_group(frame):
        return (
            frame.groupby("category_l1")["ivnett"]
            .sum()
            .reset_index()
            .rename(columns={"ivnett": "revenue"})
        )

    cur  = rev_by_group(df[df["fy"] == current_fy])
    prev = rev_by_group(df[df["fy"] == prev_fy])

    merged = cur.merge(prev, on="category_l1", suffixes=("_cur", "_prev"), how="left")
    merged["revenue_prev"] = merged["revenue_prev"].fillna(0)
    merged["yoy_pct"] = (
        (merged["revenue_cur"] - merged["revenue_prev"])
        / merged["revenue_prev"].replace(0, float("nan"))
        * 100
    ).round(1)
    merged["mix_pct"] = (merged["revenue_cur"] / merged["revenue_cur"].sum() * 100).round(1)
    merged = merged.sort_values("revenue_cur", ascending=False)

    labels   = merged["category_l1"].tolist()
    revenues = [round(v / 1_000_000, 2) for v in merged["revenue_cur"]]
    colours  = [_CHART_COLOURS[i % len(_CHART_COLOURS)] for i in range(len(labels))]

    rows = [
        {
            "group":   r["category_l1"],
            "revenue": f"R {r['revenue_cur']:,.0f}",
            "mix_pct": f"{r['mix_pct']:.1f}%",
            "yoy":     (
                f'<span style="color:var(--color-success-fg)">+{r["yoy_pct"]:.1f}%</span>'
                if r["yoy_pct"] > 0
                else f'<span style="color:var(--color-danger-fg)">{r["yoy_pct"]:.1f}%</span>'
            ) if not pd.isna(r["yoy_pct"]) else "—",
        }
        for _, r in merged.iterrows()
    ]

    top_group = merged.iloc[0]

    return {
        "id": "product-mix",
        "title": "Product Mix",
        "summary": "Revenue by product group — current financial year",
        "updated": REPORT_DATE,
        "icon": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21.21 15.89A10 10 0 118 2.83"/><path d="M22 12A10 10 0 0012 2v10z"/></svg>',
        "analysis": (
            f"Revenue breakdown by product group for FY{current_fy}. "
            f"<strong>{top_group['category_l1']}</strong> is the largest category at "
            f"<strong>{top_group['mix_pct']:.1f}%</strong> of total revenue. "
            f"YoY change compares the same financial year period against FY{prev_fy}. "
            f"Groups with negative YoY growth warrant attention from the sales and product teams."
        ),
        "chart": {
            "type": "bar",
            "options": {},
            "labels": labels,
            "datasets": [
                {
                    "label": f"FY{current_fy} Revenue (R millions)",
                    "data": revenues,
                    "backgroundColor": colours,
                    "borderRadius": 4,
                }
            ],
        },
        "table": {
            "columns": [
                {"key": "group",   "label": "Product Group"},
                {"key": "revenue", "label": "Revenue (NET)"},
                {"key": "mix_pct", "label": "% of Mix"},
                {"key": "yoy",     "label": "YoY Change"},
            ],
            "rows": rows,
        },
    }
