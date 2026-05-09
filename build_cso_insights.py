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


def build_html(insights: list) -> str:
    insights_json = json.dumps(insights, ensure_ascii=False)
    today_fmt = datetime.now().strftime("%-d %B %Y") if sys.platform != "win32" else datetime.now().strftime("%#d %B %Y")

    return f"""<!DOCTYPE html>
<html lang="en" class="theme-dark">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CSO Insights — Olympic Paints</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@300;400;700;800;900&family=Barlow:wght@300;400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script>var t=localStorage.getItem('oly-theme');if(t)document.documentElement.className=t;</script>
<style>
/* ── RAW DESIGN TOKENS ─────────────────────────────────────────── */
:root{{
  --_y50:#FEF9E0;--_y100:#FDF0A0;--_y200:#FAE04D;--_y400:#F5C400;--_y600:#D4A800;--_y800:#A88000;--_y900:#6A5000;
  --_n50:#E8EFF8;--_n100:#B8CCE8;--_n300:#6B9ED0;--_n500:#2D6BA8;--_n700:#1A3D6E;--_n900:#0D2040;--_n950:#071022;
  --_g0:#FFFFFF;--_g50:#F7F6F3;--_g100:#E8E7E2;--_g200:#C8C7C0;--_g400:#949390;--_g600:#5C5B58;--_g800:#2E2E2C;--_g900:#1A1A18;--_g950:#0D0D0B;
  --_teal:#2D8C7A;--_teal-light:#C8EDE7;--_teal-dark:#1a5c50;
  --_coral:#E86060;--_coral-light:#FDDCDC;
  --font-display:'Barlow Condensed',sans-serif;--font-body:'Barlow',sans-serif;
  --r-sm:4px;--r-md:8px;--r-lg:12px;--r-xl:16px;--r-pill:50px;
}}
.theme-light{{color-scheme:light;--color-surface-page:var(--_g50);--color-surface-base:var(--_g0);--color-surface-elevated:var(--_g0);--color-surface-sunken:var(--_g100);--color-surface-overlay:rgba(0,0,0,0.04);--color-surface-brand:var(--_y400);--color-surface-secondary:var(--_n700);--color-text-primary:var(--_g950);--color-text-secondary:var(--_g600);--color-text-tertiary:var(--_g400);--color-text-on-brand:var(--_g950);--color-text-on-navy:var(--_g0);--color-brand-primary:var(--_y400);--color-brand-hover:var(--_y600);--color-brand-secondary:var(--_n700);--color-brand-accent:var(--_y400);--color-border-subtle:var(--_g100);--color-border-default:var(--_g200);--color-border-strong:var(--_g400);--color-border-brand:var(--_y400);--color-success-bg:#EDF7F5;--color-success-fg:var(--_teal-dark);--color-success-bd:var(--_teal);--color-warning-bg:var(--_y50);--color-warning-fg:var(--_y900);--color-warning-bd:var(--_y600);--color-danger-bg:#FEF2F2;--color-danger-fg:#C0392B;--color-danger-bd:var(--_coral);--color-info-bg:var(--_n50);--color-info-fg:var(--_n700);--color-info-bd:var(--_n500);--color-neutral-bg:var(--_g100);--color-neutral-fg:var(--_g600);--color-neutral-bd:var(--_g400);--shadow-sm:0 1px 3px rgba(0,0,0,0.08);--shadow-md:0 4px 12px rgba(0,0,0,0.08);--shadow-lg:0 10px 30px rgba(0,0,0,0.10);}}
.theme-dark{{color-scheme:dark;--color-surface-page:var(--_g950);--color-surface-base:var(--_g900);--color-surface-elevated:var(--_g800);--color-surface-sunken:var(--_g950);--color-surface-overlay:rgba(255,255,255,0.04);--color-surface-brand:var(--_y400);--color-surface-secondary:var(--_n700);--color-text-primary:var(--_g100);--color-text-secondary:var(--_g400);--color-text-tertiary:var(--_g600);--color-text-on-brand:var(--_g950);--color-text-on-navy:var(--_g0);--color-brand-primary:var(--_y400);--color-brand-hover:var(--_y200);--color-brand-secondary:var(--_n700);--color-brand-accent:var(--_y400);--color-border-subtle:rgba(255,255,255,0.06);--color-border-default:rgba(255,255,255,0.10);--color-border-strong:rgba(255,255,255,0.20);--color-border-brand:var(--_y400);--color-success-bg:rgba(45,140,122,0.12);--color-success-fg:var(--_teal-light);--color-success-bd:rgba(45,140,122,0.30);--color-warning-bg:rgba(245,196,0,0.10);--color-warning-fg:var(--_y200);--color-warning-bd:rgba(245,196,0,0.25);--color-danger-bg:rgba(232,96,96,0.12);--color-danger-fg:var(--_coral-light);--color-danger-bd:rgba(232,96,96,0.30);--color-info-bg:rgba(26,61,110,0.30);--color-info-fg:var(--_n100);--color-info-bd:rgba(107,158,208,0.30);--color-neutral-bg:rgba(255,255,255,0.05);--color-neutral-fg:var(--_g400);--color-neutral-bd:rgba(255,255,255,0.10);--shadow-sm:0 1px 3px rgba(0,0,0,0.40);--shadow-md:0 4px 12px rgba(0,0,0,0.40);--shadow-lg:0 10px 30px rgba(0,0,0,0.50);}}
.theme-brand{{color-scheme:light;--color-surface-page:var(--_y400);--color-surface-base:var(--_y200);--color-surface-elevated:var(--_y50);--color-surface-sunken:var(--_y600);--color-surface-overlay:rgba(0,0,0,0.05);--color-surface-brand:var(--_y400);--color-surface-secondary:var(--_g950);--color-text-primary:var(--_g950);--color-text-secondary:var(--_y900);--color-text-tertiary:var(--_y800);--color-text-on-brand:var(--_g950);--color-text-on-navy:var(--_g0);--color-brand-primary:var(--_g950);--color-brand-hover:var(--_n700);--color-brand-secondary:var(--_n700);--color-brand-accent:var(--_g950);--color-border-subtle:rgba(0,0,0,0.08);--color-border-default:rgba(0,0,0,0.14);--color-border-strong:rgba(0,0,0,0.25);--color-border-brand:var(--_g950);--color-success-bg:rgba(45,140,122,0.12);--color-success-fg:var(--_teal-dark);--color-success-bd:var(--_teal);--color-warning-bg:rgba(0,0,0,0.08);--color-warning-fg:var(--_y900);--color-warning-bd:var(--_y900);--color-danger-bg:rgba(232,96,96,0.12);--color-danger-fg:#C0392B;--color-danger-bd:var(--_coral);--color-info-bg:rgba(26,61,110,0.10);--color-info-fg:var(--_n900);--color-info-bd:var(--_n700);--color-neutral-bg:rgba(0,0,0,0.06);--color-neutral-fg:var(--_y900);--color-neutral-bd:rgba(0,0,0,0.15);--shadow-sm:0 1px 3px rgba(0,0,0,0.12);--shadow-md:0 4px 12px rgba(0,0,0,0.14);--shadow-lg:0 10px 30px rgba(0,0,0,0.18);}}
.theme-navy{{color-scheme:dark;--color-surface-page:var(--_n950);--color-surface-base:var(--_n900);--color-surface-elevated:var(--_n700);--color-surface-sunken:var(--_n950);--color-surface-overlay:rgba(255,255,255,0.04);--color-surface-brand:var(--_y400);--color-surface-secondary:var(--_n700);--color-text-primary:var(--_g0);--color-text-secondary:var(--_n100);--color-text-tertiary:var(--_n300);--color-text-on-brand:var(--_g950);--color-text-on-navy:var(--_g0);--color-brand-primary:var(--_y400);--color-brand-hover:var(--_y200);--color-brand-secondary:var(--_n500);--color-brand-accent:var(--_y400);--color-border-subtle:rgba(107,158,208,0.12);--color-border-default:rgba(107,158,208,0.20);--color-border-strong:rgba(107,158,208,0.35);--color-border-brand:var(--_y400);--color-success-bg:rgba(45,140,122,0.15);--color-success-fg:var(--_teal-light);--color-success-bd:rgba(45,140,122,0.35);--color-warning-bg:rgba(245,196,0,0.12);--color-warning-fg:var(--_y200);--color-warning-bd:rgba(245,196,0,0.30);--color-danger-bg:rgba(232,96,96,0.14);--color-danger-fg:var(--_coral-light);--color-danger-bd:rgba(232,96,96,0.35);--color-info-bg:rgba(45,107,168,0.20);--color-info-fg:var(--_n100);--color-info-bd:rgba(107,158,208,0.35);--color-neutral-bg:rgba(255,255,255,0.05);--color-neutral-fg:var(--_n300);--color-neutral-bd:rgba(255,255,255,0.12);--shadow-sm:0 1px 3px rgba(0,0,0,0.50);--shadow-md:0 4px 12px rgba(0,0,0,0.50);--shadow-lg:0 10px 30px rgba(0,0,0,0.60);}}

*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:100%;background:var(--color-surface-page);color:var(--color-text-primary);font-family:var(--font-body)}}

/* ── HEADER ── */
.site-header{{position:sticky;top:0;z-index:100;background:var(--color-surface-secondary);display:flex;align-items:center;justify-content:space-between;padding:10px 24px;gap:16px;border-bottom:1px solid var(--color-border-subtle)}}
.header-left{{display:flex;align-items:center;gap:12px}}
.logo-wrap{{width:48px;height:48px;border-radius:50%;overflow:hidden;flex-shrink:0}}
.logo-wrap img{{display:block;width:100%;height:100%;object-fit:cover}}
.site-title{{font-family:var(--font-display);font-weight:900;font-size:22px;letter-spacing:0.04em;text-transform:uppercase;color:var(--color-text-primary)}}
.site-date{{font-size:12px;color:var(--color-text-secondary);margin-top:2px}}
.theme-bar{{display:flex;gap:4px}}
.theme-bar button{{font-family:var(--font-body);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;padding:5px 12px;border:1px solid var(--color-border-default);border-radius:var(--r-pill);background:transparent;color:var(--color-text-secondary);cursor:pointer;transition:all .15s}}
.theme-bar button:hover{{border-color:var(--color-brand-primary);color:var(--color-brand-primary)}}
.theme-bar button.active{{background:var(--color-brand-primary);border-color:var(--color-brand-primary);color:var(--color-text-on-brand)}}

/* ── MAIN LAYOUT ── */
.main{{max-width:1200px;margin:0 auto;padding:32px 24px}}

/* ── GRID VIEW ── */
.eyebrow{{font-family:var(--font-display);font-weight:700;font-size:11px;letter-spacing:0.12em;text-transform:uppercase;color:var(--color-text-tertiary);margin-bottom:24px}}
.insight-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:20px}}
@media(max-width:900px){{.insight-grid{{grid-template-columns:repeat(2,1fr)}}}}
@media(max-width:560px){{.insight-grid{{grid-template-columns:1fr}}}}

.insight-card{{background:var(--color-surface-base);border:1px solid var(--color-border-default);border-left:4px solid var(--color-border-brand);border-radius:var(--r-lg);padding:22px 20px;cursor:pointer;transition:transform .18s,box-shadow .18s;display:flex;flex-direction:column;gap:12px}}
.insight-card:hover{{transform:translateY(-3px);box-shadow:var(--shadow-md)}}
.card-icon{{width:36px;height:36px;color:var(--color-brand-primary)}}
.card-icon svg{{width:100%;height:100%}}
.card-title{{font-family:var(--font-display);font-weight:700;font-size:18px;color:var(--color-text-primary)}}
.card-summary{{font-size:13px;color:var(--color-text-secondary);line-height:1.5}}
.card-updated{{font-size:11px;color:var(--color-neutral-fg);background:var(--color-neutral-bg);border:1px solid var(--color-neutral-bd);border-radius:var(--r-pill);padding:3px 10px;align-self:flex-start}}

/* ── DETAIL VIEW ── */
#detail-view{{display:none}}
.back-btn{{display:inline-flex;align-items:center;gap:6px;font-family:var(--font-body);font-size:13px;font-weight:600;color:var(--color-brand-primary);background:none;border:none;cursor:pointer;padding:0;margin-bottom:24px;transition:opacity .15s}}
.back-btn:hover{{opacity:.7}}
.detail-title{{font-family:var(--font-display);font-weight:900;font-size:40px;text-transform:uppercase;letter-spacing:0.02em;color:var(--color-text-primary);margin-bottom:8px}}
.detail-updated{{font-size:12px;color:var(--color-text-tertiary);margin-bottom:32px}}
.detail-section{{background:var(--color-surface-base);border:1px solid var(--color-border-default);border-radius:var(--r-lg);padding:24px;margin-bottom:20px}}
.section-heading{{font-family:var(--font-display);font-weight:800;font-size:13px;text-transform:uppercase;letter-spacing:0.10em;color:var(--color-text-tertiary);margin-bottom:14px}}
.analysis-text{{font-size:14px;color:var(--color-text-secondary);line-height:1.7}}
.chart-wrap{{position:relative;height:340px;margin-top:8px}}

/* ── TABLE ── */
.tbl-filter{{width:100%;padding:8px 12px;font-family:var(--font-body);font-size:13px;border:1px solid var(--color-border-default);border-radius:var(--r-md);background:var(--color-surface-sunken);color:var(--color-text-primary);margin-bottom:12px;outline:none}}
.tbl-filter:focus{{border-color:var(--color-border-brand)}}
.tbl-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
thead th{{font-family:var(--font-display);font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:0.08em;color:var(--color-text-tertiary);text-align:left;padding:10px 12px;border-bottom:1px solid var(--color-border-default);position:sticky;top:0;background:var(--color-surface-base)}}
tbody tr:nth-child(even){{background:var(--color-surface-sunken)}}
tbody td{{padding:9px 12px;color:var(--color-text-secondary);border-bottom:1px solid var(--color-border-subtle)}}
tbody tr:hover td{{color:var(--color-text-primary)}}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-left">
    <div class="logo-wrap">
      <img src="logo.jpg" alt="Olympic Paints" width="48" height="48">
    </div>
    <div>
      <div class="site-title">CSO Insights</div>
      <div class="site-date">{today_fmt}</div>
    </div>
  </div>
  <div class="theme-bar">
    <button onclick="olyTheme('theme-light',this)">Light</button>
    <button onclick="olyTheme('theme-dark',this)" class="active">Dark</button>
    <button onclick="olyTheme('theme-brand',this)">Brand</button>
    <button onclick="olyTheme('theme-navy',this)">Navy</button>
  </div>
</header>

<main class="main">

  <!-- GRID VIEW -->
  <div id="grid-view">
    <div class="eyebrow">Insights</div>
    <div class="insight-grid" id="insight-grid"></div>
  </div>

  <!-- DETAIL VIEW -->
  <div id="detail-view">
    <button class="back-btn" onclick="showGrid()">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
      Back to Insights
    </button>
    <div class="detail-title" id="detail-title"></div>
    <div class="detail-updated" id="detail-updated"></div>
    <div class="detail-section">
      <div class="section-heading">Analysis</div>
      <div class="analysis-text" id="detail-analysis"></div>
    </div>
    <div class="detail-section">
      <div class="section-heading">Chart</div>
      <div class="chart-wrap"><canvas id="detail-chart"></canvas></div>
    </div>
    <div class="detail-section">
      <div class="section-heading">Data</div>
      <input class="tbl-filter" id="tbl-filter" type="text" placeholder="Filter rows…" oninput="filterTable()">
      <div class="tbl-wrap"><table id="detail-table"><thead id="tbl-head"></thead><tbody id="tbl-body"></tbody></table></div>
    </div>
  </div>

</main>

<script>
const OLY_THEMES=['theme-light','theme-dark','theme-brand','theme-navy'];
function olyTheme(t,btn){{
  document.documentElement.classList.remove(...OLY_THEMES);
  document.documentElement.classList.add(t);
  localStorage.setItem('oly-theme',t);
  document.querySelectorAll('.theme-bar button').forEach(b=>b.classList.toggle('active',b===btn));
}}

// ── DATA ──────────────────────────────────────────────────────────
const INSIGHTS = {insights_json};

// ── barLabels plugin ──────────────────────────────────────────────
const _lblMap = new Map();
const barLabels = {{
  id:'barLabels',
  afterDatasetsDraw(chart){{
    const{{ctx,data}}=chart;
    chart.data.datasets.forEach((ds,di)=>{{
      const meta=chart.getDatasetMeta(di);
      meta.data.forEach((bar,i)=>{{
        const cfg=_lblMap.get(`${{di}}-${{i}}`);
        if(!cfg)return;
        const{{x,y,width,height}}=bar.getProps(['x','y','width','height'],true);
        ctx.save();
        ctx.font=`600 11px 'Barlow',sans-serif`;
        ctx.textAlign='center';
        ctx.textBaseline='middle';
        ctx.fillStyle=cfg.color||'#fff';
        ctx.fillText(cfg.text,cfg.inside?x:x,cfg.inside?y+height/2:y-8);
        ctx.restore();
      }});
    }});
  }}
}};
Chart.register(barLabels);

// ── CHART INSTANCE ────────────────────────────────────────────────
let _chart = null;

function renderChart(insight){{
  if(_chart){{ _chart.destroy(); _chart=null; }}
  _lblMap.clear();
  const canvas=document.getElementById('detail-chart');
  const cfg=insight.chart;
  const opts={{
    responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{labels:{{color:'var(--color-text-secondary)',font:{{family:'Barlow',size:12}}}}}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.dataset.label}}: ${{ctx.parsed.y??ctx.parsed.x}}`}}}}}},
    scales:{{
      x:{{ticks:{{color:'var(--color-text-tertiary)',font:{{family:'Barlow',size:11}}}},grid:{{color:'var(--color-border-subtle)'}}}},
      y:{{ticks:{{color:'var(--color-text-tertiary)',font:{{family:'Barlow',size:11}}}},grid:{{color:'var(--color-border-subtle)'}}}},
    }},
    ...cfg.options
  }};
  _chart=new Chart(canvas,{{type:cfg.type,data:{{labels:cfg.labels,datasets:cfg.datasets}},options:opts}});
}}

// ── TABLE ─────────────────────────────────────────────────────────
let _allRows=[];
function renderTable(insight){{
  const{{columns,rows}}=insight.table;
  _allRows=rows;
  const head=document.getElementById('tbl-head');
  const body=document.getElementById('tbl-body');
  head.innerHTML='<tr>'+columns.map(c=>`<th>${{c.label}}</th>`).join('')+'</tr>';
  document.getElementById('tbl-filter').value='';
  paintRows(rows,columns);
}}
function paintRows(rows,columns){{
  columns=columns||window._curInsight.table.columns;
  document.getElementById('tbl-body').innerHTML=rows.map(r=>
    '<tr>'+columns.map(c=>`<td>${{r[c.key]??''}}</td>`).join('')+'</tr>'
  ).join('');
}}
function filterTable(){{
  const q=document.getElementById('tbl-filter').value.toLowerCase();
  const cols=window._curInsight.table.columns;
  const filtered=_allRows.filter(r=>cols.some(c=>String(r[c.key]||'').toLowerCase().includes(q)));
  paintRows(filtered,cols);
}}

// ── NAVIGATION ────────────────────────────────────────────────────
function showGrid(){{
  document.getElementById('grid-view').style.display='';
  document.getElementById('detail-view').style.display='none';
  history.pushState(null,'',location.pathname);
}}

function showDetail(id){{
  const insight=INSIGHTS.find(i=>i.id===id);
  if(!insight)return;
  window._curInsight=insight;
  document.getElementById('grid-view').style.display='none';
  document.getElementById('detail-view').style.display='';
  document.getElementById('detail-title').textContent=insight.title;
  document.getElementById('detail-updated').textContent='Last updated: '+insight.updated;
  document.getElementById('detail-analysis').innerHTML=insight.analysis;
  renderChart(insight);
  renderTable(insight);
  history.pushState({{id}},'','#'+id);
  window.scrollTo(0,0);
}}

window.addEventListener('popstate',e=>{{
  if(e.state&&e.state.id) showDetail(e.state.id);
  else showGrid();
}});

// ── GRID RENDER ───────────────────────────────────────────────────
function renderGrid(){{
  const grid=document.getElementById('insight-grid');
  grid.innerHTML=INSIGHTS.map(i=>`
    <div class="insight-card" onclick="showDetail('${{i.id}}')">
      <div class="card-icon">${{i.icon}}</div>
      <div class="card-title">${{i.title}}</div>
      <div class="card-summary">${{i.summary}}</div>
      <div class="card-updated">Updated ${{i.updated}}</div>
    </div>
  `).join('');
}}

// ── INIT ──────────────────────────────────────────────────────────
renderGrid();
const initHash=location.hash.slice(1);
if(initHash)showDetail(initHash);

// Sync active theme button on load
(function(){{
  const t=document.documentElement.className;
  document.querySelectorAll('.theme-bar button').forEach(b=>{{
    b.classList.toggle('active',('theme-'+b.textContent.toLowerCase())===t||b.textContent.toLowerCase()===t.replace('theme-',''));
  }});
}})();
</script>
</body>
</html>"""
