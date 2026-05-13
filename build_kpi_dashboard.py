"""
build_kpi_dashboard.py
Builds the Olympic Paints Weekly KPI Dashboard from the Weekly Progress folder.
Data is sourced from QuickSight-generated PDFs in the Weekly Progress folder.
Because QuickSight PDFs render charts as images (not extractable text), the key
figures are maintained here as structured data and updated weekly.

Run manually or via Windows Task Scheduler whenever new weekly reports arrive.
"""

import os
import json
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path

import pandas as pd

BASE_DIR        = Path(__file__).parent
WEEKLY_DIR      = BASE_DIR.parent.parent / "1.Projects" / "KPI Report" / "Weekly Progress"
DASHBOARD       = BASE_DIR / "KPI Dashboard.html"
INDEX           = BASE_DIR / "index.html"
WORKSPACE_DASH  = Path(r"C:\Users\quint\workspace-dashboard")
MERCH_FILE      = BASE_DIR / "zoho_meetings" / "data" / "Meetings_Report_Merchandising.xlsx"  # Migrated 2026-05-13 from Zoho/Meetings_Report_AWS_Merchandising.xlsx (now direct Zoho REST API)
LEADS_FILE      = BASE_DIR / "zoho_meetings" / "data" / "OP_Lead_Tracking_new.csv"  # Migrated 2026-05-13 from Zoho/OP_Lead_Tracking_new.csv (now direct Zoho REST API)

# KPI F "Merchandising" — the full 10% bucket is now scored on merchandising visits.
# Training is excluded from scoring until a data source exists.
# Scoring is GRADED: score = min(visits/target, 1) × 10%.
MERCH_TARGET_PER_MONTH       = 15        # visits per rep per month
MERCH_WEIGHT_OF_F            = 10.0      # full 10% — training portion no longer scored
NAME_TO_REP_CODE = {"NIKHIL":"NP","BYRON":"BM","ABOO":"AC","AMIT":"AP","BHADRESH":"BV"}
LEAD_OWNER_TO_CODE = {
    "aboo cassim": "AC", "amit patel": "AP", "bhadresh vallabh": "BV",
    "nikhil panchal": "NP", "byron minnie": "BM",
}
KPI_B_TARGET = 5  # min leads/month — hit or miss
MONTH_FULL_NAMES = ["January","February","March","April","May","June",
                    "July","August","September","October","November","December"]

# ── WEEKLY DATA ───────────────────────────────────────────────────────────────
# Update this block each week from the QuickSight Weekly Sales Report PDFs.
# Source: Weekly Progress folder — week ending 1 May 2026

REPORT_WEEK   = "1 May 2026"
REPORT_DATE   = "2026-05-01"

# ── Sales & Target (from Weekly_Sales_Report__2026-05-01T04_50_35.pdf) ──────
MTD_SALES      = 12_796_044.69
MTD_TARGET     = 12_814_305.26
MTD_PCT_TARGET = -0.14   # % vs target (negative = below)

# ── Debtors ──────────────────────────────────────────────────────────────────
# Source: Weekly_Sales_Report__2026-05-01T04_50_35.pdf
DEBTORS_TOTAL  = 20_922_618.50
DEBTORS_90D    = 4_495_542.11
OVERDUE_60D_PCT = 21.49   # % overdue > 60 days

# ── Margin ───────────────────────────────────────────────────────────────────
ABOVE_RB_AVG   = 9.75    # company average rock bottom % (store avg: 4.25%)
ABOVE_RB_TARGET = 15.0   # target

# ── Rep MTD Performance ───────────────────────────────────────────────────────
REPS = [
    {"code": "AC", "name": "Aboo Cassim",     "sales": 1_626_404.30, "target": 1_160_055.78, "pct": 28.67,  "rb_pct": None,  "q2_target": 2_756_633.76,  "yoy": None,   "orders_approved": 152_469},
    {"code": "AP", "name": "Amit Patel",      "sales": 1_996_047.54, "target": 1_582_503.57, "pct": 20.72,  "rb_pct": 8.85,  "q2_target": 2_336_788.38,  "yoy": None,   "orders_approved": 320_042},
    {"code": "BV", "name": "Bhadresh Vallabh","sales": 4_411_457.73, "target": 5_131_146.02, "pct": -16.31, "rb_pct": None,  "q2_target": 10_753_757.53, "yoy": None,   "orders_approved": 1_568_153},
    {"code": "NP", "name": "Nikhil Panchal",  "sales": 4_738_186.05, "target": 4_940_599.89, "pct": -4.27,  "rb_pct": 8.82,  "q2_target": 9_896_452.05,  "yoy": None,   "orders_approved": 977_024},
    {"code": "BM", "name": "Byron Minnie",    "sales": 23_949.07,    "target": None,          "pct": None,   "rb_pct": None,  "q2_target": None,           "yoy": None,   "orders_approved": None},
]

# ── YOY Monthly Comparison ────────────────────────────────────────────────────
YOY = [
    {"month": "Jan", "sales_2025": 7_297_999.94, "sales_2026": 7_617_705.29, "yoy_pct": 4.38},
    {"month": "Feb", "sales_2025": 8_251_126.12, "sales_2026": 9_875_474.09, "yoy_pct": 19.69},
    {"month": "Mar", "sales_2025": 11_342_188.94,"sales_2026": 10_504_834.10,"yoy_pct": -7.38},
    {"month": "Apr", "sales_2025": 9_857_157.89, "sales_2026": 12_796_044.69,"yoy_pct": 29.81},
]

# ── Product Mix by Revenue % ──────────────────────────────────────────────────
PRODUCT_MIX = [
    {"group": "Enamel",     "pct": 31, "avg_sale": 133_140},
    {"group": "QD Enamel",  "pct": 29, "avg_sale": 124_130},
    {"group": "Accessories","pct": 13, "avg_sale": 54_890},
    {"group": "Sealers",    "pct": 10, "avg_sale": 42_100},
    {"group": "PVA",        "pct": 7,  "avg_sale": 31_340},
    {"group": "Roof",       "pct": 6,  "avg_sale": 25_750},
    {"group": "Primer",     "pct": 1,  "avg_sale": 3_630},
]

# ── Rock Bottom % by Product Group ────────────────────────────────────────────
RB_BY_PRODUCT = [
    {"group": "Ultimate Shine",      "rb_pct": -24.30},
    {"group": "Membrane",            "rb_pct": -23.69},
    {"group": "Etch Primer",         "rb_pct": -8.63},
    {"group": "Rust Remover",        "rb_pct": -6.81},
    {"group": "Distemper",           "rb_pct": -6.15},
    {"group": "LIBERTY",             "rb_pct": -4.82},
    {"group": "Wood Primer",         "rb_pct": -4.52},
    {"group": "7 in 1 PVA",          "rb_pct": -3.94},
    {"group": "Kalahari Contractors", "rb_pct": -3.77},
    {"group": "All In One",           "rb_pct": -2.72},
    {"group": "Hi Hiding Contr",      "rb_pct": -1.73},
    {"group": "Oxide",                "rb_pct": -1.48},
    {"group": "Eclipse PVA",          "rb_pct": -0.08},
    {"group": "Road Marking",         "rb_pct": 0.63},
    {"group": "Rugged Beauty",        "rb_pct": 0.92},
    {"group": "RainProof",            "rb_pct": 1.08},
    {"group": "Fibre Restore",        "rb_pct": 1.27},
    {"group": "Master Decorators",    "rb_pct": 1.31},
    {"group": "High Gloss",           "rb_pct": 1.65},
    {"group": "Primer",               "rb_pct": 1.87},
    {"group": "Putty",                "rb_pct": 2.57},
    {"group": "Thinner",              "rb_pct": 2.69},
    {"group": "Sanding Sealer",       "rb_pct": 3.23},
    {"group": "Universal Roof Paint", "rb_pct": 3.23},
    {"group": "Plaster n Tile Bond",  "rb_pct": 3.38},
    {"group": "Stainer",              "rb_pct": 3.85},
    {"group": "Suburban Bliss",       "rb_pct": 3.95},
    {"group": "Zinc Phosphate",       "rb_pct": 4.98},
    {"group": "Pick & Save",          "rb_pct": 5.18},
    {"group": "Wood Varnish",       "rb_pct": 5.22},
    {"group": "Natural Elegance",   "rb_pct": 5.33},
    {"group": "Decor",              "rb_pct": 5.34},
    {"group": "Universal Undercoat","rb_pct": 5.45},
    {"group": "Roof & Stoep",       "rb_pct": 5.60},
    {"group": "FB Dressing",        "rb_pct": 6.33},
    {"group": "Bonding Liquid",     "rb_pct": 6.65},
    {"group": "3 in 1 Roof",        "rb_pct": 6.80},
    {"group": "Plaster Primer",     "rb_pct": 7.16},
    {"group": "3 in 1 Gripcoat",    "rb_pct": 7.70},
    {"group": "Q.D. Primer",        "rb_pct": 9.84},
    {"group": "Carbolineum",        "rb_pct": 10.57},
]

# ── Customers at Risk (below rock bottom) ─────────────────────────────────────
CUSTOMERS_AT_RISK = [
    {"name": "Cassim Tayob",         "rb_pct": -37.14},
    {"name": "Patels Hardware",      "rb_pct": -17.15},
    {"name": "Dainty's Wholesale",   "rb_pct": -12.79},
    {"name": "Dada's World of Hard.","rb_pct": -11.78},
    {"name": "HGM Steelboys CC",     "rb_pct": -2.09},
    {"name": "Brits Hardware & Gla.","rb_pct": 4.12},
    {"name": "Del Piero Trading CC", "rb_pct": 1.58},
    {"name": "Myka Trade 2",         "rb_pct": 5.30},
]

# ── E-Commerce Orders (from Daily_Eccomerce_Repo_2026-04-21) ─────────────────
ECOM_TOTAL_ORDERS = 69
ECOM_TOTAL_QTY    = 123
ECOM_OVERDUE      = 0   # per "Overdue Status" column (0 flagged as formally overdue)
ECOM_AGING = {
    "1-5 days":   12,
    "6-9 days":   21,
    "10-14 days": 14,
    "15+ days":   22,
}  # approximate from the data (19 orders at 19 days, multiple at 15+, etc.)

# ── KPI Agreement Categories (6-category framework) ───────────────────────────
KPI_CATEGORIES = [
    {"id": "A", "name": "Sales Growth & Collections", "weight": 30,
     "description": "Must exceed 10% YoY growth. Commission on sales above prior × 1.10."},
    {"id": "B", "name": "Discount Management",        "weight": 20,
     "description": "Average markup above Rock Bottom — banded: ≥15%=20pts, ≥14%=16, ≥13%=12, ≥12%=8, ≥11%=4, below=0."},
    {"id": "C", "name": "Customer Development & CRM", "weight": 20,
     "description": "Min 5 new leads created in CRM per month (15/quarter). Hit or Miss."},
    {"id": "D", "name": "Product Development",        "weight": 10,
     "description": "Min 1 upsell/week (4-5/month) of focus products."},
    {"id": "E", "name": "New Customer Onboarding",    "weight": 10,
     "description": "Min 2 new verified trading customers per month (6/quarter full, 3/quarter half)."},
    {"id": "F", "name": "Merchandising",              "weight": 10,
     "description": "Merchandising visits — graded vs 15 visits/rep/month (45/quarter). Training excluded until data feed exists."},
]

# ── Discount Management banding (KPI B) ───────────────────────────────────────
# (markup_pct_threshold, points_awarded) — same as rep/sales dashboards
RB_BANDS = [(15, 20), (14, 16), (13, 12), (12, 8), (11, 4)]

# ── HELPERS ────────────────────────────────────────────────────────────────────

def fmt_r(val, decimals=0):
    if val is None: return "—"
    if val >= 1_000_000: return f"R{val/1_000_000:.2f}M"
    if val >= 1_000:     return f"R{val/1_000:.1f}K"
    return f"R{val:.{decimals}f}"

def pct_str(val):
    if val is None: return "—"
    return f"{val:+.1f}%" if val != 0 else "0.0%"

def pct_plain(val):
    if val is None: return "—"
    return f"{val:.1f}%"

def js_arr(vals):
    return "[" + ",".join("null" if v is None else str(round(v, 2)) for v in vals) + "]"

def rb_color_class(pct):
    if pct is None:  return "neutral"
    if pct >= 8:     return "green"
    if pct >= 0:     return "amber"
    return "red"

def sales_color_class(pct):
    if pct is None:  return "neutral"
    if pct >= 0:     return "green"
    if pct >= -15:   return "amber"
    return "red"

def _kpi_b_discount_pill(rb_pct):
    """KPI B — Discount Management: banded markup above Rock Bottom."""
    if rb_pct is None:
        return '<span class="pill neutral" title="Discount Management: per-rep rock bottom % not available">⚠ No data</span>'
    score = 0
    for thr, pts in RB_BANDS:
        if rb_pct >= thr:
            score = pts; break
    if score == 20:
        cls, icon = "green", "✓"
    elif score > 0:
        cls, icon = "amber", "⚠"
    else:
        cls, icon = "red", "✗"
    title = f"Discount Mgmt: +{rb_pct:.1f}% avg above RB → {score}/20 pts (target ≥11%)"
    return f'<span class="pill {cls}" title="{title}">{icon} +{rb_pct:.1f}% &middot; {score}/20</span>'

def _kpi_c_pill(count, period):
    title = f"Customer Dev / CRM: {count} leads created in {period} (target ≥{KPI_B_TARGET}/month)"
    if count >= KPI_B_TARGET:
        return f'<span class="pill green" title="{title}">✓ Hit &middot; {count} leads</span>'
    if count > 0:
        return f'<span class="pill red" title="{title}">✗ Miss &middot; {count}/{KPI_B_TARGET}</span>'
    return f'<span class="pill red" title="{title}">✗ Miss &middot; 0 leads</span>'

# ── MERCHANDISING DATA (KPI E — 70% of 10%) ───────────────────────────────────

def load_merch_visits_by_month():
    """Parse Zoho's Meetings_Report_AWS_Merchandising export.

    Rep is buried in the free-text Note Content as
    'WHO IS THE REP THAT SERVICES THE STORE: <FIRSTNAME>'.
    Returns dict keyed by (year, month) -> {rep_code: visit_count}.
    Both PLANNED and VISITED rows count (per KPI agreement).
    """
    if not MERCH_FILE.exists():
        return {}
    df = pd.read_excel(MERCH_FILE, header=6)
    rep_re = re.compile(r"WHO IS THE REP[^:]*:\s*\n?\s*([A-Z][A-Z\s]*)", re.IGNORECASE)
    out = defaultdict(lambda: defaultdict(int))
    for _, row in df.iterrows():
        note = str(row.get("Note Content", "") or "")
        m = rep_re.search(note)
        if not m:
            continue
        first = m.group(1).strip().upper().split()[0]
        code = NAME_TO_REP_CODE.get(first)
        if not code:
            continue
        dt = row.get("Created Time")
        if pd.isna(dt):
            continue
        if isinstance(dt, str):
            dt = pd.to_datetime(dt, errors="coerce")
            if pd.isna(dt):
                continue
        out[(dt.year, dt.month)][code] += 1
    return {k: dict(v) for k, v in out.items()}


def load_lead_surveys_by_month():
    """Count prospect-store surveys per rep per month from leads.parquet.

    A "lead survey" is a Lead record where Lead_Quality has been graded
    (Good / Medium / Bad). That's the explicit "rep assessed this prospect store"
    action — separate from lead creation (which Lelani does as recon) and from
    formal Merchandising Visit events.

    Bucketed by Modified_Time (when the grade was set), not Created_Time.
    Returns dict keyed by (year, month) -> {rep_code: survey_count}.
    """
    leads_pq = BASE_DIR / "zoho_meetings" / "data" / "leads.parquet"
    if not leads_pq.exists():
        return {}
    df = pd.read_parquet(leads_pq)
    if "Lead_Quality" not in df.columns or "Owner_name" not in df.columns:
        return {}

    def _qual(v):
        if isinstance(v, dict):
            return v.get("name") or v.get("display_value")
        if v is None or v == "":
            return None
        try:
            if pd.isna(v):
                return None
        except (TypeError, ValueError):
            pass
        return str(v)

    graded = df.copy()
    graded["_q"] = graded["Lead_Quality"].apply(_qual)
    graded = graded[graded["_q"].notna()]
    if not len(graded):
        return {}
    graded["Modified_Time"] = pd.to_datetime(graded["Modified_Time"], errors="coerce", utc=True)
    graded = graded[graded["Modified_Time"].notna()].copy()

    out = defaultdict(lambda: defaultdict(int))
    for _, row in graded.iterrows():
        owner_first = str(row["Owner_name"]).strip().split()[0].upper() if row["Owner_name"] else ""
        code = NAME_TO_REP_CODE.get(owner_first)
        if not code:
            continue
        dt = row["Modified_Time"]
        out[(dt.year, dt.month)][code] += 1
    return {k: dict(v) for k, v in out.items()}


def _merge_counts(a: dict, b: dict) -> dict:
    """Merge two {rep_code: count} dicts by summing values."""
    out = dict(a)
    for k, v in b.items():
        out[k] = out.get(k, 0) + v
    return out


def get_merch_scoring_window():
    """Return (period_label, fallback_from, visits_dict, breakdown_dict).

    Combines two merch activity streams:
      • Tag="Merchandising Visit" Event records (formal visits, ~87)
      • Leads where Lead_Quality has been graded (prospect surveys, ~95)

    Default window = last completed calendar month. If that month has no logged
    activity from any rep across either stream, fall back to the most recent
    month that does have data.

    Returns:
        period_label   : month being scored (e.g. 'April 2026')
        fallback_from  : label of month requested if we fell back, else None
        combined_by_rep: {rep_code: visits + surveys}
        breakdown      : {rep_code: {'visits': N, 'surveys': N}}
    """
    visits = load_merch_visits_by_month()
    surveys = load_lead_surveys_by_month()

    today = date.today()
    last_y = today.year if today.month > 1 else today.year - 1
    last_m = today.month - 1 if today.month > 1 else 12
    target_key = (last_y, last_m)
    target_label = f"{MONTH_FULL_NAMES[last_m-1]} {last_y}"

    def _at(key):
        v = visits.get(key, {})
        s = surveys.get(key, {})
        combined = _merge_counts(v, s)
        breakdown = {code: {"visits": v.get(code, 0), "surveys": s.get(code, 0)}
                     for code in set(v) | set(s)}
        return combined, breakdown

    combined, breakdown = _at(target_key)
    if combined:
        return target_label, None, combined, breakdown

    all_keys = sorted(set(visits) | set(surveys), reverse=True)
    candidates = [k for k in all_keys if k <= target_key]
    if candidates:
        fk = candidates[0]
        combined, breakdown = _at(fk)
        return f"{MONTH_FULL_NAMES[fk[1]-1]} {fk[0]}", target_label, combined, breakdown
    return target_label, None, {}, {}

# ── LEADS DATA (KPI B — Customer Development & CRM) ───────────────────────────

_ZOHO_DT_RE = re.compile(r"\w+\s+(\w+)\s+(\d+)\s+(\d{4})\s+(\d+:\d+:\d+)")

def load_leads_by_month():
    """Count leads per rep per month from OP_Lead_Tracking_new.csv.

    Zoho exports timestamps as "Wed May 06 2026 07:57:00 GMT+0000 (...)" —
    non-standard, so we extract MonthName Day Year HH:MM:SS with a regex.
    Returns dict keyed by (year, month) -> {rep_code: lead_count}.
    """
    if not LEADS_FILE.exists():
        return {}
    df = pd.read_csv(LEADS_FILE)
    out = defaultdict(lambda: defaultdict(int))
    for _, row in df.iterrows():
        owner = str(row.get("Lead Owner", "") or "").strip().lower()
        code = LEAD_OWNER_TO_CODE.get(owner)
        if not code:
            continue
        raw_dt = str(row.get("Created Time", "") or "")
        m = _ZOHO_DT_RE.search(raw_dt)
        if not m:
            continue
        month_name, day, year, time_part = m.groups()
        try:
            dt = datetime.strptime(f"{month_name} {day} {year} {time_part}", "%b %d %Y %H:%M:%S")
        except ValueError:
            continue
        out[(dt.year, dt.month)][code] += 1
    return {k: dict(v) for k, v in out.items()}


def get_leads_scoring_window():
    """Return (period_label, fallback_from, leads_dict).

    Uses last completed calendar month. Falls back to most recent month with data.
    """
    leads = load_leads_by_month()
    today = date.today()
    last_y = today.year if today.month > 1 else today.year - 1
    last_m = today.month - 1 if today.month > 1 else 12
    target_key = (last_y, last_m)
    target_label = f"{MONTH_FULL_NAMES[last_m-1]} {last_y}"
    if leads.get(target_key):
        return target_label, None, leads[target_key]
    candidates = sorted([k for k in leads if k <= target_key], reverse=True)
    if candidates:
        fk = candidates[0]
        return f"{MONTH_FULL_NAMES[fk[1]-1]} {fk[0]}", target_label, leads[fk]
    return target_label, None, {}


# ── HTML BUILD ─────────────────────────────────────────────────────────────────

def build_html() -> str:
    generated = datetime.now().strftime("%d %B %Y %H:%M")

    merch_period, merch_fallback_from, merch_by_rep, merch_breakdown = get_merch_scoring_window()
    leads_period, leads_fallback_from, leads_by_rep = get_leads_scoring_window()

    # Rep table rows
    rep_rows = ""
    for r in REPS:
        if r["target"] is None:
            pct_bar = '<span class="pill neutral">Internal</span>'
            pct_cell = "—"
        else:
            cls = sales_color_class(r["pct"])
            pct_cell = pct_str(r["pct"])
            pct_bar = f'<span class="pill {cls}">{pct_cell}</span>'

        rb = r["rb_pct"]
        rb_cls = rb_color_class(rb)
        rb_cell = f'<span class="pill {rb_cls}">{pct_plain(rb)}</span>' if rb is not None else '<span class="pill neutral">No data</span>'

        orders = fmt_r(r["orders_approved"])

        # KPI A scoring (sales growth vs prior year baseline)
        yoy = r["yoy"]
        if yoy is None:
            kpi_a = '<span class="pill neutral">—</span>'
        elif yoy >= 10:
            kpi_a = f'<span class="pill green">✓ {pct_str(yoy)} YOY</span>'
        elif yoy >= 0:
            kpi_a = f'<span class="pill amber">⚠ {pct_str(yoy)} YOY</span>'
        else:
            kpi_a = f'<span class="pill red">✗ {pct_str(yoy)} YOY</span>'

        # KPI F (Merchandising) — graded score from formal merch visits + lead surveys.
        # Score = min(activity/target, 1) × 10%. Training is no longer scored.
        merch_count   = merch_by_rep.get(r["code"], 0)
        merch_pct     = min(merch_count / MERCH_TARGET_PER_MONTH, 1.0) * 100  # achievement %
        merch_score   = round(min(merch_count / MERCH_TARGET_PER_MONTH, 1.0) * MERCH_WEIGHT_OF_F, 1)
        _brk          = merch_breakdown.get(r["code"], {"visits": 0, "surveys": 0})
        merch_title   = (f"Merchandising: {merch_count}/{MERCH_TARGET_PER_MONTH} total in {merch_period} "
                         f"({_brk['visits']} visits + {_brk['surveys']} lead surveys) "
                         f"→ {merch_pct:.0f}% of target → {merch_score:.1f} of 10 pts")
        if merch_pct >= 100:
            kpi_f = f'<span class="pill green" title="{merch_title}">✓ {merch_pct:.0f}% &middot; {merch_count}/{MERCH_TARGET_PER_MONTH}</span>'
        elif merch_pct >= 50:
            kpi_f = f'<span class="pill amber" title="{merch_title}">⚠ {merch_pct:.0f}% &middot; {merch_count}/{MERCH_TARGET_PER_MONTH}</span>'
        else:
            kpi_f = f'<span class="pill red" title="{merch_title}">✗ {merch_pct:.0f}% &middot; {merch_count}/{MERCH_TARGET_PER_MONTH}</span>'

        rep_rows += f"""
          <tr>
            <td><strong>{r["code"]}</strong><br><small style="color:var(--muted)">{r["name"]}</small></td>
            <td>{fmt_r(r["sales"])}</td>
            <td>{fmt_r(r["target"])}</td>
            <td>{pct_bar}</td>
            <td>{rb_cell}</td>
            <td>{kpi_a}</td>
            <td>{_kpi_b_discount_pill(r["rb_pct"])}</td>
            <td>{_kpi_c_pill(leads_by_rep.get(r["code"], 0), leads_period)}</td>
            <td><span class="pill neutral" title="Product dev data required">⚠ No data</span></td>
            <td><span class="pill neutral" title="New customer data required">⚠ No data</span></td>
            <td>{kpi_f}</td>
            <td>{orders}</td>
          </tr>"""

    # KPI Achievement Summary card — per-rep merchandising roll-up.
    # Graded scoring: score_earned = min(visits/target, 1) × 10% (full bucket, training excluded).
    merch_rows_html = ""
    team_visits = 0
    team_surveys = 0
    team_target = MERCH_TARGET_PER_MONTH * len(REPS)
    team_score_earned = 0.0
    for r in REPS:
        c = merch_by_rep.get(r["code"], 0)
        brk = merch_breakdown.get(r["code"], {"visits": 0, "surveys": 0})
        team_visits  += brk["visits"]
        team_surveys += brk["surveys"]
        ach_ratio  = min(c / MERCH_TARGET_PER_MONTH, 1.0)
        ach_pct    = ach_ratio * 100
        score_earned = round(ach_ratio * MERCH_WEIGHT_OF_F, 1)
        team_score_earned += score_earned
        if ach_pct >= 100:
            pill_cls, pill_lbl = "green", f"✓ {ach_pct:.0f}%"
            bar_color = "var(--green)"
        elif ach_pct >= 50:
            pill_cls, pill_lbl = "amber", f"⚠ {ach_pct:.0f}%"
            bar_color = "var(--amber)"
        else:
            pill_cls, pill_lbl = "red", f"✗ {ach_pct:.0f}%"
            bar_color = "var(--red)"
        breakdown_html = f"<div style='font-size:11px;color:var(--muted);margin-top:2px'>{brk['visits']} visits &middot; {brk['surveys']} lead surveys</div>"
        merch_rows_html += f"""
          <tr>
            <td><strong>{r["code"]}</strong> &mdash; {r["name"]}</td>
            <td style="text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:800;font-size:18px">{c}{breakdown_html}</td>
            <td style="text-align:right;color:var(--muted)">{MERCH_TARGET_PER_MONTH}</td>
            <td style="min-width:120px"><div style="background:#e8e7e2;border-radius:4px;height:8px;overflow:hidden"><div style="width:{ach_pct:.0f}%;height:100%;background:{bar_color}"></div></div><div style="font-size:11px;color:var(--muted);margin-top:3px">{ach_pct:.0f}% of target</div></td>
            <td><span class="pill {pill_cls}">{pill_lbl}</span></td>
            <td style="text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:700">{score_earned:.1f}% / {MERCH_WEIGHT_OF_F:.0f}%</td>
          </tr>"""
    team_visits_total = team_visits + team_surveys  # for the "Team N%" pill
    overall_pct = round(team_visits_total / team_target * 100) if team_target else 0
    avg_score   = round(team_score_earned / len(REPS), 1) if REPS else 0.0
    if overall_pct >= 100: overall_pill, overall_lbl = "green", f"✓ Team {overall_pct}%"
    elif overall_pct >= 50: overall_pill, overall_lbl = "amber", f"⚠ Team {overall_pct}%"
    else: overall_pill, overall_lbl = "red", f"✗ Team {overall_pct}%"
    fallback_note = (f' &middot; <span style="color:var(--amber)"><strong>Note:</strong> '
                     f'{merch_fallback_from} had no logged visits — showing {merch_period} instead.</span>'
                     ) if merch_fallback_from else ""
    merch_summary_card = f"""
  <div class="card full">
    <div class="card-title">Merchandising Achievement &mdash; {merch_period}</div>
    <div class="card-sub">
      Source: <em>formal merch visits</em> (Tag = "Merchandising Visit") + <em>lead surveys</em> (Lead_Quality graded), both pulled from Zoho REST API &middot; target: {MERCH_TARGET_PER_MONTH} activities / rep / month &middot;
      <strong>Graded scoring:</strong> score = min(activity/target, 1) × 10% &middot; full 10% of KPI F (training excluded){fallback_note}
    </div>
    <div class="tw" style="margin-top:14px">
      <table>
        <thead>
          <tr>
            <th>Rep</th>
            <th style="text-align:right">Visits</th>
            <th style="text-align:right">Target</th>
            <th>Achievement</th>
            <th>Status</th>
            <th style="text-align:right">Score Earned</th>
          </tr>
        </thead>
        <tbody>{merch_rows_html}
          <tr style="background:#f7f6f3;font-weight:700">
            <td><strong>TEAM TOTAL</strong></td>
            <td style="text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:900;font-size:20px">{team_visits}</td>
            <td style="text-align:right">{team_target}</td>
            <td><div style="background:#e8e7e2;border-radius:4px;height:8px;overflow:hidden"><div style="width:{min(overall_pct,100):.0f}%;height:100%;background:var(--gold)"></div></div><div style="font-size:11px;color:var(--muted);margin-top:3px">{overall_pct}% of team target</div></td>
            <td><span class="pill {overall_pill}">{overall_lbl}</span></td>
            <td style="text-align:right;font-family:'Barlow Condensed',sans-serif;font-weight:900">avg {avg_score:.1f}% / {MERCH_WEIGHT_OF_F:.0f}%</td>
          </tr>
        </tbody>
      </table>
    </div>
    <p style="margin-top:12px;font-size:11px;color:var(--muted)">
      KPI E is now scored on merchandising visits only — graded against a 15 visits/rep/month target.
      Training is excluded from scoring until a training register feed exists.
    </p>
  </div>"""

    # YOY chart arrays
    yoy_labels = js_arr(None for _ in [])
    months = [y["month"] for y in YOY]
    s25 = [y["sales_2025"] for y in YOY]
    s26 = [y["sales_2026"] for y in YOY]
    yoy_pcts = [y["yoy_pct"] for y in YOY]
    months_js = "[" + ",".join(f'"{m}"' for m in months) + "]"
    s25_js = js_arr(s25)
    s26_js = js_arr(s26)
    yoy_pct_js = js_arr(yoy_pcts)

    # Rep chart arrays
    rep_names_js = "[" + ",".join(f'"{r["name"]}"' for r in REPS) + "]"
    rep_sales_js = js_arr([r["sales"] for r in REPS])
    rep_target_js = js_arr([r["target"] for r in REPS])

    # RB product chart (top 15 worst + best, sorted)
    rb_sorted = sorted(RB_BY_PRODUCT, key=lambda x: x["rb_pct"])
    rb_display = rb_sorted[:10] + rb_sorted[-8:]  # worst 10 + best 8
    rb_labels_js = "[" + ",".join(f'"{x["group"]}"' for x in rb_display) + "]"
    rb_vals_js = js_arr([x["rb_pct"] for x in rb_display])
    rb_colors_js = "[" + ",".join(
        f'"#E63946"' if x["rb_pct"] < 0 else (f'"#F4A261"' if x["rb_pct"] < 8 else f'"#2DC653"')
        for x in rb_display
    ) + "]"

    # Product mix
    pm_labels_js = "[" + ",".join(f'"{p["group"]}"' for p in PRODUCT_MIX) + "]"
    pm_vals_js = js_arr([p["pct"] for p in PRODUCT_MIX])

    # Q2 tracking
    q2_reps = [r for r in REPS if r["q2_target"]]
    q2_names_js = "[" + ",".join(f'"{r["name"]}"' for r in q2_reps) + "]"
    q2_target_js = js_arr([r["q2_target"] for r in q2_reps])
    q2_sales_js = js_arr([r["sales"] for r in q2_reps])

    # Customer risk rows
    cust_rows = ""
    for c in CUSTOMERS_AT_RISK:
        cls = rb_color_class(c["rb_pct"])
        icon = "✗" if c["rb_pct"] < 0 else "⚠"
        cust_rows += f"""
          <tr>
            <td>{c["name"]}</td>
            <td><span class="pill {cls}">{icon} {pct_plain(c["rb_pct"])}</span></td>
          </tr>"""

    # RB product worst rows (below 0)
    rb_risk_rows = ""
    for p in sorted([x for x in RB_BY_PRODUCT if x["rb_pct"] < 0], key=lambda x: x["rb_pct"]):
        cls = "red"
        rb_risk_rows += f"""
          <tr>
            <td>{p["group"]}</td>
            <td><span class="pill {cls}">✗ {pct_plain(p["rb_pct"])}</span></td>
          </tr>"""

    # Overall attainment (reps with targets only)
    scored_reps = [r for r in REPS if r["target"] is not None]
    avg_attainment = sum(r["sales"]/r["target"]*100 for r in scored_reps) / len(scored_reps)
    total_target = sum(r["target"] for r in scored_reps)
    total_sales = sum(r["sales"] for r in scored_reps)
    total_pct = (total_sales / total_target - 1) * 100

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Olympic Paints — Weekly KPI Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --orange: #E85D04; --dark: #1A1A2E; --mid: #2D2D44; --light: #F5F5F0;
    --gold: #F4A261; --teal: #2EC4B6; --red: #E63946; --green: #2DC653;
    --amber: #F4A261; --card: #FFFFFF; --border: #E8E8E0; --muted: #6B7280;
    --warn-bg: #FFF7E6; --warn-border: #F4A261;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: var(--light); color: var(--dark); }}

  .header {{ background: linear-gradient(135deg, var(--dark), var(--mid)); color: #fff;
    padding: 22px 40px; display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 20px rgba(0,0,0,.3); position: sticky; top: 0; z-index: 100; }}
  .header-left {{ display: flex; align-items: center; gap: 18px; }}
  .logo {{ width: 52px; height: 52px; background: var(--orange); border-radius: 50%;
    display: flex; align-items: center; justify-content: center; font-size: 20px;
    font-weight: 900; color: #fff; box-shadow: 0 0 0 3px rgba(232,93,4,.3); }}
  .header-title {{ font-size: 21px; font-weight: 700; }}
  .header-sub {{ font-size: 12px; color: rgba(255,255,255,.55); margin-top: 2px; }}
  .header-right {{ text-align: right; }}
  .header-week {{ font-size: 15px; font-weight: 700; color: var(--gold); }}
  .header-gen {{ font-size: 11px; color: rgba(255,255,255,.4); margin-top: 3px; }}

  .main {{ padding: 32px 40px; max-width: 1600px; margin: 0 auto; }}

  .section-title {{ font-size: 12px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 1.5px; color: var(--orange); margin: 36px 0 16px;
    display: flex; align-items: center; gap: 10px; }}
  .section-title::after {{ content: ''; flex: 1; height: 1px; background: var(--border); }}

  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; }}
  .kpi {{ background: var(--card); border-radius: 12px; padding: 20px 22px;
    border: 1px solid var(--border); box-shadow: 0 2px 8px rgba(0,0,0,.05);
    transition: transform .2s; }}
  .kpi:hover {{ transform: translateY(-2px); }}
  .kpi-label {{ font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 1px; color: var(--muted); margin-bottom: 8px; }}
  .kpi-value {{ font-size: 26px; font-weight: 800; line-height: 1; }}
  .kpi-delta {{ margin-top: 8px; font-size: 12px; font-weight: 600; }}
  .kpi-delta.up {{ color: var(--green); }} .kpi-delta.dn {{ color: var(--red); }}
  .kpi-delta.warn {{ color: #c07000; }} .kpi-delta.neu {{ color: var(--muted); }}
  .kpi-bar {{ height: 4px; background: var(--border); border-radius: 2px; margin-top: 10px; }}
  .kpi-bar-fill {{ height: 100%; border-radius: 2px; }}

  .g2 {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; }}
  .g3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }}
  .card {{ background: var(--card); border-radius: 12px; padding: 24px;
    border: 1px solid var(--border); box-shadow: 0 2px 8px rgba(0,0,0,.05); }}
  .card.full {{ grid-column: 1 / -1; }}
  .card-title {{ font-size: 14px; font-weight: 700; margin-bottom: 4px; }}
  .card-sub {{ font-size: 11px; color: var(--muted); }}
  .chart-wrap {{ position: relative; margin-top: 16px; }}
  .chart-wrap.h260 {{ height: 260px; }} .chart-wrap.h320 {{ height: 320px; }}
  .chart-wrap.h200 {{ height: 200px; }} .chart-wrap.h400 {{ height: 400px; }}

  .tw {{ overflow-x: auto; margin-top: 12px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  thead tr {{ background: var(--dark); color: #fff; }}
  thead th {{ padding: 10px 12px; text-align: left; font-size: 10px; font-weight: 600;
    text-transform: uppercase; letter-spacing: .8px; white-space: nowrap; }}
  tbody tr {{ border-bottom: 1px solid var(--border); transition: background .15s; }}
  tbody tr:hover {{ background: rgba(232,93,4,.04); }}
  tbody td {{ padding: 9px 12px; vertical-align: middle; }}

  .pill {{ display: inline-block; padding: 2px 8px; border-radius: 20px;
    font-size: 11px; font-weight: 700; white-space: nowrap; }}
  .pill.green  {{ background: rgba(45,198,83,.12);  color: #1a9e3f; }}
  .pill.red    {{ background: rgba(230,57,70,.12);  color: #c0392b; }}
  .pill.amber  {{ background: rgba(244,162,97,.15); color: #c07000; }}
  .pill.neutral{{ background: rgba(107,114,128,.1); color: var(--muted); }}

  .warn-banner {{ background: var(--warn-bg); border: 1px solid var(--warn-border);
    border-radius: 10px; padding: 16px 20px; margin-bottom: 20px;
    font-size: 13px; color: #7a4900; line-height: 1.6; }}
  .warn-banner strong {{ color: #c07000; }}

  .data-gap-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 12px; }}
  .gap-card {{ background: #fff7f0; border: 1px solid #f4a261; border-radius: 10px;
    padding: 16px; }}
  .gap-card h4 {{ font-size: 13px; font-weight: 700; color: var(--orange); margin-bottom: 6px; }}
  .gap-card p {{ font-size: 12px; color: #6b4b2a; line-height: 1.5; }}
  .gap-badge {{ display: inline-block; background: var(--orange); color: #fff;
    border-radius: 4px; font-size: 10px; font-weight: 700; padding: 1px 6px; margin-right: 4px; }}

  footer {{ text-align: center; padding: 28px 40px; font-size: 11px; color: var(--muted);
    border-top: 1px solid var(--border); margin-top: 40px; line-height: 1.6; }}

  @media (max-width: 900px) {{
    .main {{ padding: 20px; }}
    .header {{ padding: 16px 20px; }}
    .g2, .g3 {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="logo">OP</div>
    <div>
      <div class="header-title">Olympic Paints</div>
      <div class="header-sub">Weekly KPI Dashboard — Sales Operations</div>
    </div>
  </div>
  <div class="header-right">
    <div class="header-week">Week ending {REPORT_WEEK}</div>
    <div class="header-gen">Generated {generated}</div>
  </div>
</div>

<div class="main">

  <!-- ALERT BANNER -->
  <div class="warn-banner" style="margin-top:24px">
    <strong>⚠ Partial KPI Scoring</strong><br>
    KPIs A, C and F are scored from live data (QuickSight + Zoho exports).
    KPI B (Discount Mgmt) is partial — RB% available for AP and NP only.
    KPIs D and E require CRM upsell and new-store records which are
    <strong>not yet in the automated feed</strong>.
    See the <a href="#data-gaps" style="color:var(--orange)">Data Gaps section</a> for details.
  </div>

  <!-- EXECUTIVE SUMMARY -->
  <div class="section-title">Executive Summary — Week Ending {REPORT_WEEK}</div>
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">MTD Sales</div>
      <div class="kpi-value">{fmt_r(MTD_SALES)}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(MTD_SALES/MTD_TARGET*100,100):.0f}%;background:{'var(--green)' if MTD_PCT_TARGET >= 0 else 'var(--amber)'}"></div></div>
      <div class="kpi-delta {'up' if MTD_PCT_TARGET >= 0 else 'warn'}">{'▲' if MTD_PCT_TARGET >= 0 else '▼'} {abs(MTD_PCT_TARGET):.1f}% {'above' if MTD_PCT_TARGET >= 0 else 'below'} monthly target</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Monthly Target</div>
      <div class="kpi-value">{fmt_r(MTD_TARGET)}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:100%;background:var(--muted)"></div></div>
      <div class="kpi-delta neu">Gap: {fmt_r(MTD_TARGET - MTD_SALES)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Outstanding Debtors</div>
      <div class="kpi-value" style="font-size:20px">{fmt_r(DEBTORS_TOTAL)}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:100%;background:var(--red)"></div></div>
      <div class="kpi-delta dn">▲ {fmt_r(DEBTORS_90D)} (&gt;90 days)</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">% Overdue &gt;60 Days</div>
      <div class="kpi-value">{pct_plain(OVERDUE_60D_PCT)}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(OVERDUE_60D_PCT*4,100):.0f}%;background:var(--red)"></div></div>
      <div class="kpi-delta dn">▲ Above 10% threshold — collections attention needed</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Rock Bottom % (avg)</div>
      <div class="kpi-value">{pct_plain(ABOVE_RB_AVG)}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{ABOVE_RB_AVG/ABOVE_RB_TARGET*100:.0f}%;background:var(--amber)"></div></div>
      <div class="kpi-delta warn">▼ Target: {pct_plain(ABOVE_RB_TARGET)} — {pct_plain(ABOVE_RB_TARGET - ABOVE_RB_AVG)} gap</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Team Sales Attainment</div>
      <div class="kpi-value">{total_pct:+.1f}%</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(total_sales/total_target*100,100):.0f}%;background:{'var(--green)' if total_pct >= 0 else 'var(--amber)'}"></div></div>
      <div class="kpi-delta {'up' if total_pct >= 0 else 'warn'}">{'▲ Ahead' if total_pct >= 0 else '▼ Behind'} — {fmt_r(total_sales)} vs {fmt_r(total_target)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">E-Commerce Orders</div>
      <div class="kpi-value">{ECOM_TOTAL_ORDERS}</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:60%;background:var(--teal)"></div></div>
      <div class="kpi-delta neu">{ECOM_TOTAL_QTY} units — see fulfilment section</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">YOY Growth (Apr)</div>
      <div class="kpi-value">+16.1%</div>
      <div class="kpi-bar"><div class="kpi-bar-fill" style="width:60%;background:var(--green)"></div></div>
      <div class="kpi-delta up">▲ Apr 2026 vs Apr 2025</div>
    </div>
  </div>

  <!-- KPI COMPLIANCE SCORECARD -->
  <div class="section-title">KPI Compliance Scorecard — Per Rep</div>
  <div class="warn-banner">
    <strong>Scoring Status:</strong>
    &nbsp;<span class="pill green" style="font-size:11px">A — Sales Growth (30%)</span>
    <span class="pill green" style="font-size:11px">C — Customer Dev/CRM (20%)</span>
    <span class="pill green" style="font-size:11px">F — Merchandising (10%)</span>
    Scored from QuickSight + Zoho exports.
    &nbsp;<span class="pill amber" style="font-size:11px">B — Discount Mgmt (20%, partial)</span>
    RB % available for some reps only.
    &nbsp;<span class="pill neutral" style="font-size:11px">D — Product Dev (10%)</span>
    <span class="pill neutral" style="font-size:11px">E — New Customers (10%)</span>
    Awaiting data exports.
    <strong>Total scoreable weight this week: 60–80% of 100%.</strong>
  </div>
  <div class="card full">
    <div class="card-title">Rep KPI Agreement Scorecard</div>
    <div class="card-sub">Sales Reps Incentive &amp; KPI Framework — 6 weighted categories</div>
    <div class="tw">
      <table>
        <thead>
          <tr>
            <th>Rep</th>
            <th>MTD Sales</th>
            <th>Monthly Target</th>
            <th>vs Target</th>
            <th>Rock Bottom %</th>
            <th>A — Sales Growth (30%)</th>
            <th>B — Discount Mgmt (20%)</th>
            <th>C — Customer Dev (20%)</th>
            <th>D — Product Dev (10%)</th>
            <th>E — New Customers (10%)</th>
            <th>F — Merchandising (10%)</th>
            <th>Approved Orders</th>
          </tr>
        </thead>
        <tbody>{rep_rows}</tbody>
      </table>
    </div>
    <p style="margin-top:12px;font-size:11px;color:var(--muted)">
      KPI A scoring uses YOY growth vs prior year baseline. Per the KPI Agreement: growth must exceed 10% above baseline
      to qualify for commission. Reps below 10% YOY growth do not qualify for KPI A incentive this period.
    </p>
  </div>

  <!-- KPI ACHIEVEMENT SUMMARY (per-rep roll-up) -->
  <div class="section-title">KPI Achievement Summary — Merchandising (KPI F, 10%)</div>
  {merch_summary_card}

  <!-- REP PERFORMANCE CHARTS -->
  <div class="section-title">Rep Sales Performance — April 2026</div>
  <div class="g2">
    <div class="card">
      <div class="card-title">MTD Sales vs Monthly Target</div>
      <div class="card-sub">R values for week ending {REPORT_WEEK}</div>
      <div class="chart-wrap h320"><canvas id="cRepTarget"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Q2 Target Tracking — YTD vs Q2 Target</div>
      <div class="card-sub">Quarter 2 cumulative — April 2026</div>
      <div class="chart-wrap h320"><canvas id="cQ2Track"></canvas></div>
    </div>
  </div>

  <!-- YOY -->
  <div class="section-title">Year-on-Year Sales Comparison</div>
  <div class="g2">
    <div class="card">
      <div class="card-title">Monthly Sales: 2025 vs 2026</div>
      <div class="card-sub">Jan–Apr comparison</div>
      <div class="chart-wrap h260"><canvas id="cYOY"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">YOY Growth % by Month</div>
      <div class="card-sub">Positive = ahead of prior year</div>
      <div class="chart-wrap h260"><canvas id="cYOYpct"></canvas></div>
    </div>
  </div>

  <!-- ROCK BOTTOM ANALYSIS -->
  <div class="section-title">Margin Health — Rock Bottom Analysis</div>
  <div class="warn-banner">
    <strong>⚠ Margin Risk:</strong> Company average rock bottom margin is <strong>{pct_plain(ABOVE_RB_AVG)}</strong>
    against a target of <strong>{pct_plain(ABOVE_RB_TARGET)}</strong>.
    {len([x for x in RB_BY_PRODUCT if x["rb_pct"] < 0])} product groups are trading <strong>below rock bottom</strong> (negative margin).
    Immediate pricing review required for highlighted accounts and products.
  </div>
  <div class="g2">
    <div class="card">
      <div class="card-title">Rock Bottom % — Product Groups (Selected)</div>
      <div class="card-sub">Red = below rock bottom, Amber = below 8% target, Green = healthy</div>
      <div class="chart-wrap h400"><canvas id="cRBProduct"></canvas></div>
    </div>
    <div class="g2" style="display:flex;flex-direction:column;gap:20px">
      <div class="card">
        <div class="card-title">Customers at Risk — Rock Bottom %</div>
        <div class="card-sub">Accounts trading below or near rock bottom price</div>
        <div class="tw">
          <table>
            <thead><tr><th>Customer</th><th>Rock Bottom %</th></tr></thead>
            <tbody>{cust_rows}</tbody>
          </table>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Product Groups Below Rock Bottom</div>
        <div class="card-sub">Negative margin — requires immediate price correction</div>
        <div class="tw">
          <table>
            <thead><tr><th>Product Group</th><th>Rock Bottom %</th></tr></thead>
            <tbody>{rb_risk_rows}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- PRODUCT MIX -->
  <div class="section-title">Product Mix — Revenue Distribution</div>
  <div class="g2">
    <div class="card">
      <div class="card-title">Revenue by Product Category (April 2026)</div>
      <div class="card-sub">Based on average group sales values from weekly report</div>
      <div class="chart-wrap h260"><canvas id="cProductMix"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Debtors &amp; Collections Health</div>
      <div class="card-sub">Outstanding debt position — week ending {REPORT_WEEK}</div>
      <div style="margin-top:20px">
        <div style="display:flex;gap:16px;flex-wrap:wrap">
          <div style="flex:1;min-width:140px;background:#fff0f0;border-radius:10px;padding:16px;border:1px solid #fcc">
            <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase">Total Debtors</div>
            <div style="font-size:22px;font-weight:800;color:var(--dark);margin-top:4px">{fmt_r(DEBTORS_TOTAL)}</div>
          </div>
          <div style="flex:1;min-width:140px;background:#fff0f0;border-radius:10px;padding:16px;border:1px solid #fcc">
            <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase">Overdue &gt;90 Days</div>
            <div style="font-size:22px;font-weight:800;color:var(--red);margin-top:4px">{fmt_r(DEBTORS_90D)}</div>
          </div>
          <div style="flex:1;min-width:140px;background:#fff7e6;border-radius:10px;padding:16px;border:1px solid #f4a261">
            <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase">% Overdue &gt;60d</div>
            <div style="font-size:22px;font-weight:800;color:#c07000;margin-top:4px">{pct_plain(OVERDUE_60D_PCT)}</div>
            <div style="font-size:11px;color:#c07000;margin-top:4px">Target: &lt;10%</div>
          </div>
        </div>
        <p style="margin-top:16px;font-size:12px;color:var(--muted);line-height:1.6">
          Collections are a qualifying condition for KPI A commission.
          Only invoices settled within 30–60 days qualify under the KPI Agreement.
          A collections audit is required to determine which rep sales are eligible.
        </p>
      </div>
    </div>
  </div>

  <!-- E-COMMERCE -->
  <div class="section-title">E-Commerce Fulfilment — 21 April 2026</div>
  <div class="g3">
    <div class="card">
      <div class="kpi-label">Total Open Orders</div>
      <div class="kpi-value" style="font-size:36px;color:var(--dark)">{ECOM_TOTAL_ORDERS}</div>
      <div class="kpi-delta neu">{ECOM_TOTAL_QTY} total units</div>
    </div>
    <div class="card">
      <div class="kpi-label">Orders by Age in System</div>
      <div style="margin-top:12px">
        {''.join(f'<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border);font-size:13px"><span>{k}</span><span style="font-weight:700">{v} orders</span></div>' for k, v in ECOM_AGING.items())}
      </div>
    </div>
    <div class="card">
      <div class="kpi-label">Fulfilment Risk Notes</div>
      <p style="font-size:12px;color:var(--muted);margin-top:10px;line-height:1.6">
        22 orders are 15+ days in system. Key statuses: <strong>pending</strong> (unconfirmed),
        <strong>manufacturing</strong> (in production), <strong>on-hold</strong> (blocked).
        Multiple orders to <em>Deniz</em>, <em>Latticia</em>, and <em>Lynette</em> accounts are ageing.
        Dispatch target dates have been exceeded for most open orders.
      </p>
    </div>
  </div>

  <!-- DATA GAPS -->
  <div class="section-title" id="data-gaps">Data Gaps — What Is Needed to Complete KPI Scoring</div>
  <div class="data-gap-grid">
    <div class="gap-card">
      <h4><span class="gap-badge">B</span> Discount Management — partial data (20%)</h4>
      <p>Rock bottom % available for AP (8.85%) and NP (8.82%) only. Target ≥15% for full 20 pts.
      Full per-rep markup vs RB required for AC, BV, BM.
      Currently all reps scoring below the 11% minimum band.</p>
    </div>
    <div class="gap-card">
      <h4><span class="gap-badge">C</span> Customer Dev &amp; CRM — fully scored (20%)</h4>
      <p>Sourced from <em>OP_Lead_Tracking_new.csv</em> (Zoho CRM leads export).
      Scored monthly: ≥5 leads = Hit (green), &lt;5 = Miss (red).
      Showing {leads_period} data.</p>
    </div>
    <div class="gap-card">
      <h4><span class="gap-badge">D</span> Product Development Upsells (10%)</h4>
      <p>Need: CRM records of upsell activities — product, customer, date, evidence photo.
      Focus products defined quarterly (e.g. Natural Elegance, 7-in-1 PVA, Fibre Restore).
      Minimum 4–5 per rep per month.</p>
    </div>
    <div class="gap-card">
      <h4><span class="gap-badge">E</span> New Customer Onboarding (10%)</h4>
      <p>Need: List of new verified trading customers onboarded per rep per quarter —
      credit application completed, first invoice issued, CRM logged.
      Quarterly target: 6 = full credit, 3 = half credit.</p>
    </div>
    <div class="gap-card">
      <h4><span class="gap-badge">F</span> Merchandising — fully scored (10%)</h4>
      <p>Sourced from <em>Meetings_Report_Merchandising</em> (direct Zoho REST API).
      Graded vs 15 visits / rep / month (45 / quarter).
      Training is no longer part of KPI F scoring — re-introduce only once a training register feed exists.</p>
    </div>
    <div class="gap-card">
      <h4><span class="gap-badge">A</span> Collections Audit (30%)</h4>
      <p>Need: Per-rep collections rate — which invoices have been settled within 30–60 days.
      Only settled invoices qualify for KPI A commission.
      Current data shows sales only, not settlement status.</p>
    </div>
  </div>

</div>

<footer>
  <strong>Olympic Paints</strong> · Weekly KPI Dashboard · Week ending {REPORT_WEEK}<br>
  Data sourced from Amazon QuickSight Weekly Sales Reports (Weekly Progress folder).<br>
  KPI scoring requires additional CRM and activity data not yet available in automated export.<br>
  Based on: Sales Reps Incentive &amp; KPI Framework — 6 categories (Sales Growth 30% | Discount Mgmt 20% | Customer Dev 20% | Product Dev 10% | New Customers 10% | Merchandising 10%)<br>
  Generated {generated}
</footer>

<script>
const O='#E85D04', R='#E63946', T='#2EC4B6', G='#2DC653', B='#457B9D', GD='#F4A261', DK='#1A1A2E';
Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#6B7280';
const grid = {{color:'rgba(0,0,0,0.06)'}};

// Rep vs Target
new Chart(document.getElementById('cRepTarget'), {{
  type: 'bar',
  data: {{
    labels: {rep_names_js},
    datasets: [
      {{label:'MTD Sales',   data:{rep_sales_js},  backgroundColor: O+'CC', borderRadius:6}},
      {{label:'Monthly Target', data:{rep_target_js}, backgroundColor: 'rgba(107,114,128,0.25)', borderRadius:6}}
    ]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom'}},
      tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}}}}}}
  }}
}});

// Q2 Tracking
new Chart(document.getElementById('cQ2Track'), {{
  type: 'bar',
  data: {{
    labels: {q2_names_js},
    datasets: [
      {{label:'YTD Sales',  data:{q2_sales_js},  backgroundColor: O+'CC', borderRadius:6}},
      {{label:'Q2 Target',  data:{q2_target_js}, backgroundColor: 'rgba(107,114,128,0.25)', borderRadius:6}}
    ]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom'}},
      tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}}}}}}
  }}
}});

// YOY line
new Chart(document.getElementById('cYOY'), {{
  type: 'line',
  data: {{
    labels: {months_js},
    datasets: [
      {{label:'2025', data:{s25_js}, borderColor:'rgba(107,114,128,0.6)', fill:false, tension:0.3, pointRadius:5, borderDash:[5,3]}},
      {{label:'2026', data:{s26_js}, borderColor:O, fill:false, tension:0.3, pointRadius:6, borderWidth:2.5}}
    ]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom'}},
      tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}}}}}}
  }}
}});

// YOY %
new Chart(document.getElementById('cYOYpct'), {{
  type: 'bar',
  data: {{
    labels: {months_js},
    datasets: [{{
      label:'YOY Growth %', data:{yoy_pct_js},
      backgroundColor: {yoy_pct_js}.map(v=>v>=0?G+'CC':R+'CC'),
      borderRadius: 6
    }}]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},
      tooltip:{{callbacks:{{label:c=>c.parsed.y.toFixed(1)+'%'}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>v+'%'}}}},x:{{grid:{{display:false}}}}}}
  }}
}});

// RB by Product
new Chart(document.getElementById('cRBProduct'), {{
  type: 'bar',
  data: {{
    labels: {rb_labels_js},
    datasets: [{{
      label:'Rock Bottom %', data:{rb_vals_js},
      backgroundColor:{rb_colors_js},
      borderRadius:4
    }}]
  }},
  options: {{
    indexAxis:'y', responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},
      tooltip:{{callbacks:{{label:c=>c.parsed.x.toFixed(2)+'%'}}}},
      annotation:{{annotations:{{line1:{{type:'line',xMin:8,xMax:8,borderColor:'rgba(45,198,83,0.8)',borderWidth:2,borderDash:[6,4],label:{{display:true,content:'8% Target',position:'end'}}}}}}}}}},
    scales:{{x:{{grid,ticks:{{callback:v=>v+'%'}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}}}
  }}
}});

// Product Mix doughnut
new Chart(document.getElementById('cProductMix'), {{
  type: 'doughnut',
  data: {{
    labels: {pm_labels_js},
    datasets: [{{
      data: {pm_vals_js},
      backgroundColor: [O,T,G,B,GD,R,'#7B2D8B'],
      borderWidth: 2, borderColor:'#fff'
    }}]
  }},
  options: {{
    responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{position:'bottom',labels:{{padding:12,boxWidth:12}}}},
      tooltip:{{callbacks:{{label:c=>c.label+': '+c.parsed+'%'}}}}}}
  }}
}});

</script>
</body>
</html>"""
    return html


# ── KPI STATUS JSON ───────────────────────────────────────────────────────────

def write_kpi_status(generated: str):
    """Write kpi_status.json to the workspace dashboard directory so the
    workspace dashboard KPI tab always reflects the latest weekly data."""
    status = {
        "report_week": REPORT_WEEK,
        "report_date": REPORT_DATE,
        "generated_at": generated,
        "kpi_dashboard_url": "https://flomaticauto.github.io/olympic-paints-kpi/",
        "update_history": [
            {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "week": REPORT_WEEK,
                "by": "build_kpi_dashboard.py",
                "source_files": [f.name for f in sorted(WEEKLY_DIR.glob("*.pdf"))] if WEEKLY_DIR.exists() else [],
                "kpis_scored": ["A"],
                "kpis_missing": ["B", "C", "D", "E"],
            }
        ],
        "headline": {
            "mtd_sales":        MTD_SALES,
            "mtd_target":       MTD_TARGET,
            "mtd_pct":          MTD_PCT_TARGET,
            "debtors_total":    DEBTORS_TOTAL,
            "debtors_90d":      DEBTORS_90D,
            "overdue_60d_pct":  OVERDUE_60D_PCT,
            "above_rb_avg":     ABOVE_RB_AVG,
            "above_rb_target":  ABOVE_RB_TARGET,
            "yoy_apr_pct":      YOY[-1]["yoy_pct"] if YOY else None,
            "ecom_open_orders": ECOM_TOTAL_ORDERS,
        },
        "reps": [
            {
                "code": r["code"], "name": r["name"],
                "sales": r["sales"], "target": r["target"],
                "pct_target": r["pct"], "yoy": r["yoy"], "rb_pct": r["rb_pct"],
            }
            for r in REPS
        ],
        "kpi_categories": [
            {
                "id":             c["id"],
                "name":           c["name"],
                "weight":         c["weight"],
                "status":         "scored" if c["id"] == "A" else "missing",
                "data_available": c["id"] == "A",
                "description":    c["description"],
                "note": (
                    "Sales data available from QuickSight. Collections settlement audit still needed."
                    if c["id"] == "A" else
                    "Needs: Zoho CRM export of new and reactivated customers per rep."
                    if c["id"] == "B" else
                    "Needs: CRM upsell activity log with product, customer, date, evidence."
                    if c["id"] == "C" else
                    "Needs: Credit application confirmed + first invoice + CRM log per new account."
                    if c["id"] == "D" else
                    "Needs: Attendance lists, training photos, POS compliance evidence per rep."
                ),
            }
            for c in KPI_CATEGORIES
        ],
        "data_gaps": [
            {"id": "CRM_export",   "label": "Zoho CRM Export",         "affects_kpis": ["B","C","D"], "priority": "high",   "description": "Monthly Zoho CRM export per rep showing new/reactivated customers, upsell activities, and new account onboarding."},
            {"id": "collections",  "label": "Invoice Settlement Audit", "affects_kpis": ["A"],         "priority": "high",   "description": "Per-rep list of which invoices have been settled within 30-60 days. Only settled invoices qualify for KPI A commission."},
            {"id": "training",     "label": "Training Records",         "affects_kpis": ["E"],         "priority": "medium", "description": "Attendance lists, photos, and reports for key account trainings. Minimum 1 per key account per month."},
            {"id": "store_visits", "label": "Store Visit Log",          "affects_kpis": ["B"],         "priority": "medium", "description": "Store visit count per rep from Zoho or manual log. QuickSight shows No Data."},
            {"id": "leads",        "label": "Leads Created",            "affects_kpis": ["B","D"],     "priority": "medium", "description": "Leads created per rep from CRM. QuickSight shows No Data."},
            {"id": "rb_per_rep",   "label": "Rock Bottom % Per Rep",    "affects_kpis": [],            "priority": "low",    "description": "Rock bottom % by rep (only AP and NP available). Full per-rep margin data needed."},
        ],
        "scoreable_weight_pct": 50,
        "risk_items": [
            {"item": f"BV {next(r['pct'] for r in REPS if r['code']=='BV'):.1f}% vs target MTD", "severity": "high"},
            {"item": f"NP {next(r['pct'] for r in REPS if r['code']=='NP'):.1f}% vs target MTD", "severity": "high"},
            {"item": f"{len([x for x in RB_BY_PRODUCT if x['rb_pct'] < 0])} product groups trading below rock bottom", "severity": "high"},
            {"item": "Cassim Tayob customer at -37.14% rock bottom", "severity": "high"},
            {"item": f"{OVERDUE_60D_PCT}% of debtors overdue >60 days (target <10%)", "severity": "high"},
            {"item": f"MTD sales {fmt_r(MTD_TARGET - MTD_SALES)} below monthly target", "severity": "medium"},
            {"item": f"Rock bottom avg {ABOVE_RB_AVG}% vs {ABOVE_RB_TARGET}% target", "severity": "medium"},
            {"item": "22 e-commerce orders aged 15+ days", "severity": "medium"},
        ],
    }

    dest = WORKSPACE_DASH / "kpi_status.json"
    if WORKSPACE_DASH.exists():
        dest.write_text(json.dumps(status, indent=2), encoding="utf-8")
        print(f"  OK kpi_status.json -> {dest}")
    else:
        print(f"  [WARN] Workspace dashboard dir not found: {WORKSPACE_DASH}")

    # Also write alongside the KPI dashboard itself
    (BASE_DIR / "kpi_status.json").write_text(json.dumps(status, indent=2), encoding="utf-8")


# ── GIT PUSH ──────────────────────────────────────────────────────────────────

def _flomatic_token():
    """Retrieve the FlomaticAuto PAT from the GitHub CLI keyring."""
    try:
        r = subprocess.run(
            ["gh", "auth", "token", "--user", "FlomaticAuto"],
            capture_output=True, text=True, shell=True,
        )
        token = r.stdout.strip()
        if r.returncode == 0 and token.startswith("gho_"):
            return token
    except Exception as e:
        print(f"  [WARN] could not get FlomaticAuto token: {e}")
    return None


def git_push(path: Path):
    cwd = str(path)
    msg = f"Weekly KPI Dashboard update — {REPORT_WEEK} — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    token = _flomatic_token()

    def run(cmd):
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
        if r.returncode != 0:
            err = r.stderr.strip()
            if token:
                err = err.replace(token, "***")
            print(f"  [WARN] {' '.join(cmd)}: {err}")
        return r.returncode == 0

    run(["git", "config", "user.email", "auto@olympic-paints.local"])
    run(["git", "config", "user.name",  "Olympic KPI Bot"])
    run(["git", "add", "index.html"])
    run(["git", "add", "KPI Dashboard.html"])
    run(["git", "add", "build_kpi_dashboard.py"])
    run(["git", "commit", "-m", msg])

    push_cmd = ["git", "push", "origin", "master"]
    if token:
        # Push to the authed URL once without persisting it in the remote config
        authed = f"https://FlomaticAuto:{token}@github.com/FlomaticAuto/olympic-paints-kpi.git"
        push_cmd = ["git", "push", authed, "master"]

    ok = run(push_cmd)
    if not ok and not token:
        run(["git", "push", "-u", "origin", "master"])
    if ok:
        print(f"  ✓ Pushed to GitHub")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Building Olympic Paints KPI Dashboard...")
    print(f"  Report week: {REPORT_WEEK}")

    html = build_html()
    DASHBOARD.write_text(html, encoding="utf-8")
    INDEX.write_text(html, encoding="utf-8")
    print(f"  OK Written: {INDEX}")

    generated_ts = datetime.now().strftime("%Y-%m-%d %H:%M UTC")
    write_kpi_status(generated_ts)

    git_push(BASE_DIR)
    print("  Done.")


if __name__ == "__main__":
    main()
