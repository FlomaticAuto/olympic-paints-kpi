"""
build_kpi_dashboard.py
Reads all Sales_Summary_*.pdf snapshots, extracts KPI data,
regenerates KPI Dashboard.html, then commits and pushes to GitHub.

Run manually or via Windows Task Scheduler whenever new PDFs arrive.
"""

import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# ── pip-install pdfplumber on first run ──────────────────────────────────────
try:
    import pdfplumber
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pdfplumber", "-q"])
    import pdfplumber

BASE_DIR   = Path(__file__).parent
ARCHIVE    = BASE_DIR / "Daily Reports" / "Daily Reports Archive"
LIVE       = BASE_DIR / "Daily Reports"
DASHBOARD  = BASE_DIR / "KPI Dashboard.html"

# ── helpers ──────────────────────────────────────────────────────────────────

def find_snapshots():
    """Return sorted list of (date, path) for every Sales_Summary PDF."""
    pdfs = []
    for folder in [ARCHIVE, LIVE]:
        if folder.exists():
            for f in folder.glob("Sales_Summary_*.pdf"):
                m = re.search(r"(\d{4}-\d{2}-\d{2})", f.name)
                if m:
                    dt = datetime.strptime(m.group(1), "%Y-%m-%d")
                    pdfs.append((dt, f))
    pdfs.sort(key=lambda x: x[0])
    return pdfs


def extract_kpi(path: Path) -> dict:
    """Pull key numbers from a Sales_Summary PDF via text extraction."""
    data = {"path": str(path), "date": None,
            "mtd_sales": None, "debtors": None,
            "overdue_amt": None, "overdue_pct": None,
            "above_rb": None,
            "aboo_pct": None, "amit_pct": None,
            "bhadresh_pct": None, "nikhil_pct": None,
            "credits_total": None}

    def clean(s):
        return re.sub(r"[R\s,]", "", s or "")

    try:
        with pdfplumber.open(path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        print(f"  [WARN] Could not read {path.name}: {e}")
        return data

    # Date from filename
    m = re.search(r"(\d{4}-\d{2}-\d{2})", path.name)
    if m:
        data["date"] = m.group(1)

    # MTD sales — "ivnett\n<number>" or "ivnett (Sum)\n" pattern
    for pattern in [
        r"ivnett\s*\n\s*([\d,]+\.?\d*)",
        r"Total Monthly Sales\s*\n.*?\n\s*([\d,]+\.?\d*)",
    ]:
        m = re.search(pattern, text)
        if m:
            data["mtd_sales"] = float(clean(m.group(1)))
            break

    # Outstanding debtors
    m = re.search(r"Outstanding Debtors\s*\n\s*([\d,]+\.?\d*)", text)
    if m:
        data["debtors"] = float(clean(m.group(1)))

    # Overdue >90 days amount
    m = re.search(r"Outstanding Over 90 Days\s*\n\s*R?([\d,]+\.?\d*)", text)
    if m:
        data["overdue_amt"] = float(clean(m.group(1)))

    # Overdue %
    m = re.search(r"% Overdue > 90 Days\s*\n\s*([\d.]+)%", text)
    if m:
        data["overdue_pct"] = float(m.group(1))

    # Above rock bottom
    m = re.search(r"Above/Below Rock Bottom\s*\n\s*([\d.]+)%", text)
    if m:
        data["above_rb"] = float(m.group(1))

    # Rep attainment %
    for rep, key in [("Aboo", "aboo_pct"), ("Amit", "amit_pct"),
                     ("Bhadresh", "bhadresh_pct"), ("Nikhil", "nikhil_pct")]:
        # Look for pattern like "351,523\n33.19%"
        pat = rf"{rep}[^%\n]{{0,60}}?\n\s*([\d.]+)%"
        m = re.search(pat, text)
        if m:
            data[key] = float(m.group(1))

    # Credits total (absolute value of sum at bottom of credits table)
    credits = re.findall(r"-R([\d,]+)", text)
    if credits:
        total = sum(float(c.replace(",", "")) for c in credits)
        data["credits_total"] = total

    return data


def fmt_r(val, decimals=0):
    """Format as South African Rand string."""
    if val is None:
        return "—"
    if val >= 1_000_000:
        return f"R{val/1_000_000:.2f}M"
    if val >= 1_000:
        return f"R{val/1_000:.1f}K"
    return f"R{val:.{decimals}f}"


def pct(val):
    return f"{val:.1f}%" if val is not None else "—"


# ── HTML GENERATION ──────────────────────────────────────────────────────────

def build_html(snapshots: list[dict]) -> str:
    generated = datetime.now().strftime("%d %B %Y %H:%M")
    n_snaps   = len(snapshots)
    latest    = snapshots[-1] if snapshots else {}
    earliest  = snapshots[0]  if snapshots else {}

    # Build JS arrays
    snap_labels = [s["date"][5:] if s.get("date") else "?" for s in snapshots]  # MM-DD
    # Pretty labels
    snap_labels_pretty = []
    for s in snapshots:
        if s.get("date"):
            dt = datetime.strptime(s["date"], "%Y-%m-%d")
            snap_labels_pretty.append(dt.strftime("%-d %b") if sys.platform != "win32"
                                      else dt.strftime("%d %b").lstrip("0"))
        else:
            snap_labels_pretty.append("?")

    def js_arr(vals):
        parts = []
        for v in vals:
            parts.append("null" if v is None else str(round(v, 2)))
        return "[" + ",".join(parts) + "]"

    mtd_arr      = js_arr([s.get("mtd_sales")    for s in snapshots])
    debtors_arr  = js_arr([s.get("debtors")       for s in snapshots])
    overdue_arr  = js_arr([s.get("overdue_pct")   for s in snapshots])
    rb_arr       = js_arr([s.get("above_rb")       for s in snapshots])
    aboo_arr     = js_arr([s.get("aboo_pct")       for s in snapshots])
    amit_arr     = js_arr([s.get("amit_pct")       for s in snapshots])
    bhadresh_arr = js_arr([s.get("bhadresh_pct")   for s in snapshots])
    nikhil_arr   = js_arr([s.get("nikhil_pct")     for s in snapshots])
    credits_arr  = js_arr([s.get("credits_total")  for s in snapshots])

    labels_js = "[" + ",".join(f'"{l}"' for l in snap_labels_pretty) + "]"

    # KPI card values from latest snapshot
    latest_mtd     = fmt_r(latest.get("mtd_sales"))
    latest_debtors = fmt_r(latest.get("debtors"))
    latest_overdue = pct(latest.get("overdue_pct"))
    latest_rb      = pct(latest.get("above_rb"))
    latest_date    = latest.get("date", "—")
    earliest_date  = earliest.get("date", "—")

    # Table rows
    table_rows = ""
    for s in snapshots:
        dt_str = s.get("date", "—")
        try:
            dt_nice = datetime.strptime(dt_str, "%Y-%m-%d").strftime("%d %b %Y")
        except Exception:
            dt_nice = dt_str

        mtd_v = fmt_r(s.get("mtd_sales"))
        deb_v = fmt_r(s.get("debtors"))
        ovd_v = pct(s.get("overdue_pct"))
        rb_v  = pct(s.get("above_rb"))
        a_v   = pct(s.get("aboo_pct"))
        am_v  = pct(s.get("amit_pct"))
        bh_v  = pct(s.get("bhadresh_pct"))
        nk_v  = pct(s.get("nikhil_pct"))

        # Colour classes
        def ovd_cls(v):
            if v is None: return "neutral"
            return "green" if v < 10 else ("amber" if v < 15 else "red")
        def rb_cls(v):
            if v is None: return "neutral"
            return "green" if v >= 8 else "amber"

        table_rows += f"""
          <tr>
            <td><strong>{dt_nice}</strong></td>
            <td>{mtd_v}</td>
            <td>{deb_v}</td>
            <td><span class="pill {ovd_cls(s.get('overdue_pct'))}">{ovd_v}</span></td>
            <td><span class="pill {rb_cls(s.get('above_rb'))}">{rb_v}</span></td>
            <td>{a_v}</td><td>{am_v}</td><td>{bh_v}</td><td>{nk_v}</td>
          </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Olympic Paints — KPI Progress Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --orange:#E85D04; --dark:#1A1A2E; --mid:#2D2D44; --light:#F5F5F0;
    --gold:#F4A261; --teal:#2EC4B6; --red:#E63946; --green:#2DC653;
    --card:#FFFFFF; --border:#E8E8E0; --muted:#6B7280;
  }}
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Segoe UI',system-ui,sans-serif;background:var(--light);color:var(--dark)}}
  .header{{background:linear-gradient(135deg,var(--dark),var(--mid));color:#fff;padding:22px 40px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 4px 20px rgba(0,0,0,.3);position:sticky;top:0;z-index:100}}
  .logo{{width:52px;height:52px;background:var(--orange);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:900;color:#fff;flex-shrink:0;box-shadow:0 0 0 3px rgba(232,93,4,.3)}}
  .hl{{display:flex;align-items:center;gap:18px}}
  .ht{{font-size:21px;font-weight:700}}
  .hs{{font-size:12px;color:rgba(255,255,255,.55);margin-top:2px}}
  .hr{{text-align:right}}
  .hd{{font-size:11px;color:rgba(255,255,255,.4)}}
  .hsnap{{font-size:13px;color:var(--gold);font-weight:600;margin-top:2px}}
  .main{{padding:32px 40px;max-width:1600px;margin:0 auto}}
  .sec{{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:var(--orange);margin:36px 0 16px;display:flex;align-items:center;gap:10px}}
  .sec::after{{content:'';flex:1;height:1px;background:var(--border)}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px}}
  .kpi{{background:var(--card);border-radius:12px;padding:20px 22px;border:1px solid var(--border);box-shadow:0 2px 8px rgba(0,0,0,.05);transition:transform .2s,box-shadow .2s}}
  .kpi:hover{{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.1)}}
  .kl{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:8px}}
  .kv{{font-size:24px;font-weight:800;line-height:1}}
  .kd{{margin-top:8px;font-size:12px;font-weight:600}}
  .kd.up{{color:var(--green)}}.kd.dn{{color:var(--red)}}.kd.neu{{color:var(--muted)}}
  .kb{{height:4px;background:var(--border);border-radius:2px;margin-top:10px}}
  .kbf{{height:100%;border-radius:2px;background:var(--orange)}}
  .g2{{display:grid;grid-template-columns:repeat(2,1fr);gap:20px}}
  .g3{{display:grid;grid-template-columns:repeat(3,1fr);gap:20px}}
  .cc{{background:var(--card);border-radius:12px;padding:24px;border:1px solid var(--border);box-shadow:0 2px 8px rgba(0,0,0,.05)}}
  .cc.full{{grid-column:1/-1}}.cc.s2{{grid-column:span 2}}
  .ct{{font-size:14px;font-weight:700;margin-bottom:4px}}
  .cs{{font-size:11px;color:var(--muted)}}
  .cw{{position:relative;height:260px;margin-top:16px}}
  .cw.tall{{height:320px}}.cw.sh{{height:200px}}
  .tw{{overflow-x:auto;margin-top:4px}}
  table{{width:100%;border-collapse:collapse;font-size:13px}}
  thead tr{{background:var(--dark);color:#fff}}
  thead th{{padding:10px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;white-space:nowrap}}
  tbody tr{{border-bottom:1px solid var(--border);transition:background .15s}}
  tbody tr:hover{{background:rgba(232,93,4,.04)}}
  tbody td{{padding:9px 14px}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}}
  .pill.green{{background:rgba(45,198,83,.12);color:#1a9e3f}}
  .pill.red{{background:rgba(230,57,70,.12);color:#c0392b}}
  .pill.amber{{background:rgba(244,162,97,.15);color:#c07000}}
  .pill.neutral{{background:rgba(107,114,128,.1);color:var(--muted)}}
  footer{{text-align:center;padding:28px 40px;font-size:11px;color:var(--muted);border-top:1px solid var(--border);margin-top:40px}}
  footer strong{{color:var(--orange)}}
  @media(max-width:900px){{.main{{padding:20px}}.header{{padding:16px 20px}}.g2,.g3{{grid-template-columns:1fr}}.cc.s2{{grid-column:1}}}}
</style>
</head>
<body>
<div class="header">
  <div class="hl">
    <div class="logo">OP</div>
    <div>
      <div class="ht">Olympic Paints</div>
      <div class="hs">KPI Progress &amp; Trend Dashboard — Sales Operations</div>
    </div>
  </div>
  <div class="hr">
    <div class="hd">Generated {generated}</div>
    <div class="hsnap">{n_snaps} Snapshots &middot; {earliest_date} – {latest_date}</div>
  </div>
</div>

<div class="main">

  <div class="sec">Latest Snapshot — {latest_date}</div>
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kl">MTD Sales</div>
      <div class="kv">{latest_mtd}</div>
      <div class="kb"><div class="kbf" style="width:70%"></div></div>
    </div>
    <div class="kpi">
      <div class="kl">Outstanding Debtors</div>
      <div class="kv" style="font-size:20px">{latest_debtors}</div>
      <div class="kb"><div class="kbf" style="width:85%"></div></div>
    </div>
    <div class="kpi">
      <div class="kl">% Overdue &gt;90 Days</div>
      <div class="kv">{latest_overdue}</div>
      <div class="kd {'up' if latest.get('overdue_pct') and latest['overdue_pct'] < 10 else 'dn'}">{'▼ below 10% threshold' if latest.get('overdue_pct') and latest['overdue_pct'] < 10 else '▲ above 10% threshold'}</div>
    </div>
    <div class="kpi">
      <div class="kl">Above Rock Bottom %</div>
      <div class="kv">{latest_rb}</div>
      <div class="kd {'up' if latest.get('above_rb') and latest['above_rb'] >= 8 else 'dn'}">{'▲ healthy margin' if latest.get('above_rb') and latest['above_rb'] >= 8 else '▼ margin under pressure'}</div>
    </div>
    <div class="kpi">
      <div class="kl">Total Snapshots</div>
      <div class="kv">{n_snaps}</div>
      <div class="kd neu">Reports tracked</div>
    </div>
    <div class="kpi">
      <div class="kl">Data Range</div>
      <div class="kv" style="font-size:16px">{earliest_date[:7]}</div>
      <div class="kd neu">→ {latest_date[:7]}</div>
    </div>
  </div>

  <div class="sec">Monthly Sales — Cumulative MTD at Each Snapshot</div>
  <div class="g2">
    <div class="cc s2">
      <div class="ct">MTD Sales (R) — All Snapshots</div>
      <div class="cs">Each bar = total sales recorded at that snapshot date</div>
      <div class="cw tall"><canvas id="cMTD"></canvas></div>
    </div>
  </div>

  <div class="sec">Debtors Health</div>
  <div class="g2">
    <div class="cc">
      <div class="ct">Outstanding Debtors (R)</div>
      <div class="cs">Total book across all snapshots</div>
      <div class="cw"><canvas id="cDebtors"></canvas></div>
    </div>
    <div class="cc">
      <div class="ct">% Overdue &gt;90 Days</div>
      <div class="cs">Below 10% is healthy — red dashed line = danger threshold</div>
      <div class="cw"><canvas id="cOverdue"></canvas></div>
    </div>
  </div>

  <div class="sec">Rep Target Attainment — MTD % of Monthly Target</div>
  <div class="g2">
    <div class="cc s2">
      <div class="ct">Rep Attainment % — All Snapshots</div>
      <div class="cs">Cumulative MTD attainment per rep at each snapshot date</div>
      <div class="cw"><canvas id="cReps"></canvas></div>
    </div>
  </div>

  <div class="sec">Margin Health — Above Rock Bottom %</div>
  <div class="g2">
    <div class="cc s2">
      <div class="ct">Above Rock Bottom % — All Snapshots</div>
      <div class="cs">Measures how far above minimum margin the business trades. Target: ≥ 8%</div>
      <div class="cw"><canvas id="cRB"></canvas></div>
    </div>
  </div>

  <div class="sec">Credits &amp; Returns</div>
  <div class="g2">
    <div class="cc s2">
      <div class="ct">Total Credits Issued (R) — by Snapshot</div>
      <div class="cs">Sum of all credit notes in each report period</div>
      <div class="cw sh"><canvas id="cCredits"></canvas></div>
    </div>
  </div>

  <div class="sec">KPI Movement Table — Snapshot by Snapshot</div>
  <div class="cc full">
    <div class="tw">
      <table>
        <thead>
          <tr>
            <th>Snapshot</th><th>MTD Sales</th><th>Debtors</th>
            <th>Overdue &gt;90d</th><th>Above RB%</th>
            <th>Aboo %T</th><th>Amit %T</th><th>Bhadresh %T</th><th>Nikhil %T</th>
          </tr>
        </thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
  </div>

</div>

<footer><strong>Olympic Paints</strong> · KPI Dashboard · Auto-generated from QuickSight Sales Summary PDFs · {generated}</footer>

<script>
const LABELS = {labels_js};
const G = Chart.defaults;
G.font.family = "'Segoe UI',system-ui,sans-serif";
G.font.size = 12;
G.color = '#6B7280';
const grid = {{color:'rgba(0,0,0,0.06)'}};

const O='#E85D04',R='#E63946',T='#2EC4B6',G2='#2DC653',B='#457B9D',P='#7B2D8B',GD='#F4A261',DK='#1A1A2E';

new Chart(document.getElementById('cMTD'),{{
  type:'bar',
  data:{{labels:LABELS,datasets:[{{label:'MTD Sales',data:{mtd_arr},
    backgroundColor:LABELS.map((_,i)=>i===LABELS.length-1?O:O+'99'),borderRadius:6}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},
    tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('cDebtors'),{{
  type:'line',
  data:{{labels:LABELS,datasets:[{{label:'Outstanding Debtors',data:{debtors_arr},
    borderColor:R,backgroundColor:R+'18',fill:true,tension:0.4,pointRadius:5}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},
    tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('cOverdue'),{{
  type:'line',
  data:{{labels:LABELS,datasets:[
    {{label:'% Overdue >90d',data:{overdue_arr},borderColor:O,backgroundColor:O+'18',fill:true,tension:0.4,pointRadius:5}},
    {{label:'10% threshold',data:LABELS.map(()=>10),borderColor:R,borderDash:[6,4],borderWidth:1.5,pointRadius:0,fill:false}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom',labels:{{boxWidth:12,padding:12}}}},
      tooltip:{{callbacks:{{label:c=>c.parsed.y+'%'}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>v+'%'}},min:0,max:25}},x:{{grid:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('cReps'),{{
  type:'line',
  data:{{labels:LABELS,datasets:[
    {{label:'Aboo Cassim',   data:{aboo_arr},    borderColor:O,fill:false,tension:0.3,pointRadius:4}},
    {{label:'Amit Patel',    data:{amit_arr},    borderColor:T,fill:false,tension:0.3,pointRadius:4}},
    {{label:'Bhadresh Vallabh',data:{bhadresh_arr},borderColor:B,fill:false,tension:0.3,pointRadius:4}},
    {{label:'Nikhil Panchal',data:{nikhil_arr},  borderColor:P,fill:false,tension:0.3,pointRadius:4}},
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom',labels:{{boxWidth:12,padding:12}}}},
      tooltip:{{callbacks:{{label:c=>c.dataset.label+': '+c.parsed.y+'%'}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>v+'%'}},min:0,max:110}},x:{{grid:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('cRB'),{{
  type:'line',
  data:{{labels:LABELS,datasets:[
    {{label:'Above Rock Bottom %',data:{rb_arr},borderColor:G2,backgroundColor:G2+'22',fill:true,tension:0.4,pointRadius:5}},
    {{label:'8% healthy target',data:LABELS.map(()=>8),borderColor:'#aaa',borderDash:[6,4],borderWidth:1.5,pointRadius:0,fill:false}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:true,position:'bottom',labels:{{boxWidth:12,padding:12}}}},
      tooltip:{{callbacks:{{label:c=>c.parsed.y+'%'}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>v+'%'}},min:0,max:15}},x:{{grid:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('cCredits'),{{
  type:'bar',
  data:{{labels:LABELS,datasets:[{{label:'Credits (R)',data:{credits_arr},
    backgroundColor:O+'99',borderRadius:4}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},
    tooltip:{{callbacks:{{label:c=>'R '+c.parsed.y.toLocaleString('en-ZA')}}}}}},
    scales:{{y:{{grid,ticks:{{callback:v=>'R'+(v/1000).toFixed(0)+'K'}}}},x:{{grid:{{display:false}}}}}}}}
}});
</script>
</body>
</html>"""
    return html


# ── GIT PUSH ─────────────────────────────────────────────────────────────────

def git_push(dashboard_path: Path, n_snaps: int):
    cwd = dashboard_path.parent
    msg = f"Auto-update KPI dashboard — {n_snaps} snapshots — {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    def run(cmd):
        result = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  [WARN] {' '.join(cmd)}: {result.stderr.strip()}")
        return result.returncode == 0

    run(["git", "config", "user.email", "auto@olympic-paints.local"])
    run(["git", "config", "user.name",  "Olympic KPI Bot"])
    run(["git", "add", "KPI Dashboard.html"])
    run(["git", "add", "index.html"])
    run(["git", "add", "build_kpi_dashboard.py"])
    run(["git", "commit", "-m", msg])
    ok = run(["git", "push", "origin", "master"])
    if ok:
        print(f"  ✓ Pushed to GitHub ({n_snaps} snapshots)")
    else:
        # Try pushing with branch name main as fallback
        run(["git", "push", "-u", "origin", "master"])


# ── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Building KPI Dashboard…")

    pdfs = find_snapshots()
    print(f"  Found {len(pdfs)} Sales Summary PDF(s)")

    snapshots = []
    for dt, path in pdfs:
        print(f"  Extracting: {path.name}")
        kpi = extract_kpi(path)
        if kpi.get("date"):
            snapshots.append(kpi)

    if not snapshots:
        print("  [ERROR] No data extracted — aborting.")
        sys.exit(1)

    html = build_html(snapshots)
    DASHBOARD.write_text(html, encoding="utf-8")
    # index.html is what GitHub Pages serves from the root
    (BASE_DIR / "index.html").write_text(html, encoding="utf-8")
    print(f"  ✓ Dashboard written → {DASHBOARD}")

    git_push(DASHBOARD, len(snapshots))
    print("  Done.")


if __name__ == "__main__":
    main()
