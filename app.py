from flask import Flask, render_template_string, request, redirect, url_for, flash, session, send_file
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import re, io, json, base64

app = Flask(__name__)
app.secret_key = "basketkolcz2025"

# ══════════════════════════════════════════════════════════════════════════════
# PARSER
# ══════════════════════════════════════════════════════════════════════════════

ACTION_2PM  = {"2","2+1","2+0","2D","2D+1"}
ACTION_3PM  = {"3","3+1","3+0"}
ACTION_BR   = {"BR"}
ACTION_P    = {"P"}
ACTION_F    = {"F"}
ACTION_TIP  = {"2D","0/2D","2D+1"}

def extract_ft(code):
    m = re.match(r'^(\d+)/(\d+)W', code)
    if m: return int(m.group(1)), int(m.group(2))
    m2 = re.search(r'(\d+)/(\d+)W', code)
    if m2: return int(m2.group(1)), int(m2.group(2))
    return 0, 0

def safe_row(row, idx):
    return row[idx] if len(row) > idx and row[idx] is not None else None

def time_bucket(t):
    if t == 0:     return "0s"
    if t <= 4:     return "1-4s"
    if t <= 8:     return "5-8s"
    if t <= 12:    return "9-12s"
    if t <= 16:    return "13-16s"
    if t <= 20:    return "17-20s"
    return "21-24s"

BUCKETS = ["0s","1-4s","5-8s","9-12s","13-16s","17-20s","21-24s"]

ZONE_NAMES = {
    1:"Lewy blok",2:"Prawy blok",3:"Lewy łokieć",4:"Środek farby",5:"Prawy łokieć",
    6:"Lewy baseline",7:"Lewe skrzydło",8:"Góra klucza",9:"Prawe skrzydło",10:"Prawy baseline",
    11:"Lewy ekst.",12:"Prawy ekst.",13:"Daleki lewy",14:"Daleki prawy",
    15:"Lewy narożnik 3PT",16:"Lewe skrzydło 3PT",17:"Góra łuku 3PT",
    18:"Prawe skrzydło 3PT",19:"Prawy narożnik 3PT"
}
ZONE_AREA = {
    1:"Pod koszem",2:"Pod koszem",3:"Farba",4:"Farba",5:"Farba",
    6:"Mid-range",7:"Mid-range",8:"Mid-range",9:"Mid-range",10:"Mid-range",
    11:"Mid-range",12:"Mid-range",13:"Mid-range",14:"Mid-range",
    15:"3PT",16:"3PT",17:"3PT",18:"3PT",19:"3PT"
}
ZONE_PTS = {z: 3 if ZONE_AREA.get(z)=="3PT" else 2 for z in range(1,20)}

def parse_team_sheet(ws):
    stats = {
        "quarter": defaultdict(lambda: {
            "ftm":0,"fta":0,"2pm":0,"2pa":0,"3pm":0,"3pa":0,
            "tip_made":0,"tip_miss":0,"and1_2":0,"and1_3":0,
            "br":0,"p":0,"fd":0,"poss":0,"acts":0,"pts":0
        }),
        "players": defaultdict(lambda: {
            "2pm":0,"2pa":0,"3pm":0,"3pa":0,"ftm":0,"fta":0,
            "fd":0,"br":0,"finishes":0,"ast":0,"oreb":0,"dreb":0
        }),
        "lineups": defaultdict(lambda: {
            "poss":0,"acts":0,"pts":0,"2pm":0,"2pa":0,
            "3pm":0,"3pa":0,"ftm":0,"fta":0,"br":0,"fd":0,"tempo":0
        }),
        "zones": defaultdict(lambda: {"made":0,"miss":0}),
        "timing": {b: {"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}} for b in BUCKETS},
        "total_finishes": 0,
    }

    current_q = 1
    current_lineup = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v is not None for v in row[:4]):
            continue

        # Kwarta
        if row[0] is not None:
            try: current_q = int(str(row[0]).replace("*","").strip())
            except: current_q = 1

        # Zawodnicy na boisku (kolumny E-I = indeksy 4-8)
        for i in range(4, 9):
            val = safe_row(row, i)
            if val is not None:
                s = str(val).strip()
                if ";" in s:
                    parts = s.split(";")
                    try:
                        old_p, new_p = int(parts[0].strip()), int(parts[1].strip())
                        try:
                            idx = current_lineup.index(old_p)
                            current_lineup[idx] = new_p
                        except ValueError:
                            current_lineup.append(new_p)
                    except: pass
                else:
                    try:
                        p = int(s)
                        if p not in current_lineup:
                            current_lineup.append(p)
                    except: pass

        if len(current_lineup) > 5:
            current_lineup = current_lineup[-5:]

        lineup_key = ";".join(str(p) for p in sorted(current_lineup))

        # Wartości kolumn
        raw_b = str(row[1]) if row[1] is not None else ""
        raw_c = str(row[2]) if row[2] is not None else ""
        raw_d = str(row[3]) if row[3] is not None else ""
        raw_k = str(safe_row(row,10)) if safe_row(row,10) is not None else ""
        raw_l = str(safe_row(row,11)) if safe_row(row,11) is not None else ""
        raw_m = str(safe_row(row,12)) if safe_row(row,12) is not None else ""
        raw_n = str(safe_row(row,13)) if safe_row(row,13) is not None else ""

        codes     = [c.strip() for c in raw_c.split(";") if c.strip()]
        times     = [t.strip() for t in raw_b.replace(";",",").split(",") if t.strip()]
        zones     = [z.strip() for z in raw_d.split(";") if z.strip()]
        finishers = [f.strip() for f in raw_k.split(";") if f.strip()]
        assists   = [a.strip() for a in raw_l.split(";") if a.strip()]
        orebs     = [o.strip() for o in raw_m.split(";") if o.strip()]
        drebs     = [d.strip() for d in raw_n.split(";") if d.strip()]

        q = stats["quarter"][current_q]
        q["poss"] += 1
        stats["lineups"][lineup_key]["poss"] += 1

        for ai, code in enumerate(codes):
            q["acts"] += 1
            stats["lineups"][lineup_key]["acts"] += 1

            # Czas → bucket
            t_val = 0
            if ai < len(times):
                try: t_val = float(times[ai])
                except: pass
            stats["lineups"][lineup_key]["tempo"] += t_val
            bucket = time_bucket(t_val)

            # Strefa
            zone = 0
            if ai < len(zones):
                try: zone = int(zones[ai])
                except: pass

            # Zawodnik kończący
            finisher = None
            if ai < len(finishers):
                try: finisher = int(finishers[ai])
                except: pass

            # Asysta
            assister = None
            if ai < len(assists):
                try: assister = int(assists[ai])
                except: pass

            # Zbiorki
            orebler = None
            if ai < len(orebs):
                try: orebler = int(orebs[ai])
                except: pass

            drebler = None
            if ai < len(drebs):
                try: drebler = int(drebs[ai])
                except: pass

            pts = 0
            is_2pt_made = is_2pt_miss = is_3pt_made = is_3pt_miss = False

            # ── Klasyfikacja kodu ──────────────────────────────────────────
            if code in ACTION_2PM:
                q["2pm"] += 1; q["2pa"] += 1; pts = 2
                stats["lineups"][lineup_key]["2pm"] += 1
                stats["lineups"][lineup_key]["2pa"] += 1
                is_2pt_made = True
                if zone: stats["zones"][zone]["made"] += 1
                if code in ACTION_TIP: q["tip_made"] += 1
                if code in ("2+1","2D+1"): q["and1_2"] += 1
                if finisher:
                    stats["players"][finisher]["2pm"] += 1
                    stats["players"][finisher]["2pa"] += 1
                    stats["players"][finisher]["finishes"] += 1
                    stats["total_finishes"] += 1
                    if assister: stats["players"][assister]["ast"] += 1

            elif code in ("0/2","0/2D"):
                q["2pa"] += 1
                stats["lineups"][lineup_key]["2pa"] += 1
                is_2pt_miss = True
                if zone: stats["zones"][zone]["miss"] += 1
                if code == "0/2D": q["tip_miss"] += 1
                if finisher:
                    stats["players"][finisher]["2pa"] += 1
                    stats["players"][finisher]["finishes"] += 1
                    stats["total_finishes"] += 1
                if orebler: stats["players"][orebler]["oreb"] += 1

            elif code in ACTION_3PM:
                q["3pm"] += 1; q["3pa"] += 1; pts = 3
                stats["lineups"][lineup_key]["3pm"] += 1
                stats["lineups"][lineup_key]["3pa"] += 1
                is_3pt_made = True
                if zone: stats["zones"][zone]["made"] += 1
                if code in ("3+1",): q["and1_3"] += 1
                if finisher:
                    stats["players"][finisher]["3pm"] += 1
                    stats["players"][finisher]["3pa"] += 1
                    stats["players"][finisher]["finishes"] += 1
                    stats["total_finishes"] += 1
                    if assister: stats["players"][assister]["ast"] += 1

            elif code == "0/3":
                q["3pa"] += 1
                stats["lineups"][lineup_key]["3pa"] += 1
                is_3pt_miss = True
                if zone: stats["zones"][zone]["miss"] += 1
                if finisher:
                    stats["players"][finisher]["3pa"] += 1
                    stats["players"][finisher]["finishes"] += 1
                    stats["total_finishes"] += 1
                if orebler: stats["players"][orebler]["oreb"] += 1

            elif code in ACTION_BR:
                q["br"] += 1
                stats["lineups"][lineup_key]["br"] += 1
                if finisher: stats["players"][finisher]["br"] += 1
                if drebler: stats["players"][drebler]["dreb"] += 1

            elif code in ACTION_P:
                q["p"] += 1

            elif code in ACTION_F:
                q["fd"] += 1
                stats["lineups"][lineup_key]["fd"] += 1

            # Rzuty wolne
            ftm, fta = extract_ft(code)
            if fta > 0:
                q["ftm"] += ftm; q["fta"] += fta; pts += ftm
                stats["lineups"][lineup_key]["ftm"] += ftm
                stats["lineups"][lineup_key]["fta"] += fta
                if finisher:
                    stats["players"][finisher]["ftm"] += ftm
                    stats["players"][finisher]["fta"] += fta
                    stats["players"][finisher]["fd"] += 1

            # Punkty
            q["pts"] += pts
            stats["lineups"][lineup_key]["pts"] += pts

            # Timing
            if is_2pt_made:
                stats["timing"][bucket]["2PT"]["made"] += 1
            elif is_2pt_miss:
                stats["timing"][bucket]["2PT"]["miss"] += 1
            elif is_3pt_made:
                stats["timing"][bucket]["3PT"]["made"] += 1
            elif is_3pt_miss:
                stats["timing"][bucket]["3PT"]["miss"] += 1

    return stats

def calc_kpi(d):
    fga = d.get("2pa",0) + d.get("3pa",0)
    pts = d.get("pts",0)
    poss = max(d.get("poss",1), 1)
    fta = d.get("fta",0)
    ftm = d.get("ftm",0)
    pm2 = d.get("2pm",0); pa2 = d.get("2pa",0)
    pm3 = d.get("3pm",0); pa3 = d.get("3pa",0)

    def pct(n,d): return f"{n/d:.1%}" if d else "-"
    def dec(v,f=2): return f"{v:.{f}f}" if v is not None else "-"

    efg  = (pm2+1.5*pm3)/fga if fga else None
    ts   = pts/(2*(fga+0.44*fta)) if (fga+fta) else None
    ortg = pts*100/poss
    topct= d.get("br",0)/poss
    ftr  = fta/fga if fga else None
    ppp  = pts/poss
    acts = max(d.get("acts",1),1)

    return {
        "efg":    pct(pm2+1.5*pm3, fga) if fga else "-",
        "ts":     f"{ts:.1%}" if ts else "-",
        "ortg":   f"{ortg:.1f}",
        "topct":  f"{topct:.1%}",
        "ftr":    f"{ftr:.2f}" if ftr else "-",
        "ppp":    f"{ppp:.2f}",
        "p2_pct": pct(pm2,pa2),
        "p3_pct": pct(pm3,pa3),
        "ft_pct": pct(ftm,fta),
        "ppa":    f"{pts/acts:.2f}",
    }

def suma_quarters(stats):
    s = defaultdict(int)
    for qn in [1,2,3,4]:
        qd = stats["quarter"].get(qn, {})
        for k,v in qd.items():
            s[k] += v
    return dict(s)

# ══════════════════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════════════════

CSS = """
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
body{background:#f4f6fb;font-family:'Segoe UI',Arial,sans-serif;font-size:.9rem}
.navbar{background:#1a2b4a!important}
.navbar-brand{font-weight:700;font-size:1.2rem;color:#fff!important}
.brand-dot{color:#EF9F27}
.card{border:none;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.stat-card{background:#fff;border-radius:12px;padding:1rem;text-align:center;border:1px solid #eee}
.stat-value{font-size:1.8rem;font-weight:700;color:#1a2b4a;line-height:1.1}
.stat-value.sm{font-size:1.3rem}
.stat-label{font-size:.7rem;color:#999;text-transform:uppercase;letter-spacing:.5px;margin-top:.2rem}
.table th{background:#1a2b4a;color:#fff;font-size:.78rem;font-weight:600;border:none;padding:.5rem .6rem}
.table td{font-size:.82rem;vertical-align:middle;padding:.4rem .6rem}
.table-hover tbody tr:hover{background:#f0f4ff}
.qpill{display:inline-block;padding:1px 8px;border-radius:20px;font-size:.75rem;font-weight:600}
.q1{background:#e8f5e9;color:#1a6b3c}.q2{background:#e3f2fd;color:#0c447c}
.q3{background:#fff3e0;color:#854f0b}.q4{background:#fce4ec;color:#8b1a1a}
.qs{background:#1a2b4a;color:#fff}
.nav-tabs .nav-link{color:#555;font-size:.85rem;padding:.5rem .9rem}
.nav-tabs .nav-link.active{color:#1a2b4a;font-weight:600;border-bottom:2px solid #1a2b4a;background:transparent}
.section-title{font-size:.68rem;text-transform:uppercase;letter-spacing:1px;color:#aaa;font-weight:600;margin-bottom:.6rem;padding-bottom:.3rem;border-bottom:1px solid #f0f0f0}
.upload-zone{border:2px dashed #1a2b4a;border-radius:16px;padding:3rem;text-align:center;background:#fff;cursor:pointer;transition:.2s}
.upload-zone:hover{background:#f0f4ff}
.badge-area{font-size:.7rem;padding:2px 7px;border-radius:10px;font-weight:600}
.zone-paint{background:#c8e6c9;color:#1b5e20}
.zone-mid{background:#fff9c4;color:#f57f17}
.zone-3pt{background:#fce4ec;color:#880e4f}
.zone-under{background:#b2dfdb;color:#004d40}
.hero{background:linear-gradient(135deg,#1a2b4a,#2e5090);color:#fff;border-radius:16px;padding:1.75rem 2rem;margin-bottom:1.5rem}
.kpi-highlight{background:#fff8e1;border:1px solid #ffe082;border-radius:8px;padding:.5rem .75rem;display:inline-block;margin:.2rem}
.timing-bar{height:12px;border-radius:6px;background:#e3f2fd;overflow:hidden;min-width:60px}
.timing-fill{height:100%;border-radius:6px;background:#1a6b3c}
.timing-fill.miss{background:#e57373}
@media(max-width:576px){.stat-value{font-size:1.3rem}.stat-value.sm{font-size:1.1rem}}
</style>
"""

NAV = """
<nav class="navbar navbar-dark mb-3 px-3">
  <a class="navbar-brand" href="/"><span class="brand-dot">&#9679;</span> Basket Kołcz</a>
  <span style="color:#ffffff88;font-size:.8rem">Analytics Platform</span>
</nav>
"""

def base(content, scripts=""):
    return f"""<!DOCTYPE html><html lang="pl"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Basket Kołcz</title>{CSS}</head><body>{NAV}
<div class="container-fluid px-3 pb-5" style="max-width:1200px">{content}</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
{scripts}</body></html>"""

def flash_html():
    return """{% with messages = get_flashed_messages(with_categories=true) %}
{% if messages %}{% for cat,msg in messages %}
<div class="alert alert-{{ 'danger' if cat=='error' else 'info' }} alert-dismissible fade show mb-3">
{{ msg }}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>
{% endfor %}{% endif %}{% endwith %}"""

# ── Team panel builder ─────────────────────────────────────────────────────────

def build_team_panel(team_name, stats, team_id):
    suma = suma_quarters(stats)
    kpi  = calc_kpi(suma)

    # KPI row
    kpi_cards = ""
    for val, lbl in [
        (suma.get("pts",0), "Punkty"),
        (kpi["efg"],        "eFG%"),
        (kpi["ts"],         "TS%"),
        (kpi["ortg"],       "ORtg"),
        (kpi["ppp"],        "Pkt/Pos"),
        (kpi["topct"],      "TO%"),
        (kpi["p2_pct"],     "2PT%"),
        (kpi["p3_pct"],     "3PT%"),
        (kpi["ft_pct"],     "FT%"),
        (kpi["ftr"],        "FT Rate"),
    ]:
        sz = "" if str(val) == str(suma.get("pts",0)) else " sm"
        kpi_cards += f'<div class="col-6 col-sm-4 col-md-2"><div class="stat-card"><div class="stat-value{sz}">{val}</div><div class="stat-label">{lbl}</div></div></div>'

    # Per kwarta tabela
    q_rows = ""
    for qn in [1,2,3,4]:
        qd = stats["quarter"].get(qn, {})
        qk = calc_kpi(qd)
        cl = f"q{qn}"
        q_rows += f"""<tr>
<td><span class="qpill {cl}">{qn}Q</span></td>
<td><b>{qd.get('pts',0)}</b></td>
<td>{qd.get('2pm',0)}/{qd.get('2pa',0)}</td><td>{qk['p2_pct']}</td>
<td>{qd.get('3pm',0)}/{qd.get('3pa',0)}</td><td>{qk['p3_pct']}</td>
<td>{qd.get('ftm',0)}/{qd.get('fta',0)}</td><td>{qk['ft_pct']}</td>
<td>{qd.get('br',0)}</td><td>{qd.get('poss',0)}</td>
<td>{qk['ppp']}</td><td>{qk['efg']}</td><td>{qk['ortg']}</td>
</tr>"""
    sk = calc_kpi(suma)
    q_rows += f"""<tr style="font-weight:700;background:#f0f4ff">
<td><span class="qpill qs">SUMA</span></td>
<td><b>{suma.get('pts',0)}</b></td>
<td>{suma.get('2pm',0)}/{suma.get('2pa',0)}</td><td>{sk['p2_pct']}</td>
<td>{suma.get('3pm',0)}/{suma.get('3pa',0)}</td><td>{sk['p3_pct']}</td>
<td>{suma.get('ftm',0)}/{suma.get('fta',0)}</td><td>{sk['ft_pct']}</td>
<td>{suma.get('br',0)}</td><td>{suma.get('poss',0)}</td>
<td>{sk['ppp']}</td><td>{sk['efg']}</td><td>{sk['ortg']}</td>
</tr>"""

    # Zawodnicy
    p_rows = ""
    total_fin = max(stats["total_finishes"], 1)
    for pid in sorted(stats["players"].keys()):
        pd = stats["players"][pid]
        pk = calc_kpi(pd)
        pts_p = pd.get("2pm",0)*2 + pd.get("3pm",0)*3 + pd.get("ftm",0)
        fin_pct = f"{pd.get('finishes',0)/total_fin:.0%}"
        p_rows += f"""<tr>
<td class="fw-bold">#<b>{pid}</b></td>
<td class="fw-bold" style="color:#1a2b4a;font-size:.95rem">{pts_p}</td>
<td>{pd.get('2pm',0)}/{pd.get('2pa',0)}</td>
<td>{pd.get('3pm',0)}/{pd.get('3pa',0)}</td>
<td>{pd.get('ftm',0)}/{pd.get('fta',0)}</td>
<td><b>{pk['efg']}</b></td>
<td>{pk['ts']}</td>
<td>{pd.get('ast',0)}</td>
<td>{pd.get('oreb',0)}</td>
<td>{pd.get('dreb',0)}</td>
<td>{pd.get('br',0)}</td>
<td>{pd.get('fd',0)}</td>
<td>{pd.get('finishes',0)} <small class="text-muted">({fin_pct})</small></td>
</tr>"""

    # Składy
    l_rows = ""
    sorted_lineups = sorted(
        [(k,v) for k,v in stats["lineups"].items() if v["poss"] >= 2],
        key=lambda x: x[1]["pts"], reverse=True
    )
    for lid, ld in sorted_lineups[:20]:
        lk = calc_kpi(ld)
        avg_tempo = f"{ld['tempo']/max(ld['poss'],1):.1f}s" if ld.get("tempo") else "-"
        l_rows += f"""<tr>
<td style="font-family:monospace;font-size:.78rem">{lid}</td>
<td>{ld['poss']}</td>
<td><b>{ld['pts']}</b></td>
<td>{lk['ppp']}</td>
<td>{lk['efg']}</td>
<td style="color:#1a6b3c;font-weight:600">{lk['ortg']}</td>
<td>{ld.get('br',0)}</td>
<td>{avg_tempo}</td>
</tr>"""

    # Shot timing tabela
    tim_rows = ""
    for b in BUCKETS:
        bd = stats["timing"][b]
        made2 = bd["2PT"]["made"]; miss2 = bd["2PT"]["miss"]
        made3 = bd["3PT"]["made"]; miss3 = bd["3PT"]["miss"]
        tot_made = made2+made3; tot_miss = miss2+miss3; tot = tot_made+tot_miss
        eff = f"{tot_made/tot:.0%}" if tot else "-"
        bar_w = int(tot_made/(tot)*100) if tot else 0
        bar = f'<div class="timing-bar"><div class="timing-fill" style="width:{bar_w}%"></div></div>'
        tim_rows += f"""<tr>
<td><b>{b}</b></td>
<td>{made2}/{made2+miss2}</td>
<td>{made3}/{made3+miss3}</td>
<td><b>{tot_made}/{tot}</b></td>
<td>{eff} {bar}</td>
</tr>"""

    # Court zones tabela
    zone_rows = ""
    for z in range(1, 20):
        zd = stats["zones"].get(z, {"made":0,"miss":0})
        made = zd["made"]; miss = zd["miss"]; att = made+miss
        eff = f"{made/att:.0%}" if att else "-"
        pps_v = made*ZONE_PTS[z]/att if att else None
        pps = f"{pps_v:.2f}" if pps_v else "-"
        area = ZONE_AREA.get(z,"")
        area_cl = {"Pod koszem":"zone-under","Farba":"zone-paint","Mid-range":"zone-mid","3PT":"zone-3pt"}.get(area,"")
        if att == 0: continue
        zone_rows += f"""<tr>
<td><b>{z}</b></td>
<td style="font-size:.8rem">{ZONE_NAMES.get(z,'')}</td>
<td><span class="badge-area {area_cl}">{area}</span></td>
<td>{made}</td><td>{miss}</td><td>{att}</td>
<td><b>{eff}</b></td><td style="color:#1a6b3c;font-weight:600">{pps}</td>
</tr>"""

    html = f"""
<div class="row g-2 my-2">{kpi_cards}</div>

<ul class="nav nav-tabs mt-3 mb-1" id="innerTabs{team_id}">
  <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tQuarter{team_id}">Per kwarta</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tPlayers{team_id}">Zawodnicy</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tLineups{team_id}">Składy</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tTiming{team_id}">Shot Timing</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tZones{team_id}">Strefy</button></li>
</ul>

<div class="tab-content">

<div class="tab-pane fade show active" id="tQuarter{team_id}">
<div class="card mt-2"><div class="card-body p-2">
<div class="table-responsive">
<table class="table table-hover mb-0">
<thead><tr><th>Q</th><th>PKT</th><th>2PM/A</th><th>2P%</th><th>3PM/A</th><th>3P%</th><th>FTM/A</th><th>FT%</th><th>BR</th><th>POSS</th><th>PPP</th><th>eFG%</th><th>ORtg</th></tr></thead>
<tbody>{q_rows}</tbody></table></div></div></div>
</div>

<div class="tab-pane fade" id="tPlayers{team_id}">
<div class="card mt-2"><div class="card-body p-2">
<div class="table-responsive">
<table class="table table-hover mb-0">
<thead><tr><th>#</th><th>PTS</th><th>2PM/A</th><th>3PM/A</th><th>FTM/A</th><th>eFG%</th><th>TS%</th><th>AST</th><th>OREB</th><th>DREB</th><th>BR</th><th>FD</th><th>Wykończenia</th></tr></thead>
<tbody>{p_rows if p_rows else '<tr><td colspan="13" class="text-center text-muted py-3">Brak danych zawodników — dodaj kolumnę K (zawodnik kończący) w pliku zapis</td></tr>'}</tbody>
</table></div></div></div>
</div>

<div class="tab-pane fade" id="tLineups{team_id}">
<div class="card mt-2"><div class="card-body p-2">
<div class="table-responsive">
<table class="table table-hover mb-0">
<thead><tr><th>Skład (nr koszulek)</th><th>POSS</th><th>PKT</th><th>PPP</th><th>eFG%</th><th>ORtg</th><th>BR</th><th>Śr. tempo</th></tr></thead>
<tbody>{l_rows if l_rows else '<tr><td colspan="8" class="text-center text-muted py-3">Brak składów z min. 2 posiadaniami</td></tr>'}</tbody>
</table></div></div></div>
</div>

<div class="tab-pane fade" id="tTiming{team_id}">
<div class="card mt-2"><div class="card-body p-2">
<p class="text-muted mb-2" style="font-size:.8rem">Czas trwania posiadania przed oddaniem rzutu (z zegara 24s)</p>
<div class="table-responsive">
<table class="table table-hover mb-0">
<thead><tr><th>Czas posiadania</th><th>2PT Made/Att</th><th>3PT Made/Att</th><th>Razem Made/Att</th><th>Skuteczność</th></tr></thead>
<tbody>{tim_rows}</tbody></table></div></div></div>
</div>

<div class="tab-pane fade" id="tZones{team_id}">
<div class="card mt-2"><div class="card-body p-2">
<p class="text-muted mb-2" style="font-size:.8rem">PPP = punkty na próbę rzutową z danej strefy</p>
<div class="table-responsive">
<table class="table table-hover mb-0">
<thead><tr><th>#</th><th>Strefa</th><th>Obszar</th><th>Celne</th><th>Niecelne</th><th>Próby</th><th>Eff%</th><th>PPP</th></tr></thead>
<tbody>{zone_rows if zone_rows else '<tr><td colspan="8" class="text-center text-muted py-3">Brak danych stref — dodaj kolumnę D (strefa) w pliku zapis</td></tr>'}</tbody>
</table></div></div></div>
</div>

</div>"""
    return html

# ══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    content = """
<div class="hero">
  <h1 class="fw-bold mb-1" style="font-size:1.75rem">Basket Kołcz Analytics</h1>
  <p class="mb-0" style="opacity:.8">Wgraj plik zapis.xlsx — otrzymasz pełny raport w kilka sekund</p>
</div>
<div class="row justify-content-center">
<div class="col-lg-7">
  <form method="POST" action="/upload" enctype="multipart/form-data">
    <div class="upload-zone" onclick="document.getElementById('fup').click()">
      <div style="font-size:2.5rem;margin-bottom:.75rem">📊</div>
      <h5 class="fw-bold mb-1" style="color:#1a2b4a">Wgraj plik zapis.xlsx</h5>
      <p class="text-muted mb-0" style="font-size:.85rem">Kliknij lub przeciągnij plik</p>
      <input type="file" id="fup" name="file" accept=".xlsx" class="d-none" onchange="this.form.submit()">
    </div>
  </form>

  <div class="card mt-3 p-3">
    <div style="font-size:.8rem;font-weight:700;color:#1a2b4a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:.75rem">
      📥 Pobierz szablony
    </div>
    <div class="row g-2">
      <div class="col-6">
        <a href="/template/zapis" class="btn btn-outline-primary w-100" style="font-size:.82rem">
          📝 Zapis meczu<br>
          <small class="text-muted" style="font-size:.72rem">Pusty arkusz do kodowania</small>
        </a>
      </div>
      <div class="col-6">
        <a href="/template/szablon" class="btn btn-outline-success w-100" style="font-size:.82rem">
          📋 Szablon raportu<br>
          <small class="text-muted" style="font-size:.72rem">Pusty szablon statystyk</small>
        </a>
      </div>
    </div>
  </div>

  <div class="card mt-2 p-3" style="font-size:.82rem;color:#666">
    <b>Format pliku:</b> Excel (.xlsx) z dwoma arkuszami — jedna drużyna = jeden arkusz.<br>
    Obsługuje stary format (11 kolumn) i nowy (14 kolumn z asystami i zbiorkami).
  </div>
</div></div>"""
    from flask import render_template_string
    return render_template_string(base(content))

@app.route("/upload", methods=["POST"])
def upload():
    from flask import render_template_string
    if "file" not in request.files:
        flash("Nie wybrano pliku", "error")
        return redirect(url_for("index"))

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        flash("Plik musi być w formacie .xlsx", "error")
        return redirect(url_for("index"))

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        data_sheets = [s for s in wb.sheetnames if s.upper() not in ("META","KODY","LEGENDA")]

        if len(data_sheets) < 1:
            flash("Plik musi zawierać co najmniej jeden arkusz z danymi", "error")
            return redirect(url_for("index"))

        name_a = data_sheets[0]
        name_b = data_sheets[1] if len(data_sheets) > 1 else None

        stats_a = parse_team_sheet(wb[name_a])
        stats_b = parse_team_sheet(wb[name_b]) if name_b else None

        # Zapisz dane do sesji dla eksportu
        session['name_a'] = name_a
        session['name_b'] = name_b
        session['suma_a'] = json.dumps(dict(suma_quarters(stats_a)))
        session['suma_b'] = json.dumps(dict(suma_quarters(stats_b))) if stats_b else None
        session['kpi_a']  = json.dumps(calc_kpi(suma_quarters(stats_a)))
        session['kpi_b']  = json.dumps(calc_kpi(suma_quarters(stats_b))) if stats_b else None
        # Dane per kwarta
        session['quarters_a'] = json.dumps({str(q): dict(stats_a["quarter"].get(q,{})) for q in [1,2,3,4]})
        session['quarters_b'] = json.dumps({str(q): dict(stats_b["quarter"].get(q,{})) for q in [1,2,3,4]}) if stats_b else None
        # Zawodnicy
        session['players_a'] = json.dumps({str(k): dict(v) for k,v in stats_a["players"].items()})
        session['players_b'] = json.dumps({str(k): dict(v) for k,v in stats_b["players"].items()}) if stats_b else None
        # Shot timing
        session['timing_a'] = json.dumps({b: stats_a["timing"][b] for b in BUCKETS})
        session['timing_b'] = json.dumps({b: stats_b["timing"][b] for b in BUCKETS}) if stats_b else None

        suma_a = suma_quarters(stats_a)
        suma_b = suma_quarters(stats_b) if stats_b else None
        kpi_a  = calc_kpi(suma_a)
        kpi_b  = calc_kpi(suma_b) if suma_b else None

        # Wykres punktów per kwarta
        pts_a = [stats_a["quarter"].get(q,{}).get("pts",0) for q in [1,2,3,4]]
        pts_b = [stats_b["quarter"].get(q,{}).get("pts",0) for q in [1,2,3,4]] if stats_b else [0,0,0,0]

        # Panel A
        panel_a = build_team_panel(name_a, stats_a, "A")
        panel_b = build_team_panel(name_b, stats_b, "B") if stats_b else ""

        # ── Tabela Kluczowych Metryk ──────────────────────────────────────────
        def bar_html(va, vb, higher_is_better=True):
            """Pasek porównawczy — wizualizuje przewagę"""
            try:
                fa = float(str(va).replace('%','').replace('-','0') or 0)
                fb = float(str(vb).replace('%','').replace('-','0') or 0)
            except: return ""
            total = fa + fb
            if total == 0: return ""
            pct_a = int(fa / total * 100)
            pct_b = 100 - pct_a
            # kto wygrywa
            if fa > fb:
                col_a, col_b = '#1a6b3c', '#e0e0e0'
            elif fb > fa:
                col_a, col_b = '#e0e0e0', '#8b1a1a'
            else:
                col_a = col_b = '#aaa'
            return f'''<div style="display:flex;height:6px;border-radius:3px;overflow:hidden;margin-top:4px">
                <div style="width:{pct_a}%;background:{col_a}"></div>
                <div style="width:{pct_b}%;background:{col_b}"></div>
            </div>'''

        cmp_rows = ""
        if kpi_b:
            metrics = [
                ("Punkty",           suma_a.get("pts",0),    suma_b.get("pts",0),    True,  "Łączna liczba punktów zdobytych w meczu"),
                ("Posiadania",        suma_a.get("poss",0),   suma_b.get("poss",0),   True,  "Liczba posiadań piłki"),
                ("eFG%",              kpi_a["efg"],           kpi_b["efg"],           True,  "Efektywny % rzutów z pola (uwzględnia wagę trójki)"),
                ("True Shooting%",    kpi_a["ts"],            kpi_b["ts"],            True,  "Prawdziwy % skuteczności (2PT + 3PT + FT)"),
                ("Offensive Rating",  kpi_a["ortg"],          kpi_b["ortg"],          True,  "Punkty na 100 posiadań"),
                ("Pkt / Posiadanie",  kpi_a["ppp"],           kpi_b["ppp"],           True,  "Średnia punktów na jedno posiadanie"),
                ("2PT%",              kpi_a["p2_pct"],        kpi_b["p2_pct"],        True,  "Skuteczność rzutów za 2 punkty"),
                ("3PT%",              kpi_a["p3_pct"],        kpi_b["p3_pct"],        True,  "Skuteczność rzutów za 3 punkty"),
                ("FT%",               kpi_a["ft_pct"],        kpi_b["ft_pct"],        True,  "Skuteczność rzutów wolnych"),
                ("FT Rate",           kpi_a["ftr"],           kpi_b["ftr"],           True,  "Stosunek rzutów wolnych do rzutów z gry"),
                ("Turnover%",         kpi_a["topct"],         kpi_b["topct"],         False, "% posiadań zakończonych stratą (niższy = lepszy)"),
                ("Straty (BR)",       suma_a.get("br",0),     suma_b.get("br",0),     False, "Liczba strat"),
                ("Faule wymuszone",   suma_a.get("fd",0),     suma_b.get("fd",0),     True,  "Liczba wymuszonych fauli"),
                ("2PM/A",             f"{suma_a.get('2pm',0)}/{suma_a.get('2pa',0)}", f"{suma_b.get('2pm',0)}/{suma_b.get('2pa',0)}", None, "Celne/próby za 2 punkty"),
                ("3PM/A",             f"{suma_a.get('3pm',0)}/{suma_a.get('3pa',0)}", f"{suma_b.get('3pm',0)}/{suma_b.get('3pa',0)}", None, "Celne/próby za 3 punkty"),
                ("FTM/A",             f"{suma_a.get('ftm',0)}/{suma_a.get('fta',0)}", f"{suma_b.get('ftm',0)}/{suma_b.get('fta',0)}", None, "Celne/próby rzutów wolnych"),
            ]
            for lbl, va, vb, hib, desc in metrics:
                bar = bar_html(va, vb, hib) if hib is not None else ""
                # Podświetl wygraną stronę
                try:
                    fa2 = float(str(va).replace('%','').replace('/','').replace('-','0') or 0)
                    fb2 = float(str(vb).replace('%','').replace('/','').replace('-','0') or 0)
                    style_a = "font-weight:700;color:#1a6b3c" if (hib and fa2>fb2) or (hib==False and fa2<fb2) else "color:#555"
                    style_b = "font-weight:700;color:#8b1a1a" if (hib and fb2>fa2) or (hib==False and fb2<fa2) else "color:#555"
                except:
                    style_a = style_b = "color:#555"
                cmp_rows += f"""<tr>
                    <td style="font-size:.8rem">
                        <span style="font-weight:600">{lbl}</span>
                        <div style="font-size:.7rem;color:#aaa">{desc}</div>
                        {bar}
                    </td>
                    <td class="text-center" style="{style_a};font-size:.9rem">{va}</td>
                    <td class="text-center" style="{style_b};font-size:.9rem">{vb}</td>
                </tr>"""

        # ── Shot Timing porównanie ────────────────────────────────────────────
        tim_cmp_rows = ""
        if stats_b:
            for b in BUCKETS:
                ba = stats_a["timing"][b]
                bb2 = stats_b["timing"][b]
                made_a = ba["2PT"]["made"]+ba["3PT"]["made"]
                att_a  = made_a + ba["2PT"]["miss"]+ba["3PT"]["miss"]
                made_b = bb2["2PT"]["made"]+bb2["3PT"]["made"]
                att_b  = made_b + bb2["2PT"]["miss"]+bb2["3PT"]["miss"]
                eff_a = f"{made_a/att_a:.0%}" if att_a else "-"
                eff_b = f"{made_b/att_b:.0%}" if att_b else "-"
                # Pasek długości posiadań
                max_att = max(att_a, att_b, 1)
                bar_a = int(att_a/max_att*80)
                bar_b = int(att_b/max_att*80)
                tim_cmp_rows += f"""<tr>
                    <td class="fw-bold" style="font-size:.82rem">{b}</td>
                    <td class="text-center" style="font-size:.82rem">{made_a}/{att_a}</td>
                    <td class="text-center" style="font-size:.82rem;color:#1a6b3c;font-weight:600">{eff_a}</td>
                    <td style="padding:4px 8px">
                        <div style="height:8px;width:{bar_a}px;background:#1a6b3c;border-radius:4px;display:inline-block"></div>
                    </td>
                    <td style="padding:4px 8px">
                        <div style="height:8px;width:{bar_b}px;background:#8b1a1a;border-radius:4px;display:inline-block"></div>
                    </td>
                    <td class="text-center" style="font-size:.82rem;color:#8b1a1a;font-weight:600">{eff_b}</td>
                    <td class="text-center" style="font-size:.82rem">{made_b}/{att_b}</td>
                </tr>"""

        # ── Zakładka porównanie HTML ──────────────────────────────────────────
        if name_b:
            tab_cmp = (
                '<div class="tab-pane fade" id="tabCmp">'
                '<ul class="nav nav-tabs mt-2 mb-1" id="cmpTabs">'
                '<li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#cmpMetrics">Kluczowe Metryki</button></li>'
                '<li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#cmpTiming">Shot Timing</button></li>'
                '</ul>'
                '<div class="tab-content">'
                '<div class="tab-pane fade show active" id="cmpMetrics">'
                '<div class="card mt-2"><div class="card-body p-2">'
                '<div class="table-responsive">'
                '<table class="table table-sm table-hover mb-0">'
                '<thead><tr>'
                '<th>Metryka</th>'
                '<th class="text-center" style="color:#1a6b3c">' + name_a + '</th>'
                '<th class="text-center" style="color:#8b1a1a">' + name_b + '</th>'
                '</tr></thead>'
                '<tbody>' + cmp_rows + '</tbody>'
                '</table></div>'
                '</div></div>'
                '<div class="card mt-3"><div class="card-body p-3">'
                '<div class="section-title">Punkty per kwarta</div>'
                '<canvas id="qChart"></canvas>'
                '</div></div>'
                '</div>'
                '<div class="tab-pane fade" id="cmpTiming">'
                '<div class="card mt-2"><div class="card-body p-2">'
                '<p class="text-muted mb-2" style="font-size:.8rem">Porównanie skuteczności rzutów według czasu trwania posiadania (zegar 24s)</p>'
                '<div class="d-flex gap-3 mb-2" style="font-size:.78rem">'
                '<span><span style="display:inline-block;width:12px;height:8px;background:#1a6b3c;border-radius:2px;margin-right:4px"></span>' + name_a + '</span>'
                '<span><span style="display:inline-block;width:12px;height:8px;background:#8b1a1a;border-radius:2px;margin-right:4px"></span>' + name_b + '</span>'
                '</div>'
                '<div class="table-responsive">'
                '<table class="table table-hover mb-0">'
                '<thead><tr>'
                '<th>Czas</th>'
                '<th class="text-center" style="color:#1a6b3c">Celne/Próby</th>'
                '<th class="text-center" style="color:#1a6b3c">Eff%</th>'
                '<th>' + name_a + '</th>'
                '<th>' + name_b + '</th>'
                '<th class="text-center" style="color:#8b1a1a">Eff%</th>'
                '<th class="text-center" style="color:#8b1a1a">Celne/Próby</th>'
                '</tr></thead>'
                '<tbody>' + tim_cmp_rows + '</tbody>'
                '</table></div>'
                '</div></div>'
                '</div>'
                '</div>'
                '</div>'
            )
        else:
            tab_cmp = ""

        # Score hero
        score_a = suma_a.get("pts",0)
        score_b = suma_b.get("pts",0) if suma_b else "-"
        winner = name_a if isinstance(score_a,int) and isinstance(score_b,int) and score_a > score_b else (name_b if score_b != "-" and score_b > score_a else "")

        content = f"""
<div class="hero mb-3">
  <div class="d-flex justify-content-between align-items-center flex-wrap gap-2">
    <div>
      <h2 class="fw-bold mb-0">{name_a} vs {name_b or '—'}</h2>
      <p class="mb-0" style="opacity:.7;font-size:.85rem">Raport meczowy</p>
    </div>
    <div class="text-center">
      <div style="font-size:2.5rem;font-weight:700;letter-spacing:4px">{score_a} : {score_b}</div>
      {f'<small style="opacity:.7">{winner} wygrywa</small>' if winner else ''}
    </div>
    <a href="/" class="btn btn-outline-light btn-sm">← Nowy mecz</a>
    <div class="d-flex gap-2">
      <a href="/export/xlsx" class="btn btn-warning btn-sm fw-bold">⬇ Excel</a>
      <a href="/export/pdf" class="btn btn-danger btn-sm fw-bold">⬇ PDF</a>
    </div>
  </div>
</div>

<ul class="nav nav-tabs" id="mainTabs">
  <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tabA">{name_a}</button></li>
  {'<li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabB">' + name_b + '</button></li>' if name_b else ''}
  {'<li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabCmp">Porównanie</button></li>' if name_b else ''}
</ul>

<div class="tab-content">

<div class="tab-pane fade show active" id="tabA">
{panel_a}
</div>

{'<div class="tab-pane fade" id="tabB">' + panel_b + '</div>' if name_b else ''}

{tab_cmp}

</div>"""

        scripts = f"""<script>
const qc = document.getElementById('qChart');
if(qc) new Chart(qc, {{
  type:'bar',
  data:{{
    labels:['1Q','2Q','3Q','4Q'],
    datasets:[
      {{label:'{name_a}',data:{pts_a},backgroundColor:'#1a6b3c88',borderColor:'#1a6b3c',borderWidth:2,borderRadius:6}},
      {{label:'{name_b}',data:{pts_b},backgroundColor:'#8b1a1a88',borderColor:'#8b1a1a',borderWidth:2,borderRadius:6}}
    ]
  }},
  options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},scales:{{y:{{beginAtZero:true,grid:{{color:'#f0f0f0'}}}}}}}}
}});
</script>"""

        return render_template_string(base(content, scripts))

    except Exception as e:
        import traceback
        flash(f"Błąd: {str(e)}", "error")
        return redirect(url_for("index"))

@app.route("/export/xlsx")
def export_xlsx():
    try:
        name_a   = session.get('name_a', 'Drużyna A')
        name_b   = session.get('name_b', 'Drużyna B')
        suma_a   = json.loads(session.get('suma_a', '{}'))
        suma_b   = json.loads(session.get('suma_b', '{}')) if session.get('suma_b') else {}
        kpi_a    = json.loads(session.get('kpi_a',  '{}'))
        kpi_b    = json.loads(session.get('kpi_b',  '{}')) if session.get('kpi_b') else {}
        quarters_a = json.loads(session.get('quarters_a', '{}'))
        quarters_b = json.loads(session.get('quarters_b', '{}')) if session.get('quarters_b') else {}
        players_a  = json.loads(session.get('players_a', '{}'))
        players_b  = json.loads(session.get('players_b', '{}')) if session.get('players_b') else {}

        wb = openpyxl.Workbook()

        # Style
        HDR = PatternFill("solid", fgColor="1A2B4A")
        HDR_F = Font(color="FFFFFF", bold=True, size=10)
        SUB = PatternFill("solid", fgColor="E8F0FB")
        SUB_F = Font(bold=True, size=10, color="1A2B4A")
        GREEN_F = Font(bold=True, color="1A6B3C")
        RED_F   = Font(bold=True, color="8B1A1A")
        CTR = Alignment(horizontal="center", vertical="center")
        BORDER = Border(
            bottom=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC")
        )

        def set_hdr(ws, row, cols, labels):
            for i, lbl in enumerate(labels):
                c = ws.cell(row=row, column=cols+i, value=lbl)
                c.fill = HDR; c.font = HDR_F; c.alignment = CTR

        def style_row(ws, row, col_start, col_end, alt=False):
            fill = PatternFill("solid", fgColor="F5F8FF") if alt else None
            for col in range(col_start, col_end+1):
                c = ws.cell(row=row, column=col)
                c.border = BORDER; c.alignment = CTR
                if fill: c.fill = fill

        # ── Arkusz 1: OGÓLNE ──────────────────────────────────────────────────
        ws1 = wb.active; ws1.title = "OGÓLNE"
        ws1.column_dimensions['A'].width = 22
        for col in ['B','C','D','E','F','G','H','I','J','K','L']:
            ws1.column_dimensions[col].width = 12

        # Nagłówek meczu
        ws1.merge_cells('A1:L1')
        t = ws1['A1']; t.value = f"RAPORT MECZOWY — {name_a} vs {name_b}"
        t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=13)
        t.alignment = CTR

        # Wynik
        score_a = suma_a.get('pts', 0); score_b = suma_b.get('pts', 0)
        ws1['A2'] = "Wynik:"; ws1['A2'].font = Font(bold=True)
        ws1['B2'] = f"{score_a} : {score_b}"; ws1['B2'].font = Font(bold=True, size=13, color="1A2B4A")

        # Nagłówki tabeli
        row = 4
        set_hdr(ws1, row, 1, ["Metryka",
            name_a+" PKT", name_a+" POSS", name_a+" eFG%", name_a+" TS%", name_a+" ORtg", name_a+" PPP",
            name_b+" PKT", name_b+" POSS", name_b+" eFG%", name_b+" TS%", name_b+" ORtg", name_b+" PPP"])
        ws1.row_dimensions[row].height = 22

        metrics = [
            ("Łącznie", suma_a, suma_b, kpi_a, kpi_b),
        ]
        for qn in ["1","2","3","4"]:
            qa = quarters_a.get(qn, {}); qb = quarters_b.get(qn, {})
            metrics.append((f"Kwarta {qn}", qa, qb, calc_kpi(qa), calc_kpi(qb)))

        for i, (lbl, sa, sb, ka, kb) in enumerate(metrics):
            r = row + 1 + i
            ws1.cell(r, 1, lbl).font = Font(bold=True)
            ws1.cell(r, 2, sa.get('pts',0))
            ws1.cell(r, 3, sa.get('poss',0))
            ws1.cell(r, 4, ka.get('efg','-'))
            ws1.cell(r, 5, ka.get('ts','-'))
            ws1.cell(r, 6, ka.get('ortg','-'))
            ws1.cell(r, 7, ka.get('ppp','-'))
            ws1.cell(r, 8, sb.get('pts',0))
            ws1.cell(r, 9, sb.get('poss',0))
            ws1.cell(r,10, kb.get('efg','-'))
            ws1.cell(r,11, kb.get('ts','-'))
            ws1.cell(r,12, kb.get('ortg','-'))
            ws1.cell(r,13, kb.get('ppp','-'))
            style_row(ws1, r, 1, 13, i%2==1)
            # Podświetl wygraną
            try:
                if int(sa.get('pts',0)) > int(sb.get('pts',0)):
                    ws1.cell(r,2).font = GREEN_F
                elif int(sb.get('pts',0)) > int(sa.get('pts',0)):
                    ws1.cell(r,8).font = RED_F
            except: pass

        # ── Arkusz 2: ZAWODNICY ───────────────────────────────────────────────
        for team_name, players in [(name_a, players_a), (name_b, players_b)]:
            if not players: continue
            ws = wb.create_sheet(f"ZAWODNICY {team_name[:8]}")
            ws.column_dimensions['A'].width = 8
            for col in ['B','C','D','E','F','G','H','I','J','K','L','M']:
                ws.column_dimensions[col].width = 10

            ws.merge_cells('A1:M1')
            t = ws['A1']; t.value = f"STATYSTYKI ZAWODNIKÓW — {team_name}"
            t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

            hdrs = ["#","PTS","2PM","2PA","2P%","3PM","3PA","3P%","FTM","FTA","FT%","eFG%","TS%","AST","OREB","DREB","BR","FD","Wykończenia"]
            ws.column_dimensions['A'].width = 6
            for j,h in enumerate(hdrs):
                c = ws.cell(2, j+1, h)
                c.fill = HDR; c.font = HDR_F; c.alignment = CTR
                ws.column_dimensions[get_column_letter(j+1)].width = 9

            for i, (pid, pd) in enumerate(sorted(players.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 99)):
                r = 3 + i
                pts = pd.get('2pm',0)*2 + pd.get('3pm',0)*3 + pd.get('ftm',0)
                pa2 = pd.get('2pa',0); pm2 = pd.get('2pm',0)
                pa3 = pd.get('3pa',0); pm3 = pd.get('3pm',0)
                fta = pd.get('fta',0); ftm = pd.get('ftm',0)
                fga = pa2+pa3
                efg = f"{(pm2+1.5*pm3)/fga:.1%}" if fga else "-"
                ts  = f"{pts/(2*(fga+0.44*fta)):.1%}" if (fga+fta) else "-"
                p2  = f"{pm2/pa2:.1%}" if pa2 else "-"
                p3  = f"{pm3/pa3:.1%}" if pa3 else "-"
                ft  = f"{ftm/fta:.1%}" if fta else "-"
                vals = [pid, pts, pm2, pa2, p2, pm3, pa3, p3, ftm, fta, ft, efg, ts,
                        pd.get('ast',0), pd.get('oreb',0), pd.get('dreb',0),
                        pd.get('br',0), pd.get('fd',0), pd.get('finishes',0)]
                for j, v in enumerate(vals):
                    c = ws.cell(r, j+1, v); c.alignment = CTR; c.border = BORDER
                    if j==1 and pts > 0: c.font = Font(bold=True, color="1A2B4A")
                if i%2==1:
                    for col in range(1, len(vals)+1):
                        ws.cell(r, col).fill = PatternFill("solid", fgColor="F5F8FF")

        # ── Arkusz 3: SHOT TIMING ─────────────────────────────────────────────
        wst = wb.create_sheet("SHOT TIMING")
        wst.column_dimensions['A'].width = 14
        for col in ['B','C','D','E','F','G','H']:
            wst.column_dimensions[col].width = 13

        wst.merge_cells('A1:H1')
        t = wst['A1']; t.value = "SHOT TIMING — czas posiadania a skuteczność"
        t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

        set_hdr(wst, 2, 1, ["Czas", f"{name_a} 2PT", f"{name_a} 3PT", f"{name_a} Eff%",
                             f"{name_b} 2PT", f"{name_b} 3PT", f"{name_b} Eff%", "Różnica Eff%"])

        timing_a = json.loads(session.get('timing_a', '{}')) if session.get('timing_a') else {}
        timing_b = json.loads(session.get('timing_b', '{}')) if session.get('timing_b') else {}

        for i, b in enumerate(BUCKETS):
            r = 3 + i
            wst.cell(r, 1, b).font = Font(bold=True)
            ta = timing_a.get(b, {"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}})
            tb = timing_b.get(b, {"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}})
            m2a=ta["2PT"]["made"]; miss2a=ta["2PT"]["miss"]
            m3a=ta["3PT"]["made"]; miss3a=ta["3PT"]["miss"]
            m2b=tb["2PT"]["made"]; miss2b=tb["2PT"]["miss"]
            m3b=tb["3PT"]["made"]; miss3b=tb["3PT"]["miss"]
            att_a2=m2a+miss2a; att_a3=m3a+miss3a
            att_b2=m2b+miss2b; att_b3=m3b+miss3b
            tot_a=m2a+m3a; tot_att_a=att_a2+att_a3
            tot_b=m2b+m3b; tot_att_b=att_b2+att_b3
            eff_a_val = f"{tot_a/tot_att_a:.0%}" if tot_att_a else "-"
            eff_b_val = f"{tot_b/tot_att_b:.0%}" if tot_att_b else "-"
            try:
                diff = f"{tot_a/tot_att_a - tot_b/tot_att_b:+.0%}" if (tot_att_a and tot_att_b) else "-"
            except: diff = "-"
            wst.cell(r,2, f"{m2a}/{att_a2}")
            wst.cell(r,3, f"{m3a}/{att_a3}")
            wst.cell(r,4, eff_a_val)
            wst.cell(r,5, f"{m2b}/{att_b2}")
            wst.cell(r,6, f"{m3b}/{att_b3}")
            wst.cell(r,7, eff_b_val)
            wst.cell(r,8, diff)
            style_row(wst, r, 1, 8, i%2==1)
            if eff_a_val != "-" and eff_b_val != "-":
                try:
                    if tot_a/tot_att_a > tot_b/tot_att_b:
                        wst.cell(r,4).font = GREEN_F
                    elif tot_b/tot_att_b > tot_a/tot_att_a:
                        wst.cell(r,7).font = RED_F
                except: pass

        # Zapisz do bufora
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        filename = f"raport_{name_a}_vs_{name_b}.xlsx".replace(" ","_")
        return send_file(buf, as_attachment=True,
                         download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        flash(f"Błąd eksportu: {str(e)}", "error")
        return redirect(url_for("index"))


@app.route("/export/pdf")
def export_pdf():
    try:
        name_a    = session.get('name_a', 'Drużyna A')
        name_b    = session.get('name_b', 'Drużyna B')
        suma_a    = json.loads(session.get('suma_a', '{}'))
        suma_b    = json.loads(session.get('suma_b', '{}')) if session.get('suma_b') else {}
        kpi_a     = json.loads(session.get('kpi_a',  '{}'))
        kpi_b     = json.loads(session.get('kpi_b',  '{}')) if session.get('kpi_b') else {}
        quarters_a= json.loads(session.get('quarters_a', '{}'))
        quarters_b= json.loads(session.get('quarters_b', '{}')) if session.get('quarters_b') else {}
        players_a = json.loads(session.get('players_a', '{}'))
        players_b = json.loads(session.get('players_b', '{}')) if session.get('players_b') else {}
        timing_a  = json.loads(session.get('timing_a', '{}')) if session.get('timing_a') else {}
        timing_b  = json.loads(session.get('timing_b', '{}')) if session.get('timing_b') else {}

        score_a = suma_a.get('pts', 0)
        score_b = suma_b.get('pts', 0)
        winner  = name_a if score_a > score_b else (name_b if score_b > score_a else "Remis")

        def row_color(i): return "#f5f8ff" if i%2==0 else "#ffffff"
        def val(d, k): return d.get(k, '-') if d else '-'

        def kpi_row(lbl, va, vb, desc=""):
            try:
                fa = float(str(va).replace('%','').replace('-','0') or 0)
                fb = float(str(vb).replace('%','').replace('-','0') or 0)
                sa = "color:#1a6b3c;font-weight:700" if fa>fb else ""
                sb = "color:#8b1a1a;font-weight:700" if fb>fa else ""
            except: sa=sb=""
            return f"<tr><td style='padding:5px 8px;font-size:11px'><b>{lbl}</b><br><span style='color:#aaa;font-size:10px'>{desc}</span></td><td style='text-align:center;{sa};padding:5px'>{va}</td><td style='text-align:center;{sb};padding:5px'>{vb}</td></tr>"

        # Tabela metryk
        metrics_rows = ""
        for lbl, va, vb, desc in [
            ("Punkty",          score_a,                score_b,                "Łączna liczba punktów"),
            ("Posiadania",      suma_a.get('poss',0),   suma_b.get('poss',0),   "Liczba posiadań"),
            ("eFG%",            val(kpi_a,'efg'),        val(kpi_b,'efg'),        "Efektywny % rzutów z pola"),
            ("True Shooting%",  val(kpi_a,'ts'),         val(kpi_b,'ts'),         "Prawdziwy % skuteczności"),
            ("ORtg",            val(kpi_a,'ortg'),       val(kpi_b,'ortg'),       "Punkty na 100 posiadań"),
            ("PPP",             val(kpi_a,'ppp'),        val(kpi_b,'ppp'),        "Pkt na posiadanie"),
            ("2PT%",            val(kpi_a,'p2_pct'),     val(kpi_b,'p2_pct'),     "Skuteczność za 2 pkt"),
            ("3PT%",            val(kpi_a,'p3_pct'),     val(kpi_b,'p3_pct'),     "Skuteczność za 3 pkt"),
            ("FT%",             val(kpi_a,'ft_pct'),     val(kpi_b,'ft_pct'),     "Skuteczność rzutów wolnych"),
            ("Straty (BR)",     suma_a.get('br',0),      suma_b.get('br',0),      "Liczba strat"),
            ("FT Rate",         val(kpi_a,'ftr'),        val(kpi_b,'ftr'),        "FTA/FGA"),
        ]:
            metrics_rows += kpi_row(lbl, va, vb, desc)

        # Tabela per kwarta
        q_rows = ""
        for qn in ["1","2","3","4","SUMA"]:
            key = qn if qn != "SUMA" else None
            sa = suma_a if key is None else quarters_a.get(qn, {})
            sb = suma_b if key is None else quarters_b.get(qn, {})
            ka = calc_kpi(sa); kb = calc_kpi(sb)
            lbl = f"{qn}Q" if qn != "SUMA" else "SUMA"
            bold = "font-weight:700;background:#e8f0fb" if qn=="SUMA" else ""
            q_rows += f"""<tr style="{bold}">
                <td style='padding:4px 8px;font-weight:600'>{lbl}</td>
                <td style='text-align:center'>{sa.get('pts',0)}</td>
                <td style='text-align:center'>{sa.get('2pm',0)}/{sa.get('2pa',0)}</td>
                <td style='text-align:center'>{ka.get('p2_pct','-')}</td>
                <td style='text-align:center'>{sa.get('3pm',0)}/{sa.get('3pa',0)}</td>
                <td style='text-align:center'>{ka.get('p3_pct','-')}</td>
                <td style='text-align:center'>{sa.get('ftm',0)}/{sa.get('fta',0)}</td>
                <td style='text-align:center'>{sa.get('br',0)}</td>
                <td style='text-align:center'>{sa.get('poss',0)}</td>
                <td style='text-align:center'>{ka.get('efg','-')}</td>
                <td style='text-align:center'>{ka.get('ortg','-')}</td>
                <td style='text-align:center'>{sb.get('pts',0)}</td>
                <td style='text-align:center'>{sb.get('2pm',0)}/{sb.get('2pa',0)}</td>
                <td style='text-align:center'>{kb.get('p2_pct','-')}</td>
                <td style='text-align:center'>{sb.get('3pm',0)}/{sb.get('3pa',0)}</td>
                <td style='text-align:center'>{kb.get('p3_pct','-')}</td>
                <td style='text-align:center'>{sb.get('ftm',0)}/{sb.get('fta',0)}</td>
                <td style='text-align:center'>{sb.get('br',0)}</td>
                <td style='text-align:center'>{sb.get('poss',0)}</td>
                <td style='text-align:center'>{kb.get('efg','-')}</td>
                <td style='text-align:center'>{kb.get('ortg','-')}</td>
            </tr>"""

        # Tabela zawodników
        def player_table(players, team):
            if not players: return "<p style='color:#aaa;font-size:11px'>Brak danych zawodników</p>"
            rows = ""
            for i,(pid,pd) in enumerate(sorted(players.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 99)):
                pts=pd.get('2pm',0)*2+pd.get('3pm',0)*3+pd.get('ftm',0)
                pa2=pd.get('2pa',0); pm2=pd.get('2pm',0)
                pa3=pd.get('3pa',0); pm3=pd.get('3pm',0)
                fga=pa2+pa3; fta=pd.get('fta',0); ftm=pd.get('ftm',0)
                efg=f"{(pm2+1.5*pm3)/fga:.0%}" if fga else "-"
                ts =f"{pts/(2*(fga+0.44*fta)):.0%}" if (fga+fta) else "-"
                bg = row_color(i)
                rows += f"""<tr style='background:{bg};font-size:10px'>
                    <td style='padding:3px 6px;font-weight:700'>{pid}</td>
                    <td style='text-align:center;font-weight:700;color:#1a2b4a'>{pts}</td>
                    <td style='text-align:center'>{pm2}/{pa2}</td>
                    <td style='text-align:center'>{pm3}/{pa3}</td>
                    <td style='text-align:center'>{ftm}/{fta}</td>
                    <td style='text-align:center;font-weight:600'>{efg}</td>
                    <td style='text-align:center'>{ts}</td>
                    <td style='text-align:center'>{pd.get('ast',0)}</td>
                    <td style='text-align:center'>{pd.get('oreb',0)}</td>
                    <td style='text-align:center'>{pd.get('dreb',0)}</td>
                    <td style='text-align:center'>{pd.get('br',0)}</td>
                    <td style='text-align:center'>{pd.get('finishes',0)}</td>
                </tr>"""
            return f"""<table style='width:100%;border-collapse:collapse;font-size:10px'>
                <thead><tr style='background:#1a2b4a;color:#fff'>
                    <th style='padding:4px 6px'>#</th><th>PTS</th><th>2PM/A</th><th>3PM/A</th>
                    <th>FTM/A</th><th>eFG%</th><th>TS%</th><th>AST</th>
                    <th>ORB</th><th>DRB</th><th>BR</th><th>FIN</th>
                </tr></thead><tbody>{rows}</tbody></table>"""

        # Shot timing tabela
        tim_rows = ""
        for i,b in enumerate(BUCKETS):
            ta=timing_a.get(b,{"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}})
            tb=timing_b.get(b,{"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}})
            m_a=ta["2PT"]["made"]+ta["3PT"]["made"]; att_a=m_a+ta["2PT"]["miss"]+ta["3PT"]["miss"]
            m_b=tb["2PT"]["made"]+tb["3PT"]["made"]; att_b=m_b+tb["2PT"]["miss"]+tb["3PT"]["miss"]
            ea=f"{m_a/att_a:.0%}" if att_a else "-"
            eb=f"{m_b/att_b:.0%}" if att_b else "-"
            bg=row_color(i)
            tim_rows += f"""<tr style='background:{bg};font-size:10px'>
                <td style='padding:3px 8px;font-weight:700'>{b}</td>
                <td style='text-align:center'>{m_a}/{att_a}</td>
                <td style='text-align:center;font-weight:600;color:#1a6b3c'>{ea}</td>
                <td style='text-align:center'>{m_b}/{att_b}</td>
                <td style='text-align:center;font-weight:600;color:#8b1a1a'>{eb}</td>
            </tr>"""

        TH = "style='background:#1a2b4a;color:#fff;padding:4px 6px;font-size:10px;text-align:center'"
        TH_L = "style='background:#1a2b4a;color:#fff;padding:4px 6px;font-size:10px'"

        html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  body{{font-family:Arial,sans-serif;margin:0;padding:20px;color:#222;font-size:11px}}
  h1{{font-size:18px;margin:0}} h2{{font-size:13px;color:#1a2b4a;margin:16px 0 6px}}
  h3{{font-size:11px;color:#1a2b4a;margin:10px 0 4px;text-transform:uppercase;letter-spacing:.5px}}
  .hero{{background:#1a2b4a;color:#fff;padding:14px 20px;border-radius:8px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}}
  .score{{font-size:28px;font-weight:700;letter-spacing:4px}}
  table{{width:100%;border-collapse:collapse;margin-bottom:12px}}
  th{{background:#1a2b4a;color:#fff;padding:4px 6px;font-size:10px}}
  td{{padding:3px 6px;border-bottom:1px solid #eee;font-size:10px}}
  .section{{margin-bottom:20px;page-break-inside:avoid}}
  .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
  @media print{{body{{padding:8px}} .no-print{{display:none}}}}
  @page{{size:A4;margin:1.5cm}}
</style>
</head><body>

<div class="hero">
  <div>
    <h1>{name_a} vs {name_b}</h1>
    <div style='opacity:.7;font-size:11px'>Raport meczowy · Basket Kołcz Analytics</div>
  </div>
  <div class="score">{score_a} : {score_b}</div>
  <div style='font-size:12px;opacity:.8'>{winner}<br>wygrywa</div>
</div>

<div class="section">
  <h2>Kluczowe metryki</h2>
  <table>
    <thead><tr>
      <th style='text-align:left'>Metryka</th>
      <th style='color:#7dffb3'>{name_a}</th>
      <th style='color:#ffaaaa'>{name_b}</th>
    </tr></thead>
    <tbody>{metrics_rows}</tbody>
  </table>
</div>

<div class="section">
  <h2>Statystyki per kwarta</h2>
  <p style='font-size:9px;color:#aaa'>PKT · 2PM/A · 2P% · 3PM/A · 3P% · FTM/A · BR · POSS · eFG% · ORtg</p>
  <table>
    <thead>
      <tr>
        <th {TH_L}>Q</th>
        <th colspan="10" style='background:#1a6b3c;color:#fff;padding:4px;font-size:10px'>{name_a}</th>
        <th colspan="10" style='background:#8b1a1a;color:#fff;padding:4px;font-size:10px'>{name_b}</th>
      </tr>
      <tr>
        <th {TH_L}>Q</th>
        {''.join(f'<th {TH}>{h}</th>' for h in ['PKT','2PM/A','2P%','3PM/A','3P%','FTM/A','BR','POSS','eFG%','ORtg'])}
        {''.join(f'<th {TH}>{h}</th>' for h in ['PKT','2PM/A','2P%','3PM/A','3P%','FTM/A','BR','POSS','eFG%','ORtg'])}
      </tr>
    </thead>
    <tbody>{q_rows}</tbody>
  </table>
</div>

<div class="two-col">
  <div class="section">
    <h2 style='color:#1a6b3c'>{name_a} — Zawodnicy</h2>
    {player_table(players_a, name_a)}
  </div>
  <div class="section">
    <h2 style='color:#8b1a1a'>{name_b} — Zawodnicy</h2>
    {player_table(players_b, name_b)}
  </div>
</div>

<div class="section">
  <h2>Shot Timing</h2>
  <table style='width:50%'>
    <thead><tr>
      <th style='background:#1a2b4a;color:#fff;padding:4px 8px;text-align:left'>Czas</th>
      <th style='background:#1a6b3c;color:#fff;padding:4px;text-align:center'>{name_a} Celne/Att</th>
      <th style='background:#1a6b3c;color:#fff;padding:4px;text-align:center'>{name_a} Eff%</th>
      <th style='background:#8b1a1a;color:#fff;padding:4px;text-align:center'>{name_b} Celne/Att</th>
      <th style='background:#8b1a1a;color:#fff;padding:4px;text-align:center'>{name_b} Eff%</th>
    </tr></thead>
    <tbody>{tim_rows}</tbody>
  </table>
</div>

<div style='margin-top:20px;text-align:center;font-size:9px;color:#aaa;border-top:1px solid #eee;padding-top:8px'>
  Basket Kołcz Analytics · Wygenerowano automatycznie
</div>

</body></html>"""

        buf = io.BytesIO(html.encode('utf-8'))
        buf.seek(0)
        filename = f"raport_{name_a}_vs_{name_b}.html".replace(" ","_")
        return send_file(buf, as_attachment=True,
                         download_name=filename,
                         mimetype='text/html')

    except Exception as e:
        flash(f"Błąd eksportu PDF: {str(e)}", "error")
        return redirect(url_for("index"))


@app.route("/template/zapis")
def template_zapis():
    wb = openpyxl.Workbook()

    # Styl nagłówka
    HDR  = PatternFill("solid", fgColor="1A2B4A")
    HDR_F = Font(color="FFFFFF", bold=True, size=10)
    YEL  = PatternFill("solid", fgColor="FFF9C4")
    GRN  = PatternFill("solid", fgColor="E8F5E9")
    CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    BORDER = Border(
        bottom=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        left=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
    )

    KODY = [
        ("2",      "Celny rzut za 2 pkt"),
        ("0/2",    "Niecelny rzut za 2 pkt"),
        ("3",      "Celny rzut za 3 pkt"),
        ("0/3",    "Niecelny rzut za 3 pkt"),
        ("BR",     "Strata piłki"),
        ("P",      "Przewinienie"),
        ("F",      "Faul"),
        ("2+1",    "Celny 2pkt + rzut wolny"),
        ("2+0",    "Celny 2pkt + niecelny RW"),
        ("3+1",    "Celny 3pkt + rzut wolny"),
        ("3+0",    "Celny 3pkt + niecelny RW"),
        ("2D",     "Celny tip-in za 2pkt"),
        ("0/2D",   "Niecelny tip-in za 2pkt"),
        ("2D+1",   "Celny tip-in + RW"),
        ("1/2W",   "1/2 rzutów wolnych"),
        ("2/2W",   "2/2 rzutów wolnych"),
        ("0/2W",   "0/2 rzutów wolnych"),
        ("1/3W",   "1/3 rzutów wolnych"),
        ("2/3W",   "2/3 rzutów wolnych"),
        ("3/3W",   "3/3 rzutów wolnych"),
        ("0/3W",   "0/3 rzutów wolnych"),
    ]

    # ── Arkusz KODY (ściągawka) ────────────────────────────────────────────
    ws_kody = wb.active
    ws_kody.title = "KODY"
    ws_kody.column_dimensions['A'].width = 12
    ws_kody.column_dimensions['B'].width = 32

    ws_kody.merge_cells('A1:B1')
    t = ws_kody['A1']
    t.value = "KODY AKCJI — ściągawka"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR
    ws_kody.row_dimensions[1].height = 24

    ws_kody['A2'] = "KOD";      ws_kody['A2'].fill = HDR; ws_kody['A2'].font = HDR_F; ws_kody['A2'].alignment = CTR
    ws_kody['B2'] = "OPIS";     ws_kody['B2'].fill = HDR; ws_kody['B2'].font = HDR_F; ws_kody['B2'].alignment = CTR

    for i,(kod,opis) in enumerate(KODY):
        r = 3+i
        c1 = ws_kody.cell(r,1,kod);  c1.fill = YEL; c1.alignment = CTR; c1.border = BORDER
        c1.font = Font(bold=True, size=10)
        c2 = ws_kody.cell(r,2,opis); c2.border = BORDER
        if i%2==0: c2.fill = PatternFill("solid", fgColor="FAFAFA")

    ws_kody.merge_cells('A26:B26')
    ws_kody['A26'] = "Wielokrotne akcje w posiadaniu: oddziel średnikiem  np. F;2+1"
    ws_kody['A26'].font = Font(italic=True, color="555555", size=9)

    # ── Arkusze drużyn ─────────────────────────────────────────────────────
    COLS = [
        ("A","Kwarta\n(*kto rozpoczął)",7),
        ("B","Czas trwania\nakcji",9),
        ("C","Kod zakończenia\nakcji",12),
        ("D","Strefa\n(numer)",8),
        ("E","Zawodnik 1",11),
        ("F","Zawodnik 2",11),
        ("G","Zawodnik 3",11),
        ("H","Zawodnik 4",11),
        ("I","Zawodnik 5",11),
        ("J","Timeout",8),
        ("K","Zawodnik\nkończący",11),
        ("L","Asysta ★",9),
        ("M","OREB ★",9),
        ("N","DREB ★",9),
    ]

    for team_name in ["drużyna_A", "drużyna_B"]:
        ws = wb.create_sheet(team_name)

        # Nagłówek
        ws.merge_cells('A1:N1')
        t = ws['A1']
        t.value = f"ZAPIS MECZU — {team_name.upper()}   |   Wpisz nazwę drużyny w zakładce"
        t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=11); t.alignment = CTR
        ws.row_dimensions[1].height = 22

        # Kolumny
        for col_letter, label, width in COLS:
            ws.column_dimensions[col_letter].width = width
            c = ws[f"{col_letter}2"]
            c.value = label; c.fill = HDR; c.font = HDR_F
            c.alignment = CTR; c.border = BORDER
        ws.row_dimensions[2].height = 32

        # Walidacja kolumny C (kody akcji)
        from openpyxl.worksheet.datavalidation import DataValidation
        valid_codes = ",".join([k for k,_ in KODY])
        # Tylko podstawowe — Excel ma limit długości
        dv = DataValidation(
            type="list",
            formula1='"2,0/2,3,0/3,BR,P,F,2+1,2D,0/2D,1/2W,2/2W,0/2W"',
            allow_blank=True,
            showErrorMessage=False
        )
        ws.add_data_validation(dv)
        dv.add(f"C3:C500")

        # Puste wiersze z formatowaniem (50 wierszy)
        for r in range(3, 53):
            for col_idx, (col_letter, _, _) in enumerate(COLS):
                c = ws[f"{col_letter}{r}"]
                c.border = BORDER
                c.alignment = Alignment(horizontal="center", vertical="center")
                # Żółte tło dla kluczowych kolumn
                if col_letter in ['A','B','C','D']:
                    c.fill = YEL
                elif col_letter in ['K','L','M','N']:
                    c.fill = GRN
            if r % 2 == 0:
                for col_letter, _, _ in COLS:
                    existing = ws[f"{col_letter}{r}"].fill
                    if ws[f"{col_letter}{r}"].fill.fgColor.rgb in ('FFFFFFFF', '00000000'):
                        ws[f"{col_letter}{r}"].fill = PatternFill("solid", fgColor="F8F8F8")

        # Legenda kolorów
        r_leg = 55
        ws.merge_cells(f'A{r_leg}:N{r_leg}')
        ws[f'A{r_leg}'] = "🟡 Żółte = obowiązkowe   🟢 Zielone = opcjonalne (★ = tylko nowy format 14 kolumn)"
        ws[f'A{r_leg}'].font = Font(italic=True, color="666666", size=9)

        # Freeze row
        ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="ZAPIS_MECZU_szablon.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route("/template/szablon")
def template_szablon():
    wb = openpyxl.Workbook()

    HDR   = PatternFill("solid", fgColor="1A2B4A")
    HDR_F = Font(color="FFFFFF", bold=True, size=10)
    AMB   = PatternFill("solid", fgColor="FFF8E1")   # INPUT
    KPI   = PatternFill("solid", fgColor="E8F5E9")   # KPI
    CTR   = Alignment(horizontal="center", vertical="center")
    BORDER = Border(
        bottom=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        left=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
    )

    def hdr(ws, row, col, val, w=10):
        c = ws.cell(row, col, val)
        c.fill = HDR; c.font = HDR_F; c.alignment = CTR; c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = w

    def inp(ws, row, col, val=""):
        c = ws.cell(row, col, val)
        c.fill = AMB; c.alignment = CTR; c.border = BORDER

    def kpi_cell(ws, row, col, val=""):
        c = ws.cell(row, col, val)
        c.fill = KPI; c.alignment = CTR; c.border = BORDER

    def section_title(ws, row, col_start, col_end, title):
        ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
        c = ws.cell(row, col_start, title)
        c.fill = PatternFill("solid", fgColor="E3F2FD")
        c.font = Font(bold=True, color="0C447C", size=10)
        c.alignment = CTR

    # ── TEAM GENERAL ──────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "TEAM GENERAL"
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells('A1:V1')
    t = ws1['A1']; t.value = "TEAM GENERAL — Statystyki drużyny per kwarta"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    hdrs_gen = [
        ("Q",4),("PKT",7),("POSS",6),
        ("2PM",6),("2PA",6),("2P%",7),
        ("3PM",6),("3PA",6),("3P%",7),
        ("FTM",6),("FTA",6),("FT%",7),
        ("BR",6),("P",6),("FD",6),
        ("eFG%",7),("TS%",7),("ORtg",7),
        ("DRtg",7),("NetRtg",8),("PPP",7),("TO%",7),
    ]
    for i,(h,w) in enumerate(hdrs_gen):
        hdr(ws1, 2, i+1, h, w)

    for r, qname in enumerate(["1Q","2Q","3Q","4Q","SUMA"], 3):
        ws1.cell(r, 1, qname).font = Font(bold=True)
        for col in range(2, len(hdrs_gen)+1):
            inp(ws1, r, col)
        if qname == "SUMA":
            ws1.row_dimensions[r].height = 20
            for col in range(1, len(hdrs_gen)+1):
                ws1.cell(r, col).fill = PatternFill("solid", fgColor="E8F0FB")
                ws1.cell(r, col).font = Font(bold=True)

    ws1.freeze_panes = "B3"

    # ── PLAYERS ───────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("PLAYERS")
    ws2.merge_cells('A1:P1')
    t = ws2['A1']; t.value = "PLAYERS — Statystyki zawodników"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    hdrs_p = [
        ("#",5),("MIN",7),("PTS",7),
        ("2PM",6),("2PA",6),("2P%",7),
        ("3PM",6),("3PA",6),("3P%",7),
        ("FTM",6),("FTA",6),("FT%",7),
        ("eFG%",7),("TS%",7),
        ("AST",6),("OREB",6),("DREB",6),
        ("BR",6),("FD",6),("FIN",6),
    ]
    for i,(h,w) in enumerate(hdrs_p):
        hdr(ws2, 2, i+1, h, w)

    for r in range(3, 18):
        for col in range(1, len(hdrs_p)+1):
            inp(ws2, r, col)
        if r % 2 == 0:
            for col in range(1, len(hdrs_p)+1):
                ws2.cell(r, col).fill = PatternFill("solid", fgColor="FFF3E0")

    ws2.freeze_panes = "B3"

    # ── LINEUPS ───────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("LINEUPS")
    ws3.merge_cells('A1:L1')
    t = ws3['A1']; t.value = "LINEUPS — Statystyki składów 5-osobowych"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    hdrs_l = [
        ("Skład",22),("POSS",7),("PKT",7),("PPP",7),
        ("eFG%",7),("ORtg",7),("DRtg",7),("NetRtg",9),
        ("BR",6),("FD",6),("2P%",7),("3P%",7),("Śr.tempo",10),
    ]
    for i,(h,w) in enumerate(hdrs_l):
        hdr(ws3, 2, i+1, h, w)

    for r in range(3, 18):
        for col in range(1, len(hdrs_l)+1):
            inp(ws3, r, col)

    ws3.freeze_panes = "B3"

    # ── SHOT TIMING ───────────────────────────────────────────────────────
    ws4 = wb.create_sheet("SHOT TIMING")
    ws4.merge_cells('A1:H1')
    t = ws4['A1']; t.value = "SHOT TIMING — Skuteczność vs czas posiadania"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    hdrs_t = [
        ("Czas posiadania",18),("2PT Made",10),("2PT Att",10),("2PT%",8),
        ("3PT Made",10),("3PT Att",10),("3PT%",8),("Eff% łącznie",12),
    ]
    for i,(h,w) in enumerate(hdrs_t):
        hdr(ws4, 2, i+1, h, w)

    for i,b in enumerate(["0s","1-4s","5-8s","9-12s","13-16s","17-20s","21-24s"]):
        r = 3+i
        ws4.cell(r, 1, b).font = Font(bold=True)
        for col in range(2, len(hdrs_t)+1):
            inp(ws4, r, col)

    ws4.freeze_panes = "B3"

    # ── COURT ZONES ───────────────────────────────────────────────────────
    ws5 = wb.create_sheet("COURT ZONES")
    ws5.merge_cells('A1:H1')
    t = ws5['A1']; t.value = "COURT ZONES — Statystyki per strefa boiska"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    zone_names_list = [
        (1,"Pod koszem"),(2,"Prawy blok bliski"),(3,"Lewy blok bliski"),
        (4,"Środek farby"),(5,"Lewy blok daleki"),(6,"Prawy blok daleki"),
        (7,"Lewy baseline"),(8,"Prawy baseline"),(9,"Lewy corner 3PT"),
        (10,"Prawy corner 3PT"),(11,"Lewe skrzydło 3PT"),(12,"Góra łuku 3PT"),
        (13,"Prawe skrzydło 3PT"),
    ]

    hdrs_z = [
        ("#",5),("Strefa",22),("Obszar",12),
        ("Celne",8),("Niecelne",9),("Próby",7),("Eff%",8),("PPP",7),
    ]
    for i,(h,w) in enumerate(hdrs_z):
        hdr(ws5, 2, i+1, h, w)

    for i,(zn,zname) in enumerate(zone_names_list):
        r = 3+i
        area = "Pod koszem" if zn<=2 else ("Farba" if zn<=6 else ("Mid-range" if zn<=8 else "3PT"))
        ws5.cell(r,1,zn).font = Font(bold=True); ws5.cell(r,1).alignment = CTR
        ws5.cell(r,2,zname)
        ws5.cell(r,3,area)
        for col in range(4, len(hdrs_z)+1):
            inp(ws5, r, col)
        if i%2==0:
            ws5.cell(r,2).fill = PatternFill("solid", fgColor="FAFAFA")
            ws5.cell(r,3).fill = PatternFill("solid", fgColor="FAFAFA")

    ws5.freeze_panes = "D3"

    # ── LEGENDA ───────────────────────────────────────────────────────────
    ws6 = wb.create_sheet("LEGENDA")
    ws6.column_dimensions['A'].width = 22
    ws6.column_dimensions['B'].width = 48

    ws6.merge_cells('A1:B1')
    t = ws6['A1']; t.value = "LEGENDA — Opis metryk"
    t.fill = HDR; t.font = Font(color="FFFFFF", bold=True, size=12); t.alignment = CTR

    legend = [
        ("eFG%",       "Efektywny % rzutów: (2PM + 1.5×3PM) / FGA"),
        ("TS%",        "Prawdziwy %: PTS / (2 × (FGA + 0.44×FTA))"),
        ("ORtg",       "Offensive Rating: PTS×100/POSS"),
        ("DRtg",       "Defensive Rating: OPP_PTS×100/OPP_POSS"),
        ("NetRtg",     "ORtg − DRtg"),
        ("PPP",        "Punkty na posiadanie: PTS/POSS"),
        ("TO%",        "% posiadań ze stratą: BR/POSS"),
        ("FT Rate",    "Stosunek rzutów wolnych: FTA/FGA"),
        ("FIN",        "Wykończenia — rzuty kończące posiadanie"),
        ("POSS",       "Posiadania — liczba ataków drużyny"),
        ("BR",         "Ball lost — strata piłki"),
        ("FD",         "Faule wymuszone na atakującym"),
    ]
    hdr(ws6, 2, 1, "Metryka", 22); hdr(ws6, 2, 2, "Opis", 48)
    for i,(m,o) in enumerate(legend):
        r = 3+i
        c1=ws6.cell(r,1,m); c1.font=Font(bold=True); c1.border=BORDER
        c2=ws6.cell(r,2,o); c2.border=BORDER
        if i%2==0:
            c1.fill=PatternFill("solid",fgColor="F5F5F5")
            c2.fill=PatternFill("solid",fgColor="F5F5F5")

    # Legenda kolorów
    r_leg = 3+len(legend)+2
    ws6.merge_cells(f'A{r_leg}:B{r_leg}')
    ws6[f'A{r_leg}'] = "Kolory komórek:"
    ws6[f'A{r_leg}'].font = Font(bold=True)
    for i, (col, desc) in enumerate([
        ("FFF8E1","🟡 Żółty — pole INPUT (wpisz dane)"),
        ("E8F5E9","🟢 Zielony — pole KPI (obliczone)"),
    ]):
        r2 = r_leg+1+i
        ws6.cell(r2,1).fill = PatternFill("solid",fgColor=col)
        ws6.cell(r2,1).border = BORDER
        ws6.cell(r2,2,desc).border = BORDER

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="SZABLON_MECZ.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == "__main__":
    app.run(debug=True)
