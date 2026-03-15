from flask import Flask, render_template_string, request, redirect, url_for, flash
import openpyxl
from collections import defaultdict
import re

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
  <div class="card mt-3 p-3" style="font-size:.82rem;color:#666">
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

if __name__ == "__main__":
    app.run(debug=True)
