from flask import Flask, render_template_string, request, redirect, url_for, flash, session, send_file, g
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import re, io, json, os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "basketkolcz2025secret")

DATABASE_URL = os.environ.get("DATABASE_URL", "")

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════════════════════════════

def get_db():
    if "db" not in g:
        g.db = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def init_db():
    db = get_db()
    cur = db.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS matches (
        id SERIAL PRIMARY KEY,
        sezon VARCHAR(20) NOT NULL DEFAULT '2024/25',
        data_meczu DATE,
        przeciwnik VARCHAR(100) NOT NULL,
        wynik_gtk INTEGER DEFAULT 0,
        wynik_opp INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS match_stats (
        id SERIAL PRIMARY KEY,
        match_id INTEGER REFERENCES matches(id) ON DELETE CASCADE,
        druzyna VARCHAR(10) NOT NULL,
        kwarta INTEGER NOT NULL,
        pts INTEGER DEFAULT 0,
        poss INTEGER DEFAULT 0,
        p2m INTEGER DEFAULT 0,
        p2a INTEGER DEFAULT 0,
        p3m INTEGER DEFAULT 0,
        p3a INTEGER DEFAULT 0,
        ftm INTEGER DEFAULT 0,
        fta INTEGER DEFAULT 0,
        br INTEGER DEFAULT 0,
        fd INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS player_stats (
        id SERIAL PRIMARY KEY,
        match_id INTEGER REFERENCES matches(id) ON DELETE CASCADE,
        druzyna VARCHAR(10) NOT NULL,
        nr INTEGER NOT NULL,
        pts INTEGER DEFAULT 0,
        p2m INTEGER DEFAULT 0,
        p2a INTEGER DEFAULT 0,
        p3m INTEGER DEFAULT 0,
        p3a INTEGER DEFAULT 0,
        ftm INTEGER DEFAULT 0,
        fta INTEGER DEFAULT 0,
        ast INTEGER DEFAULT 0,
        oreb INTEGER DEFAULT 0,
        dreb INTEGER DEFAULT 0,
        br INTEGER DEFAULT 0,
        fd INTEGER DEFAULT 0,
        finishes INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS timing_stats (
        id SERIAL PRIMARY KEY,
        match_id INTEGER REFERENCES matches(id) ON DELETE CASCADE,
        druzyna VARCHAR(10) NOT NULL,
        bucket VARCHAR(10) NOT NULL,
        made2 INTEGER DEFAULT 0,
        att2 INTEGER DEFAULT 0,
        made3 INTEGER DEFAULT 0,
        att3 INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS settings (
        key VARCHAR(50) PRIMARY KEY,
        value VARCHAR(200)
    );

    INSERT INTO settings (key, value) VALUES ('gtk_name', 'GTK') ON CONFLICT DO NOTHING;
    INSERT INTO settings (key, value) VALUES ('current_season', '2024/25') ON CONFLICT DO NOTHING;
    """)
    db.commit()
    cur.close()

def get_setting(key):
    db = get_db()
    cur = db.cursor()
    cur.execute("SELECT value FROM settings WHERE key=%s", (key,))
    row = cur.fetchone()
    cur.close()
    return row['value'] if row else None

def set_setting(key, value):
    db = get_db()
    cur = db.cursor()
    cur.execute("INSERT INTO settings (key,value) VALUES (%s,%s) ON CONFLICT (key) DO UPDATE SET value=%s",
                (key, value, value))
    db.commit()
    cur.close()

# ══════════════════════════════════════════════════════════════════════════════
# PARSER (identyczny jak w app_v2.py)
# ══════════════════════════════════════════════════════════════════════════════

ACTION_2PM = {"2","2+1","2+0","2D","2D+1"}
ACTION_3PM = {"3","3+1","3+0"}
ACTION_BR  = {"BR"}
ACTION_F   = {"F"}
BUCKETS    = ["0s","1-4s","5-8s","9-12s","13-16s","17-20s","21-24s"]

def extract_ft(code):
    m = re.match(r'^(\d+)/(\d+)W', code)
    if m: return int(m.group(1)), int(m.group(2))
    m2 = re.search(r'(\d+)/(\d+)W', code)
    if m2: return int(m2.group(1)), int(m2.group(2))
    return 0, 0

def time_bucket(t):
    if t == 0:    return "0s"
    if t <= 4:    return "1-4s"
    if t <= 8:    return "5-8s"
    if t <= 12:   return "9-12s"
    if t <= 16:   return "13-16s"
    if t <= 20:   return "17-20s"
    return "21-24s"

def parse_team_sheet(ws):
    stats = {
        "quarter": defaultdict(lambda: {"ftm":0,"fta":0,"p2m":0,"p2a":0,"p3m":0,"p3a":0,"br":0,"fd":0,"poss":0,"pts":0}),
        "players": defaultdict(lambda: {"p2m":0,"p2a":0,"p3m":0,"p3a":0,"ftm":0,"fta":0,"fd":0,"br":0,"finishes":0,"ast":0,"oreb":0,"dreb":0}),
        "timing":  {b: {"2PT":{"made":0,"miss":0},"3PT":{"made":0,"miss":0}} for b in BUCKETS},
    }
    current_q = 1
    current_lineup = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v is not None for v in row[:4]): continue
        if row[0] is not None:
            try: current_q = int(str(row[0]).replace("*","").strip())
            except: current_q = 1

        for i in range(4, 9):
            val = row[i] if len(row) > i and row[i] is not None else None
            if val is not None:
                s = str(val).strip()
                if ";" in s:
                    parts = s.split(";")
                    try:
                        old_p, new_p = int(parts[0].strip()), int(parts[1].strip())
                        try: current_lineup[current_lineup.index(old_p)] = new_p
                        except ValueError: current_lineup.append(new_p)
                    except: pass
                else:
                    try:
                        p = int(s)
                        if p not in current_lineup: current_lineup.append(p)
                    except: pass
        if len(current_lineup) > 5: current_lineup = current_lineup[-5:]

        raw_b = str(row[1]) if row[1] is not None else ""
        raw_c = str(row[2]) if row[2] is not None else ""
        raw_d = str(row[3]) if row[3] is not None else ""
        raw_k = str(row[10]) if len(row)>10 and row[10] is not None else ""
        raw_l = str(row[11]) if len(row)>11 and row[11] is not None else ""
        raw_m = str(row[12]) if len(row)>12 and row[12] is not None else ""
        raw_n = str(row[13]) if len(row)>13 and row[13] is not None else ""

        codes     = [c.strip() for c in raw_c.split(";") if c.strip()]
        times     = [t.strip() for t in raw_b.replace(";",",").split(",") if t.strip()]
        finishers = [f.strip() for f in raw_k.split(";") if f.strip()]
        assists   = [a.strip() for a in raw_l.split(";") if a.strip()]
        orebs     = [o.strip() for o in raw_m.split(";") if o.strip()]
        drebs     = [d.strip() for d in raw_n.split(";") if d.strip()]

        q = stats["quarter"][current_q]
        q["poss"] += 1

        for ai, code in enumerate(codes):
            t_val = 0
            if ai < len(times):
                try: t_val = float(times[ai])
                except: pass
            bucket = time_bucket(t_val)

            finisher = None
            if ai < len(finishers):
                try: finisher = int(finishers[ai])
                except: pass

            assister = None
            if ai < len(assists):
                try: assister = int(assists[ai])
                except: pass

            orebler = None
            if ai < len(orebs):
                try: orebler = int(orebs[ai])
                except: pass

            drebler = None
            if ai < len(drebs):
                try: drebler = int(drebs[ai])
                except: pass

            pts = 0

            if code in ACTION_2PM:
                q["p2m"]+=1; q["p2a"]+=1; pts=2
                stats["timing"][bucket]["2PT"]["made"]+=1
                if finisher:
                    stats["players"][finisher]["p2m"]+=1; stats["players"][finisher]["p2a"]+=1
                    stats["players"][finisher]["finishes"]+=1
                    if assister: stats["players"][assister]["ast"]+=1
            elif code in ("0/2","0/2D"):
                q["p2a"]+=1
                stats["timing"][bucket]["2PT"]["miss"]+=1
                if finisher: stats["players"][finisher]["p2a"]+=1; stats["players"][finisher]["finishes"]+=1
                if orebler: stats["players"][orebler]["oreb"]+=1
            elif code in ACTION_3PM:
                q["p3m"]+=1; q["p3a"]+=1; pts=3
                stats["timing"][bucket]["3PT"]["made"]+=1
                if finisher:
                    stats["players"][finisher]["p3m"]+=1; stats["players"][finisher]["p3a"]+=1
                    stats["players"][finisher]["finishes"]+=1
                    if assister: stats["players"][assister]["ast"]+=1
            elif code == "0/3":
                q["p3a"]+=1
                stats["timing"][bucket]["3PT"]["miss"]+=1
                if finisher: stats["players"][finisher]["p3a"]+=1; stats["players"][finisher]["finishes"]+=1
                if orebler: stats["players"][orebler]["oreb"]+=1
            elif code in ACTION_BR:
                q["br"]+=1
                if finisher: stats["players"][finisher]["br"]+=1
                if drebler: stats["players"][drebler]["dreb"]+=1
            elif code in ACTION_F:
                q["fd"]+=1

            ftm, fta = extract_ft(code)
            if fta > 0:
                q["ftm"]+=ftm; q["fta"]+=fta; pts+=ftm
                if finisher:
                    stats["players"][finisher]["ftm"]+=ftm; stats["players"][finisher]["fta"]+=fta
                    stats["players"][finisher]["fd"]+=1

            q["pts"] += pts

    return stats

def suma_quarters(stats):
    s = defaultdict(int)
    for qn in [1,2,3,4]:
        for k,v in stats["quarter"].get(qn,{}).items():
            s[k] += v
    return dict(s)

def calc_kpi(d):
    fga = d.get("p2a",0) + d.get("p3a",0)
    pts = d.get("pts",0)
    poss = max(d.get("poss",1),1)
    fta = d.get("fta",0); ftm = d.get("ftm",0)
    pm2 = d.get("p2m",0); pa2 = d.get("p2a",0)
    pm3 = d.get("p3m",0); pa3 = d.get("p3a",0)
    def pct(n,d): return f"{n/d:.1%}" if d else "-"
    efg  = (pm2+1.5*pm3)/fga if fga else None
    ts   = pts/(2*(fga+0.44*fta)) if (fga+fta) else None
    return {
        "efg":   pct(pm2+1.5*pm3,fga) if fga else "-",
        "ts":    f"{ts:.1%}" if ts else "-",
        "ortg":  f"{pts*100/poss:.1f}",
        "ppp":   f"{pts/poss:.2f}",
        "topct": f"{d.get('br',0)/poss:.1%}",
        "ftr":   f"{fta/fga:.2f}" if fga else "-",
        "p2_pct":pct(pm2,pa2),
        "p3_pct":pct(pm3,pa3),
        "ft_pct":pct(ftm,fta),
    }

def save_match_to_db(przeciwnik, sezon, data_meczu, stats_gtk, stats_opp):
    db = get_db()
    cur = db.cursor()
    suma_gtk = suma_quarters(stats_gtk)
    suma_opp = suma_quarters(stats_opp)

    # Wstaw mecz
    cur.execute("""
        INSERT INTO matches (sezon, data_meczu, przeciwnik, wynik_gtk, wynik_opp)
        VALUES (%s, %s, %s, %s, %s) RETURNING id
    """, (sezon, data_meczu, przeciwnik, suma_gtk.get("pts",0), suma_opp.get("pts",0)))
    match_id = cur.fetchone()["id"]

    # Statystyki per kwarta
    for qn in [1,2,3,4]:
        qg = stats_gtk["quarter"].get(qn,{})
        qo = stats_opp["quarter"].get(qn,{})
        for druzyna, qd in [("gtk", qg), ("opp", qo)]:
            cur.execute("""
                INSERT INTO match_stats (match_id,druzyna,kwarta,pts,poss,p2m,p2a,p3m,p3a,ftm,fta,br,fd)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (match_id, druzyna, qn,
                  qd.get("pts",0), qd.get("poss",0),
                  qd.get("p2m",0), qd.get("p2a",0),
                  qd.get("p3m",0), qd.get("p3a",0),
                  qd.get("ftm",0), qd.get("fta",0),
                  qd.get("br",0),  qd.get("fd",0)))

    # Zawodnicy
    for druzyna, stats in [("gtk", stats_gtk), ("opp", stats_opp)]:
        for nr, pd in stats["players"].items():
            pts = pd.get("p2m",0)*2 + pd.get("p3m",0)*3 + pd.get("ftm",0)
            cur.execute("""
                INSERT INTO player_stats (match_id,druzyna,nr,pts,p2m,p2a,p3m,p3a,ftm,fta,ast,oreb,dreb,br,fd,finishes)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (match_id, druzyna, int(nr), pts,
                  pd.get("p2m",0), pd.get("p2a",0),
                  pd.get("p3m",0), pd.get("p3a",0),
                  pd.get("ftm",0), pd.get("fta",0),
                  pd.get("ast",0), pd.get("oreb",0),
                  pd.get("dreb",0), pd.get("br",0),
                  pd.get("fd",0), pd.get("finishes",0)))

    # Shot timing
    for druzyna, stats in [("gtk", stats_gtk), ("opp", stats_opp)]:
        for b in BUCKETS:
            td = stats["timing"][b]
            cur.execute("""
                INSERT INTO timing_stats (match_id,druzyna,bucket,made2,att2,made3,att3)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (match_id, druzyna, b,
                  td["2PT"]["made"], td["2PT"]["made"]+td["2PT"]["miss"],
                  td["3PT"]["made"], td["3PT"]["made"]+td["3PT"]["miss"]))

    db.commit()
    cur.close()
    return match_id

# ══════════════════════════════════════════════════════════════════════════════
# HTML COMPONENTS
# ══════════════════════════════════════════════════════════════════════════════

CSS = """
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
:root{--gtk:#1a6b3c;--opp:#8b1a1a;--navy:#1a2b4a}
body{background:#f0f2f7;font-family:'Segoe UI',Arial,sans-serif;font-size:.9rem}

/* SIDEBAR */
.sidebar{position:fixed;top:0;left:0;height:100vh;width:240px;background:#1a2b4a;z-index:1000;display:flex;flex-direction:column;overflow-y:auto;transition:.3s}
.sidebar-logo{padding:1.25rem 1rem;border-bottom:1px solid #ffffff22}
.sidebar-logo .brand{font-size:1.1rem;font-weight:700;color:#fff}
.sidebar-logo .brand span{color:#EF9F27}
.sidebar-logo .sub{font-size:.7rem;color:#ffffff66;margin-top:2px}
.nav-section{padding:.5rem 1rem .25rem;font-size:.65rem;text-transform:uppercase;letter-spacing:1px;color:#ffffff44;font-weight:700}
.nav-item-link{display:flex;align-items:center;gap:.65rem;padding:.6rem 1rem;color:#ffffffbb;text-decoration:none;font-size:.85rem;border-radius:6px;margin:1px 8px;transition:.15s}
.nav-item-link:hover{background:#ffffff15;color:#fff}
.nav-item-link.active{background:#EF9F2722;color:#EF9F27;font-weight:600}
.nav-item-link .icon{width:18px;text-align:center;font-size:.9rem}
.nav-season{margin:0 8px 8px;padding:.5rem .75rem;background:#ffffff0f;border-radius:8px;font-size:.75rem;color:#ffffff88}
.nav-season strong{color:#EF9F27;display:block;font-size:.8rem;margin-bottom:2px}

/* MAIN */
.main-content{margin-left:240px;min-height:100vh;padding:1.5rem}

/* CARDS */
.card{border:none;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.07)}
.stat-card{background:#fff;border-radius:12px;padding:1rem;text-align:center}
.stat-val{font-size:1.6rem;font-weight:700;color:#1a2b4a;line-height:1.1}
.stat-val.sm{font-size:1.1rem}
.stat-lbl{font-size:.68rem;color:#999;text-transform:uppercase;letter-spacing:.5px;margin-top:.2rem}

/* TABLES */
.table th{background:#1a2b4a;color:#fff;font-size:.77rem;font-weight:600;border:none;padding:.45rem .6rem}
.table td{font-size:.82rem;vertical-align:middle;padding:.38rem .6rem}
.table-hover tbody tr:hover{background:#f0f4ff}

/* MISC */
.hero{background:linear-gradient(135deg,#1a2b4a,#2e5090);color:#fff;border-radius:14px;padding:1.5rem 2rem}
.page-title{font-size:1.3rem;font-weight:700;color:#1a2b4a;margin-bottom:1rem}
.badge-win{background:#c8e6c9;color:#1a6b3c;font-size:.72rem;padding:3px 8px;border-radius:20px;font-weight:700}
.badge-loss{background:#ffcdd2;color:#8b1a1a;font-size:.72rem;padding:3px 8px;border-radius:20px;font-weight:700}
.badge-draw{background:#e0e0e0;color:#555;font-size:.72rem;padding:3px 8px;border-radius:20px;font-weight:700}
.gtk-color{color:#1a6b3c;font-weight:700}
.opp-color{color:#8b1a1a;font-weight:700}
.upload-zone{border:2px dashed #1a2b4a;border-radius:14px;padding:2.5rem;text-align:center;background:#fff;cursor:pointer;transition:.2s}
.upload-zone:hover{background:#f0f4ff}
.nav-tabs .nav-link{color:#666;font-size:.83rem}
.nav-tabs .nav-link.active{color:#1a2b4a;font-weight:600}
.section-hdr{font-size:.68rem;text-transform:uppercase;letter-spacing:1px;color:#aaa;font-weight:700;margin:.75rem 0 .4rem;padding-bottom:.3rem;border-bottom:1px solid #f0f0f0}
@media(max-width:768px){.sidebar{width:60px}.sidebar .brand,.nav-section,.nav-item-link span,.nav-season{display:none}.main-content{margin-left:60px}}
</style>
"""

def nav(active="home"):
    gtk_name = "GTK"
    season = "2024/25"
    try:
        gtk_name = get_setting("gtk_name") or "GTK"
        season   = get_setting("current_season") or "2024/25"
    except: pass

    items = [
        ("home",    "/",            "🏠", "Strona główna"),
        ("history", "/historia",    "📋", "Historia meczów"),
        ("season",  "/sezon",       "📊", "Statystyki sezonu"),
        ("players", "/zawodnicy",   "👤", "Zawodnicy sezonu"),
        ("settings","ustawienia", "⚙️", "Ustawienia"),
    ]
    links = ""
    for key, href, icon, label in items:
        a_class = "nav-item-link active" if active==key else "nav-item-link"
        real_href = href if href != "ustawienia" else "/ustawienia"
        links += f'<a href="{real_href}" class="{a_class}"><span class="icon">{icon}</span><span>{label}</span></a>'

    return f"""
<div class="sidebar">
  <div class="sidebar-logo">
    <div class="brand"><span>●</span> Basket Kołcz</div>
    <div class="sub">Analytics Platform</div>
  </div>
  <div class="nav-season"><strong>{gtk_name}</strong>Sezon {season}</div>
  <div class="nav-section">Nawigacja</div>
  {links}
</div>"""

def base(content, scripts="", active="home"):
    return f"""<!DOCTYPE html><html lang="pl"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Basket Kołcz Analytics</title>{CSS}</head>
<body>
{nav(active)}
<div class="main-content">
{content}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
{scripts}
</body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
# ROUTES — STRONA GŁÓWNA
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    try: init_db()
    except: pass

    gtk_name = "GTK"
    season = "2024/25"
    try:
        gtk_name = get_setting("gtk_name") or "GTK"
        season   = get_setting("current_season") or "2024/25"
    except: pass

    # Ostatnie 5 meczów
    recent = []
    try:
        db = get_db(); cur = db.cursor()
        cur.execute("""SELECT id,data_meczu,przeciwnik,wynik_gtk,wynik_opp
                       FROM matches WHERE sezon=%s ORDER BY created_at DESC LIMIT 5""", (season,))
        recent = cur.fetchall(); cur.close()
    except: pass

    recent_rows = ""
    for m in recent:
        wynik = f"{m['wynik_gtk']}:{m['wynik_opp']}"
        if m['wynik_gtk'] > m['wynik_opp']: badge = '<span class="badge-win">W</span>'
        elif m['wynik_gtk'] < m['wynik_opp']: badge = '<span class="badge-loss">L</span>'
        else: badge = '<span class="badge-draw">D</span>'
        dt = m['data_meczu'].strftime('%d.%m.%Y') if m['data_meczu'] else "-"
        recent_rows += f"""<tr>
            <td>{badge}</td>
            <td>{dt}</td>
            <td><a href="/mecz/{m['id']}" class="fw-bold text-decoration-none" style="color:#1a2b4a">{m['przeciwnik']}</a></td>
            <td class="text-center fw-bold">{wynik}</td>
        </tr>"""

    content = f"""
<div class="hero mb-3">
  <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
    <div>
      <h1 style="font-size:1.5rem;font-weight:700;margin:0">{gtk_name} Analytics</h1>
      <p style="opacity:.75;margin:4px 0 0;font-size:.85rem">Sezon {season} · Wgraj nowy mecz</p>
    </div>
    <div class="d-flex gap-2 flex-wrap">
      <a href="/template/zapis" class="btn btn-outline-light btn-sm">📝 Zapis meczu</a>
      <a href="/template/szablon" class="btn btn-outline-light btn-sm">📋 Szablon</a>
    </div>
  </div>
</div>

<div class="row g-3">
<div class="col-lg-7">
  <div class="card p-3">
    <div class="section-hdr">Wgraj nowy mecz</div>
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <div class="row g-2 mb-2">
        <div class="col-6">
          <label class="form-label" style="font-size:.8rem;font-weight:600">Sezon</label>
          <input type="text" name="sezon" class="form-control form-control-sm" value="{season}" required>
        </div>
        <div class="col-6">
          <label class="form-label" style="font-size:.8rem;font-weight:600">Data meczu</label>
          <input type="date" name="data_meczu" class="form-control form-control-sm" value="{datetime.now().strftime('%Y-%m-%d')}">
        </div>
      </div>
      <div class="upload-zone" onclick="document.getElementById('fup').click()">
        <div style="font-size:2rem;margin-bottom:.5rem">📊</div>
        <h6 class="fw-bold mb-1" style="color:#1a2b4a">Wgraj plik zapis.xlsx</h6>
        <p class="text-muted mb-0" style="font-size:.8rem">Format: drużyna 1 = arkusz 1 ({gtk_name}), drużyna 2 = arkusz 2 (przeciwnik)</p>
        <input type="file" id="fup" name="file" accept=".xlsx" class="d-none" onchange="this.form.submit()">
      </div>
    </form>
  </div>
</div>

<div class="col-lg-5">
  <div class="card p-3">
    <div class="section-hdr">Ostatnie mecze — {season}</div>
    {'<p class="text-muted" style="font-size:.82rem">Brak meczów w tym sezonie</p>' if not recent_rows else f'<div class="table-responsive"><table class="table table-hover mb-0"><thead><tr><th></th><th>Data</th><th>Przeciwnik</th><th class="text-center">Wynik</th></tr></thead><tbody>{recent_rows}</tbody></table></div>'}
    <div class="mt-2">
      <a href="/historia" class="btn btn-outline-primary btn-sm w-100">Zobacz wszystkie mecze →</a>
    </div>
  </div>
</div>
</div>"""

    return render_template_string(base(content, active="home"))

# ══════════════════════════════════════════════════════════════════════════════
# UPLOAD
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("Nie wybrano pliku","error"); return redirect(url_for("index"))
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        flash("Plik musi być .xlsx","error"); return redirect(url_for("index"))

    sezon      = request.form.get("sezon","2024/25")
    data_meczu = request.form.get("data_meczu") or None

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        sheets = [s for s in wb.sheetnames if s.upper() not in ("META","KODY","LEGENDA")]
        if len(sheets) < 2:
            flash("Plik musi mieć 2 arkusze (GTK + przeciwnik)","error")
            return redirect(url_for("index"))

        name_gtk = sheets[0]
        name_opp = sheets[1]
        stats_gtk = parse_team_sheet(wb[name_gtk])
        stats_opp = parse_team_sheet(wb[name_opp])

        match_id = save_match_to_db(name_opp, sezon, data_meczu, stats_gtk, stats_opp)

        # Zapisz do sesji dla raportu
        session["match_id"] = match_id
        session["name_gtk"] = name_gtk
        session["name_opp"] = name_opp

        flash(f"Mecz {name_gtk} vs {name_opp} zapisany!","success")
        return redirect(url_for("mecz", match_id=match_id))

    except Exception as e:
        flash(f"Błąd: {str(e)}","error")
        return redirect(url_for("index"))

# ══════════════════════════════════════════════════════════════════════════════
# HISTORIA MECZÓW
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/historia")
def historia():
    sezon_filter = request.args.get("sezon","")
    db = get_db(); cur = db.cursor()

    if sezon_filter:
        cur.execute("SELECT * FROM matches WHERE sezon=%s ORDER BY created_at DESC", (sezon_filter,))
    else:
        cur.execute("SELECT * FROM matches ORDER BY created_at DESC")
    matches = cur.fetchall()

    cur.execute("SELECT DISTINCT sezon FROM matches ORDER BY sezon DESC")
    sezony = [r["sezon"] for r in cur.fetchall()]
    cur.close()

    rows = ""
    for i, m in enumerate(matches):
        wynik = f"{m['wynik_gtk']} : {m['wynik_opp']}"
        if m['wynik_gtk'] > m['wynik_opp']:   badge = '<span class="badge-win">WYGRANA</span>'
        elif m['wynik_gtk'] < m['wynik_opp']: badge = '<span class="badge-loss">PRZEGRANA</span>'
        else:                                  badge = '<span class="badge-draw">REMIS</span>'
        dt = m['data_meczu'].strftime('%d.%m.%Y') if m['data_meczu'] else "-"
        bg = "background:#f8f9ff" if i%2==0 else ""
        rows += f"""<tr style="{bg}">
            <td>{badge}</td>
            <td style="font-size:.8rem;color:#888">{dt}</td>
            <td><a href="/mecz/{m['id']}" class="fw-bold text-decoration-none" style="color:#1a2b4a">{m['przeciwnik']}</a></td>
            <td class="text-center"><span style="font-size:1rem;font-weight:700">{wynik}</span></td>
            <td class="text-center" style="font-size:.8rem;color:#888">{m['sezon']}</td>
            <td class="text-center">
              <a href="/mecz/{m['id']}" class="btn btn-outline-primary btn-sm" style="font-size:.75rem">Raport</a>
              <a href="/mecz/{m['id']}/delete" class="btn btn-outline-danger btn-sm ms-1" style="font-size:.75rem"
                 onclick="return confirm('Usunąć ten mecz?')">✕</a>
            </td>
        </tr>"""

    season_opts = "".join([f'<option value="{s}" {"selected" if s==sezon_filter else ""}>{s}</option>' for s in sezony])

    content = f"""
<div class="page-title">📋 Historia meczów</div>
<div class="card p-3 mb-3">
  <div class="d-flex gap-2 align-items-center flex-wrap">
    <form method="GET" class="d-flex gap-2 align-items-center">
      <label style="font-size:.82rem;font-weight:600">Sezon:</label>
      <select name="sezon" class="form-select form-select-sm" style="width:120px" onchange="this.form.submit()">
        <option value="">Wszystkie</option>
        {season_opts}
      </select>
    </form>
    <span style="font-size:.82rem;color:#888">{len(matches)} meczów</span>
  </div>
</div>
<div class="card">
  <div class="card-body p-2">
    <div class="table-responsive">
      <table class="table table-hover mb-0">
        <thead><tr>
          <th>Wynik</th><th>Data</th><th>Przeciwnik</th>
          <th class="text-center">Punkty</th><th class="text-center">Sezon</th><th class="text-center">Akcje</th>
        </tr></thead>
        <tbody>
          {rows if rows else '<tr><td colspan="6" class="text-center text-muted py-4">Brak meczów — wgraj pierwszy plik zapis.xlsx</td></tr>'}
        </tbody>
      </table>
    </div>
  </div>
</div>"""

    return render_template_string(base(content, active="history"))

# ══════════════════════════════════════════════════════════════════════════════
# RAPORT MECZU
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/mecz/<int:match_id>")
def mecz(match_id):
    db = get_db(); cur = db.cursor()
    cur.execute("SELECT * FROM matches WHERE id=%s", (match_id,))
    m = cur.fetchone()
    if not m: flash("Mecz nie istnieje","error"); return redirect(url_for("historia"))

    gtk_name = get_setting("gtk_name") or "GTK"
    name_opp = m["przeciwnik"]

    cur.execute("SELECT * FROM match_stats WHERE match_id=%s ORDER BY kwarta", (match_id,))
    all_stats = cur.fetchall()

    cur.execute("SELECT * FROM player_stats WHERE match_id=%s", (match_id,))
    all_players = cur.fetchall()

    cur.execute("SELECT * FROM timing_stats WHERE match_id=%s", (match_id,))
    all_timing = cur.fetchall()
    cur.close()

    def build_suma(druzyna):
        s = {"pts":0,"poss":0,"p2m":0,"p2a":0,"p3m":0,"p3a":0,"ftm":0,"fta":0,"br":0,"fd":0}
        for row in all_stats:
            if row["druzyna"] == druzyna:
                for k in s: s[k] += row.get(k,0)
        return s

    suma_gtk = build_suma("gtk")
    suma_opp = build_suma("opp")
    kpi_gtk  = calc_kpi(suma_gtk)
    kpi_opp  = calc_kpi(suma_opp)

    dt = m['data_meczu'].strftime('%d.%m.%Y') if m['data_meczu'] else ""

    # KPI cards
    def kpi_cards(suma, kpi):
        cards = ""
        for val, lbl in [
            (suma.get("pts",0),"Punkty"),
            (kpi["efg"],"eFG%"),
            (kpi["ts"],"TS%"),
            (kpi["ortg"],"ORtg"),
            (kpi["ppp"],"PPP"),
            (kpi["p2_pct"],"2PT%"),
            (kpi["p3_pct"],"3PT%"),
            (kpi["ft_pct"],"FT%"),
        ]:
            cards += f'<div class="col-6 col-sm-3 col-md-2"><div class="stat-card"><div class="stat-val sm">{val}</div><div class="stat-lbl">{lbl}</div></div></div>'
        return cards

    # Per kwarta tabela
    def q_table(druzyna):
        rows = ""
        for qn in [1,2,3,4]:
            qd = next((r for r in all_stats if r["druzyna"]==druzyna and r["kwarta"]==qn), {})
            kq = calc_kpi(dict(qd) if qd else {})
            rows += f"""<tr>
                <td><span class="badge" style="background:{'#c8e6c9' if qn==1 else '#bbdefb' if qn==2 else '#fff9c4' if qn==3 else '#fce4ec'};color:#333">{qn}Q</span></td>
                <td><b>{qd.get('pts',0)}</b></td>
                <td>{qd.get('p2m',0)}/{qd.get('p2a',0)}</td><td>{kq['p2_pct']}</td>
                <td>{qd.get('p3m',0)}/{qd.get('p3a',0)}</td><td>{kq['p3_pct']}</td>
                <td>{qd.get('ftm',0)}/{qd.get('fta',0)}</td>
                <td>{qd.get('br',0)}</td><td>{qd.get('poss',0)}</td>
                <td>{kq['efg']}</td><td>{kq['ortg']}</td>
            </tr>"""
        sk = calc_kpi(suma_gtk if druzyna=="gtk" else suma_opp)
        sd = suma_gtk if druzyna=="gtk" else suma_opp
        rows += f"""<tr style="background:#f0f4ff;font-weight:700">
            <td>SUMA</td><td><b>{sd.get('pts',0)}</b></td>
            <td>{sd.get('p2m',0)}/{sd.get('p2a',0)}</td><td>{sk['p2_pct']}</td>
            <td>{sd.get('p3m',0)}/{sd.get('p3a',0)}</td><td>{sk['p3_pct']}</td>
            <td>{sd.get('ftm',0)}/{sd.get('fta',0)}</td>
            <td>{sd.get('br',0)}</td><td>{sd.get('poss',0)}</td>
            <td>{sk['efg']}</td><td>{sk['ortg']}</td>
        </tr>"""
        return f"""<div class="table-responsive"><table class="table table-hover mb-0">
            <thead><tr><th>Q</th><th>PKT</th><th>2PM/A</th><th>2P%</th><th>3PM/A</th><th>3P%</th><th>FTM/A</th><th>BR</th><th>POSS</th><th>eFG%</th><th>ORtg</th></tr></thead>
            <tbody>{rows}</tbody></table></div>"""

    # Zawodnicy tabela
    def p_table(druzyna):
        players = [r for r in all_players if r["druzyna"]==druzyna]
        if not players:
            return '<p class="text-muted p-2" style="font-size:.82rem">Brak danych zawodników</p>'
        rows = ""
        for pd in sorted(players, key=lambda x: x["pts"], reverse=True):
            fga = pd.get("p2a",0)+pd.get("p3a",0)
            fta = pd.get("fta",0); ftm = pd.get("ftm",0)
            pm2 = pd.get("p2m",0); pm3 = pd.get("p3m",0)
            efg = f"{(pm2+1.5*pm3)/fga:.0%}" if fga else "-"
            ts  = f"{pd.get('pts',0)/(2*(fga+0.44*fta)):.0%}" if (fga+fta) else "-"
            rows += f"""<tr>
                <td class="fw-bold">#{pd['nr']}</td>
                <td class="fw-bold" style="color:#1a2b4a">{pd.get('pts',0)}</td>
                <td>{pm2}/{pd.get('p2a',0)}</td>
                <td>{pm3}/{pd.get('p3a',0)}</td>
                <td>{ftm}/{fta}</td>
                <td><b>{efg}</b></td><td>{ts}</td>
                <td>{pd.get('ast',0)}</td><td>{pd.get('oreb',0)}</td>
                <td>{pd.get('dreb',0)}</td><td>{pd.get('br',0)}</td>
                <td>{pd.get('finishes',0)}</td>
            </tr>"""
        return f"""<div class="table-responsive"><table class="table table-hover mb-0">
            <thead><tr><th>#</th><th>PTS</th><th>2PM/A</th><th>3PM/A</th><th>FTM/A</th><th>eFG%</th><th>TS%</th><th>AST</th><th>OREB</th><th>DREB</th><th>BR</th><th>FIN</th></tr></thead>
            <tbody>{rows}</tbody></table></div>"""

    # Shot timing
    def tim_table(druzyna):
        rows = ""
        for b in BUCKETS:
            td = next((r for r in all_timing if r["druzyna"]==druzyna and r["bucket"]==b), {})
            m2=td.get("made2",0); a2=td.get("att2",0)
            m3=td.get("made3",0); a3=td.get("att3",0)
            tot=m2+m3; att=a2+a3
            eff=f"{tot/att:.0%}" if att else "-"
            rows += f"<tr><td><b>{b}</b></td><td>{m2}/{a2}</td><td>{m3}/{a3}</td><td><b>{eff}</b></td></tr>"
        return f"""<div class="table-responsive"><table class="table table-hover mb-0">
            <thead><tr><th>Czas</th><th>2PT</th><th>3PT</th><th>Eff%</th></tr></thead>
            <tbody>{rows}</tbody></table></div>"""

    pts_q_gtk = [next((r["pts"] for r in all_stats if r["druzyna"]=="gtk" and r["kwarta"]==q),0) for q in [1,2,3,4]]
    pts_q_opp = [next((r["pts"] for r in all_stats if r["druzyna"]=="opp" and r["kwarta"]==q),0) for q in [1,2,3,4]]

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <div>
    <div class="page-title mb-0">{gtk_name} vs {name_opp}</div>
    <div style="font-size:.8rem;color:#888">{dt} · Sezon {m['sezon']}</div>
  </div>
  <div class="d-flex gap-2">
    <a href="/historia" class="btn btn-outline-secondary btn-sm">← Historia</a>
    <a href="/mecz/{match_id}/export/xlsx" class="btn btn-warning btn-sm fw-bold">⬇ Excel</a>
    <a href="/mecz/{match_id}/export/pdf" class="btn btn-danger btn-sm fw-bold">⬇ PDF</a>
  </div>
</div>

<div class="hero mb-3">
  <div class="d-flex justify-content-between align-items-center flex-wrap gap-3">
    <div>
      <div style="font-size:.8rem;opacity:.7">{gtk_name}</div>
      <div style="font-size:2rem;font-weight:700">{m['wynik_gtk']}</div>
    </div>
    <div class="text-center">
      {'<span class="badge-win" style="font-size:.9rem;padding:6px 14px">WYGRANA</span>' if m['wynik_gtk']>m['wynik_opp'] else '<span class="badge-loss" style="font-size:.9rem;padding:6px 14px">PRZEGRANA</span>' if m['wynik_gtk']<m['wynik_opp'] else '<span class="badge-draw" style="font-size:.9rem;padding:6px 14px">REMIS</span>'}
    </div>
    <div class="text-end">
      <div style="font-size:.8rem;opacity:.7">{name_opp}</div>
      <div style="font-size:2rem;font-weight:700">{m['wynik_opp']}</div>
    </div>
  </div>
</div>

<ul class="nav nav-tabs mb-2" id="mainTabs">
  <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tabGTK">{gtk_name}</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabOPP">{name_opp}</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabCMP">Porównanie</button></li>
</ul>

<div class="tab-content">

<div class="tab-pane fade show active" id="tabGTK">
  <div class="row g-2 mb-2">{kpi_cards(suma_gtk, kpi_gtk)}</div>
  <ul class="nav nav-tabs mt-2 mb-1" id="gtkTabs">
    <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#gtk_q">Per kwarta</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#gtk_p">Zawodnicy</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#gtk_t">Shot Timing</button></li>
  </ul>
  <div class="tab-content">
    <div class="tab-pane fade show active" id="gtk_q"><div class="card mt-1"><div class="card-body p-2">{q_table('gtk')}</div></div></div>
    <div class="tab-pane fade" id="gtk_p"><div class="card mt-1"><div class="card-body p-2">{p_table('gtk')}</div></div></div>
    <div class="tab-pane fade" id="gtk_t"><div class="card mt-1"><div class="card-body p-2">{tim_table('gtk')}</div></div></div>
  </div>
</div>

<div class="tab-pane fade" id="tabOPP">
  <div class="row g-2 mb-2">{kpi_cards(suma_opp, kpi_opp)}</div>
  <ul class="nav nav-tabs mt-2 mb-1" id="oppTabs">
    <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#opp_q">Per kwarta</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#opp_p">Zawodnicy</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#opp_t">Shot Timing</button></li>
  </ul>
  <div class="tab-content">
    <div class="tab-pane fade show active" id="opp_q"><div class="card mt-1"><div class="card-body p-2">{q_table('opp')}</div></div></div>
    <div class="tab-pane fade" id="opp_p"><div class="card mt-1"><div class="card-body p-2">{p_table('opp')}</div></div></div>
    <div class="tab-pane fade" id="opp_t"><div class="card mt-1"><div class="card-body p-2">{tim_table('opp')}</div></div></div>
  </div>
</div>

<div class="tab-pane fade" id="tabCMP">
  <ul class="nav nav-tabs mt-2 mb-1" id="cmpTabs">
    <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#cmp_metrics">Kluczowe Metryki</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#cmp_timing">Shot Timing</button></li>
  </ul>
  <div class="tab-content">

  <div class="tab-pane fade show active" id="cmp_metrics">
    <div class="row g-3 mt-1">
      <div class="col-lg-5">
        <div class="card"><div class="card-body p-2">
          <table class="table table-sm table-hover mb-0">
            <thead><tr>
              <th>Metryka</th>
              <th class="text-center gtk-color">{gtk_name}</th>
              <th class="text-center opp-color">{name_opp}</th>
            </tr></thead>
            <tbody>
              {''.join(f'''<tr>
                <td style="font-size:.82rem"><b>{l}</b><br><span style="font-size:.7rem;color:#aaa">{desc}</span>
                  <div style="display:flex;height:5px;border-radius:3px;overflow:hidden;margin-top:3px">
                    <div style="flex:{fa};background:{'#1a6b3c' if fa>=fb else '#ddd'}"></div>
                    <div style="flex:{fb};background:{'#8b1a1a' if fb>fa else '#ddd'}"></div>
                  </div>
                </td>
                <td class="text-center" style="{'font-weight:700;color:#1a6b3c' if fa>fb else ''}">{va}</td>
                <td class="text-center" style="{'font-weight:700;color:#8b1a1a' if fb>fa else ''}">{vb}</td>
              </tr>''' for l,va,vb,desc,fa,fb in [
                ("Punkty",suma_gtk.get('pts',0),suma_opp.get('pts',0),"Łączna liczba punktów",suma_gtk.get('pts',0),suma_opp.get('pts',0)),
                ("Posiadania",suma_gtk.get('poss',0),suma_opp.get('poss',0),"Liczba posiadań",suma_gtk.get('poss',0),suma_opp.get('poss',0)),
                ("eFG%",kpi_gtk['efg'],kpi_opp['efg'],"Efektywny % rzutów z pola",
                 float(kpi_gtk['efg'].replace('%','')) if kpi_gtk['efg']!='-' else 0,
                 float(kpi_opp['efg'].replace('%','')) if kpi_opp['efg']!='-' else 0),
                ("TS%",kpi_gtk['ts'],kpi_opp['ts'],"Prawdziwy % skuteczności",
                 float(kpi_gtk['ts'].replace('%','')) if kpi_gtk['ts']!='-' else 0,
                 float(kpi_opp['ts'].replace('%','')) if kpi_opp['ts']!='-' else 0),
                ("ORtg",kpi_gtk['ortg'],kpi_opp['ortg'],"Punkty na 100 posiadań",
                 float(kpi_gtk['ortg']) if kpi_gtk['ortg']!='-' else 0,
                 float(kpi_opp['ortg']) if kpi_opp['ortg']!='-' else 0),
                ("PPP",kpi_gtk['ppp'],kpi_opp['ppp'],"Punkty na posiadanie",
                 float(kpi_gtk['ppp']) if kpi_gtk['ppp']!='-' else 0,
                 float(kpi_opp['ppp']) if kpi_opp['ppp']!='-' else 0),
                ("2PT%",kpi_gtk['p2_pct'],kpi_opp['p2_pct'],"Skuteczność za 2 pkt",
                 float(kpi_gtk['p2_pct'].replace('%','')) if kpi_gtk['p2_pct']!='-' else 0,
                 float(kpi_opp['p2_pct'].replace('%','')) if kpi_opp['p2_pct']!='-' else 0),
                ("3PT%",kpi_gtk['p3_pct'],kpi_opp['p3_pct'],"Skuteczność za 3 pkt",
                 float(kpi_gtk['p3_pct'].replace('%','')) if kpi_gtk['p3_pct']!='-' else 0,
                 float(kpi_opp['p3_pct'].replace('%','')) if kpi_opp['p3_pct']!='-' else 0),
                ("FT%",kpi_gtk['ft_pct'],kpi_opp['ft_pct'],"Skuteczność rzutów wolnych",
                 float(kpi_gtk['ft_pct'].replace('%','')) if kpi_gtk['ft_pct']!='-' else 0,
                 float(kpi_opp['ft_pct'].replace('%','')) if kpi_opp['ft_pct']!='-' else 0),
                ("FT Rate",kpi_gtk['ftr'],kpi_opp['ftr'],"FTA / FGA",
                 float(kpi_gtk['ftr']) if kpi_gtk['ftr']!='-' else 0,
                 float(kpi_opp['ftr']) if kpi_opp['ftr']!='-' else 0),
                ("Straty (BR)",suma_gtk.get('br',0),suma_opp.get('br',0),"Liczba strat — niższy = lepszy",
                 suma_opp.get('br',0),suma_gtk.get('br',0)),
                ("Faule wymuszone",suma_gtk.get('fd',0),suma_opp.get('fd',0),"Liczba wymuszonych fauli",
                 suma_gtk.get('fd',0),suma_opp.get('fd',0)),
                ("2PM/A",f"{suma_gtk.get('p2m',0)}/{suma_gtk.get('p2a',0)}",f"{suma_opp.get('p2m',0)}/{suma_opp.get('p2a',0)}","Celne / próby za 2 pkt",suma_gtk.get('p2m',0),suma_opp.get('p2m',0)),
                ("3PM/A",f"{suma_gtk.get('p3m',0)}/{suma_gtk.get('p3a',0)}",f"{suma_opp.get('p3m',0)}/{suma_opp.get('p3a',0)}","Celne / próby za 3 pkt",suma_gtk.get('p3m',0),suma_opp.get('p3m',0)),
                ("FTM/A",f"{suma_gtk.get('ftm',0)}/{suma_gtk.get('fta',0)}",f"{suma_opp.get('ftm',0)}/{suma_opp.get('fta',0)}","Celne / próby rzutów wolnych",suma_gtk.get('ftm',0),suma_opp.get('ftm',0)),
              ])}
            </tbody>
          </table>
        </div></div>
      </div>
      <div class="col-lg-7">
        <div class="card"><div class="card-body p-3">
          <div class="section-hdr">Punkty per kwarta</div>
          <canvas id="qChart"></canvas>
        </div></div>
      </div>
    </div>
  </div>

  <div class="tab-pane fade" id="cmp_timing">
    <div class="card mt-2"><div class="card-body p-2">
      <p class="text-muted mb-2" style="font-size:.8rem">Porównanie skuteczności rzutów według czasu trwania posiadania (zegar 24s)</p>
      <div class="d-flex gap-3 mb-2" style="font-size:.78rem">
        <span><span style="display:inline-block;width:12px;height:8px;background:#1a6b3c;border-radius:2px;margin-right:4px"></span>{gtk_name}</span>
        <span><span style="display:inline-block;width:12px;height:8px;background:#8b1a1a;border-radius:2px;margin-right:4px"></span>{name_opp}</span>
      </div>
      <div class="table-responsive">
        <table class="table table-hover mb-0">
          <thead><tr>
            <th>Czas</th>
            <th class="text-center" style="color:#1a6b3c">Celne/Próby</th>
            <th class="text-center" style="color:#1a6b3c">Eff%</th>
            <th style="width:100px">{gtk_name}</th>
            <th style="width:100px">{name_opp}</th>
            <th class="text-center" style="color:#8b1a1a">Eff%</th>
            <th class="text-center" style="color:#8b1a1a">Celne/Próby</th>
          </tr></thead>
          <tbody>
            {''.join(f"""<tr>
              <td class="fw-bold" style="font-size:.82rem">{b}</td>
              <td class="text-center" style="font-size:.82rem">{(lambda gd: f"{gd.get('made2',0)+gd.get('made3',0)}/{gd.get('att2',0)+gd.get('att3',0)}")(next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==b),{}))}</td>
              <td class="text-center" style="font-size:.82rem;font-weight:600;color:#1a6b3c">{(lambda gd: f"{(gd.get('made2',0)+gd.get('made3',0))/(gd.get('att2',0)+gd.get('att3',0)):.0%}" if (gd.get('att2',0)+gd.get('att3',0)) else "-")(next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==b),{}))}</td>
              <td><div style="height:8px;width:{int((next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==b),{}).get('att2',0)+next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==b),{}).get('att3',0))/max(max((next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==bb),{}).get('att2',0)+next((r for r in all_timing if r['druzyna']=='gtk' and r['bucket']==bb),{}).get('att3',0)) for bb in BUCKETS),1)*80)}px;background:#1a6b3c;border-radius:4px"></div></td>
              <td><div style="height:8px;width:{int((next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==b),{}).get('att2',0)+next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==b),{}).get('att3',0))/max(max((next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==bb),{}).get('att2',0)+next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==bb),{}).get('att3',0)) for bb in BUCKETS),1)*80)}px;background:#8b1a1a;border-radius:4px"></div></td>
              <td class="text-center" style="font-size:.82rem;font-weight:600;color:#8b1a1a">{(lambda od: f"{(od.get('made2',0)+od.get('made3',0))/(od.get('att2',0)+od.get('att3',0)):.0%}" if (od.get('att2',0)+od.get('att3',0)) else "-")(next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==b),{}))}</td>
              <td class="text-center" style="font-size:.82rem">{(lambda od: f"{od.get('made2',0)+od.get('made3',0)}/{od.get('att2',0)+od.get('att3',0)}")(next((r for r in all_timing if r['druzyna']=='opp' and r['bucket']==b),{}))}</td>
            </tr>""" for b in BUCKETS)}
          </tbody>
        </table>
      </div>
    </div></div>
  </div>

  </div>
</div>

</div>"""

    scripts = f"""<script>
new Chart(document.getElementById('qChart'),{{
  type:'bar',
  data:{{labels:['1Q','2Q','3Q','4Q'],datasets:[
    {{label:'{gtk_name}',data:{pts_q_gtk},backgroundColor:'#1a6b3c88',borderColor:'#1a6b3c',borderWidth:2,borderRadius:6}},
    {{label:'{name_opp}',data:{pts_q_opp},backgroundColor:'#8b1a1a88',borderColor:'#8b1a1a',borderWidth:2,borderRadius:6}}
  ]}},
  options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},scales:{{y:{{beginAtZero:true}}}}}}
}});
</script>"""

    return render_template_string(base(content, scripts, active="history"))

@app.route("/mecz/<int:match_id>/delete")
def mecz_delete(match_id):
    db = get_db(); cur = db.cursor()
    cur.execute("DELETE FROM matches WHERE id=%s", (match_id,))
    db.commit(); cur.close()
    flash("Mecz usunięty","success")
    return redirect(url_for("historia"))

# ══════════════════════════════════════════════════════════════════════════════
# STATYSTYKI SEZONU
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/sezon")
def sezon():
    sezon_filter = request.args.get("sezon", get_setting("current_season") or "2024/25")
    db = get_db(); cur = db.cursor()

    cur.execute("SELECT DISTINCT sezon FROM matches ORDER BY sezon DESC")
    sezony = [r["sezon"] for r in cur.fetchall()]

    cur.execute("SELECT COUNT(*) as cnt FROM matches WHERE sezon=%s", (sezon_filter,))
    n_matches = cur.fetchone()["cnt"]

    if n_matches == 0:
        cur.close()
        content = f"""
<div class="page-title">📊 Statystyki sezonu {sezon_filter}</div>
<div class="card p-4 text-center text-muted">Brak meczów w sezonie {sezon_filter}. Wgraj pliki zapis.xlsx.</div>"""
        return render_template_string(base(content, active="season"))

    # Agregaty GTK
    cur.execute("""
        SELECT
            SUM(pts) as pts, SUM(poss) as poss,
            SUM(p2m) as p2m, SUM(p2a) as p2a,
            SUM(p3m) as p3m, SUM(p3a) as p3a,
            SUM(ftm) as ftm, SUM(fta) as fta,
            SUM(br)  as br,  SUM(fd)  as fd
        FROM match_stats ms
        JOIN matches m ON ms.match_id = m.id
        WHERE m.sezon=%s AND ms.druzyna='gtk'
    """, (sezon_filter,))
    gtk_tot = dict(cur.fetchone())

    # Agregaty OPP
    cur.execute("""
        SELECT
            SUM(pts) as pts, SUM(poss) as poss,
            SUM(p2m) as p2m, SUM(p2a) as p2a,
            SUM(p3m) as p3m, SUM(p3a) as p3a,
            SUM(ftm) as ftm, SUM(fta) as fta,
            SUM(br)  as br,  SUM(fd)  as fd
        FROM match_stats ms
        JOIN matches m ON ms.match_id = m.id
        WHERE m.sezon=%s AND ms.druzyna='opp'
    """, (sezon_filter,))
    opp_tot = dict(cur.fetchone())

    # Wyniki meczów
    cur.execute("""SELECT wynik_gtk, wynik_opp FROM matches WHERE sezon=%s""", (sezon_filter,))
    results = cur.fetchall()
    wins   = sum(1 for r in results if r["wynik_gtk"] > r["wynik_opp"])
    losses = sum(1 for r in results if r["wynik_gtk"] < r["wynik_opp"])
    draws  = n_matches - wins - losses

    # Shot timing per sezon
    cur.execute("""
        SELECT ts.druzyna, ts.bucket,
               SUM(ts.made2) as made2, SUM(ts.att2) as att2,
               SUM(ts.made3) as made3, SUM(ts.att3) as att3
        FROM timing_stats ts
        JOIN matches m ON ts.match_id=m.id
        WHERE m.sezon=%s
        GROUP BY ts.druzyna, ts.bucket
    """, (sezon_filter,))
    timing_rows = cur.fetchall()
    cur.close()

    gtk_name = get_setting("gtk_name") or "GTK"

    def avg(d, k):
        v = d.get(k) or 0
        return round(v / n_matches, 1)

    def avg_kpi(d):
        k = calc_kpi(d)
        return k

    gtk_kpi = calc_kpi(gtk_tot)
    opp_kpi = calc_kpi(opp_tot)

    def cmp_row(lbl, vg, vo, higher_is_better=True):
        try:
            fg = float(str(vg).replace('%','').replace('-','0'))
            fo = float(str(vo).replace('%','').replace('-','0'))
            sg = "font-weight:700;color:#1a6b3c" if (higher_is_better and fg>fo) or (not higher_is_better and fg<fo) else ""
            so = "font-weight:700;color:#8b1a1a" if (higher_is_better and fo>fg) or (not higher_is_better and fo<fg) else ""
        except: sg=so=""
        return f"<tr><td><b>{lbl}</b></td><td class='text-center' style='{sg}'>{vg}</td><td class='text-center' style='{so}'>{vo}</td></tr>"

    kpi_rows = (
        cmp_row("Pkt / mecz",       avg(gtk_tot,"pts"),  avg(opp_tot,"pts")) +
        cmp_row("Posiadania / mecz", avg(gtk_tot,"poss"), avg(opp_tot,"poss")) +
        cmp_row("eFG%",             gtk_kpi["efg"],      opp_kpi["efg"]) +
        cmp_row("TS%",              gtk_kpi["ts"],       opp_kpi["ts"]) +
        cmp_row("ORtg",             gtk_kpi["ortg"],     opp_kpi["ortg"]) +
        cmp_row("PPP",              gtk_kpi["ppp"],      opp_kpi["ppp"]) +
        cmp_row("2PT%",             gtk_kpi["p2_pct"],   opp_kpi["p2_pct"]) +
        cmp_row("3PT%",             gtk_kpi["p3_pct"],   opp_kpi["p3_pct"]) +
        cmp_row("FT%",              gtk_kpi["ft_pct"],   opp_kpi["ft_pct"]) +
        cmp_row("Straty / mecz",    avg(gtk_tot,"br"),   avg(opp_tot,"br"), False) +
        cmp_row("FT Rate",          gtk_kpi["ftr"],      opp_kpi["ftr"])
    )

    # Timing tabela
    def timing_row(bucket):
        gd = next((r for r in timing_rows if r["druzyna"]=="gtk" and r["bucket"]==bucket), {})
        od = next((r for r in timing_rows if r["druzyna"]=="opp" and r["bucket"]==bucket), {})
        gm=gd.get("made2",0)+gd.get("made3",0); ga=gd.get("att2",0)+gd.get("att3",0)
        om=od.get("made2",0)+od.get("made3",0); oa=od.get("att2",0)+od.get("att3",0)
        ge=f"{gm/ga:.0%}" if ga else "-"; oe=f"{om/oa:.0%}" if oa else "-"
        return f"<tr><td><b>{bucket}</b></td><td>{gm}/{ga}</td><td style='color:#1a6b3c;font-weight:600'>{ge}</td><td>{om}/{oa}</td><td style='color:#8b1a1a;font-weight:600'>{oe}</td></tr>"

    tim_rows = "".join(timing_row(b) for b in BUCKETS)

    pts_per_match_gtk = [0,0,0,0]
    pts_per_match_opp = [0,0,0,0]
    season_opts = "".join([f'<option value="{s}" {"selected" if s==sezon_filter else ""}>{s}</option>' for s in sezony])

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <div class="page-title mb-0">📊 Statystyki sezonu</div>
  <form method="GET" class="d-flex gap-2 align-items-center">
    <label style="font-size:.82rem;font-weight:600">Sezon:</label>
    <select name="sezon" class="form-select form-select-sm" style="width:120px" onchange="this.form.submit()">
      {season_opts}
    </select>
  </form>
</div>

<div class="row g-2 mb-3">
  <div class="col-6 col-md-3"><div class="stat-card"><div class="stat-val">{n_matches}</div><div class="stat-lbl">Mecze</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card"><div class="stat-val" style="color:#1a6b3c">{wins}</div><div class="stat-lbl">Wygrane</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card"><div class="stat-val" style="color:#8b1a1a">{losses}</div><div class="stat-lbl">Przegrane</div></div></div>
  <div class="col-6 col-md-3"><div class="stat-card"><div class="stat-val">{wins/n_matches:.0%}</div><div class="stat-lbl">Win%</div></div></div>
</div>

<ul class="nav nav-tabs mb-2">
  <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#sMetrics">Metryki średnie</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#sTiming">Shot Timing</button></li>
</ul>

<div class="tab-content">

<div class="tab-pane fade show active" id="sMetrics">
  <div class="row g-3">
    <div class="col-lg-6">
      <div class="card"><div class="card-body p-2">
        <div class="section-hdr">Średnie na mecz — {gtk_name} vs Przeciwnicy</div>
        <p style="font-size:.75rem;color:#aaa">Suma wszystkich danych ÷ liczba meczów ({n_matches})</p>
        <table class="table table-sm table-hover mb-0">
          <thead><tr>
            <th>Metryka</th>
            <th class="text-center" style="color:#1a6b3c">{gtk_name}</th>
            <th class="text-center" style="color:#8b1a1a">Przeciwnicy</th>
          </tr></thead>
          <tbody>{kpi_rows}</tbody>
        </table>
      </div></div>
    </div>
    <div class="col-lg-6">
      <div class="card"><div class="card-body p-3">
        <div class="section-hdr">Bilans meczów</div>
        <canvas id="winChart" style="max-height:200px"></canvas>
      </div></div>
    </div>
  </div>
</div>

<div class="tab-pane fade" id="sTiming">
  <div class="card mt-1"><div class="card-body p-2">
    <div class="section-hdr">Shot Timing sezonu (łącznie)</div>
    <table class="table table-hover mb-0">
      <thead><tr>
        <th>Czas</th>
        <th class="text-center" style="color:#1a6b3c">{gtk_name} Made/Att</th>
        <th class="text-center" style="color:#1a6b3c">Eff%</th>
        <th class="text-center" style="color:#8b1a1a">Opp Made/Att</th>
        <th class="text-center" style="color:#8b1a1a">Eff%</th>
      </tr></thead>
      <tbody>{tim_rows}</tbody>
    </table>
  </div></div>
</div>

</div>"""

    scripts = f"""<script>
new Chart(document.getElementById('winChart'),{{
  type:'doughnut',
  data:{{
    labels:['Wygrane','Przegrane','Remisy'],
    datasets:[{{data:[{wins},{losses},{draws}],
      backgroundColor:['#1a6b3c','#8b1a1a','#888'],
      borderWidth:0}}]
  }},
  options:{{responsive:true,plugins:{{legend:{{position:'bottom'}}}}}}
}});
</script>"""

    return render_template_string(base(content, scripts, active="season"))

# ══════════════════════════════════════════════════════════════════════════════
# ZAWODNICY SEZONU
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/zawodnicy")
def zawodnicy():
    sezon_filter = request.args.get("sezon", get_setting("current_season") or "2024/25")
    db = get_db(); cur = db.cursor()

    cur.execute("SELECT DISTINCT sezon FROM matches ORDER BY sezon DESC")
    sezony = [r["sezon"] for r in cur.fetchall()]

    cur.execute("SELECT COUNT(*) as cnt FROM matches WHERE sezon=%s", (sezon_filter,))
    n_matches = cur.fetchone()["cnt"]

    cur.execute("""
        SELECT ps.nr,
               SUM(ps.pts) as pts, SUM(ps.p2m) as p2m, SUM(ps.p2a) as p2a,
               SUM(ps.p3m) as p3m, SUM(ps.p3a) as p3a,
               SUM(ps.ftm) as ftm, SUM(ps.fta) as fta,
               SUM(ps.ast) as ast, SUM(ps.oreb) as oreb, SUM(ps.dreb) as dreb,
               SUM(ps.br) as br, SUM(ps.fd) as fd, SUM(ps.finishes) as finishes,
               COUNT(DISTINCT ps.match_id) as mecze
        FROM player_stats ps
        JOIN matches m ON ps.match_id=m.id
        WHERE m.sezon=%s AND ps.druzyna='gtk'
        GROUP BY ps.nr ORDER BY pts DESC
    """, (sezon_filter,))
    players = cur.fetchall()
    cur.close()

    rows = ""
    for i, p in enumerate(players):
        fga = p.get("p2a",0)+p.get("p3a",0)
        fta = p.get("fta",0); ftm = p.get("ftm",0)
        pm2=p.get("p2m",0); pm3=p.get("p3m",0)
        efg=f"{(pm2+1.5*pm3)/fga:.1%}" if fga else "-"
        ts =f"{p.get('pts',0)/(2*(fga+0.44*fta)):.1%}" if (fga+fta) else "-"
        n = p.get("mecze",1); bg = "background:#f8f9ff" if i%2==0 else ""
        rows += f"""<tr style="{bg}">
            <td class="fw-bold">#{p['nr']}</td>
            <td class="fw-bold" style="color:#1a2b4a;font-size:.95rem">{p.get('pts',0)}</td>
            <td style="font-size:.78rem;color:#888">{p.get('pts',0)/n:.1f}</td>
            <td>{pm2}/{p.get('p2a',0)}</td>
            <td>{pm3}/{p.get('p3a',0)}</td>
            <td>{ftm}/{fta}</td>
            <td><b>{efg}</b></td><td>{ts}</td>
            <td>{p.get('ast',0)}</td>
            <td>{p.get('oreb',0)}</td><td>{p.get('dreb',0)}</td>
            <td>{p.get('br',0)}</td>
            <td>{p.get('finishes',0)}</td>
            <td style="font-size:.78rem;color:#888">{n}</td>
        </tr>"""

    season_opts = "".join([f'<option value="{s}" {"selected" if s==sezon_filter else ""}>{s}</option>' for s in sezony])
    gtk_name = get_setting("gtk_name") or "GTK"

    content = f"""
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <div class="page-title mb-0">👤 Zawodnicy sezonu — {gtk_name}</div>
  <form method="GET" class="d-flex gap-2 align-items-center">
    <select name="sezon" class="form-select form-select-sm" style="width:120px" onchange="this.form.submit()">
      {season_opts}
    </select>
  </form>
</div>
<div class="card">
  <div class="card-body p-2">
    <p style="font-size:.78rem;color:#aaa">Suma wszystkich meczów sezonu {sezon_filter} ({n_matches} meczów)</p>
    <div class="table-responsive">
      <table class="table table-hover mb-0">
        <thead><tr>
          <th>#</th><th>PTS łącznie</th><th>PPG</th>
          <th>2PM/A</th><th>3PM/A</th><th>FTM/A</th>
          <th>eFG%</th><th>TS%</th><th>AST</th>
          <th>OREB</th><th>DREB</th><th>BR</th><th>FIN</th><th>Mecze</th>
        </tr></thead>
        <tbody>
          {rows if rows else '<tr><td colspan="14" class="text-center text-muted py-4">Brak danych zawodników</td></tr>'}
        </tbody>
      </table>
    </div>
  </div>
</div>"""

    return render_template_string(base(content, active="players"))

# ══════════════════════════════════════════════════════════════════════════════
# USTAWIENIA
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/ustawienia", methods=["GET","POST"])
def ustawienia():
    if request.method == "POST":
        set_setting("gtk_name", request.form.get("gtk_name","GTK"))
        set_setting("current_season", request.form.get("current_season","2024/25"))
        flash("Ustawienia zapisane!","success")
        return redirect(url_for("ustawienia"))

    gtk_name = get_setting("gtk_name") or "GTK"
    season   = get_setting("current_season") or "2024/25"

    content = f"""
<div class="page-title">⚙️ Ustawienia</div>
<div class="row justify-content-center">
<div class="col-lg-6">
  <div class="card p-3">
    <form method="POST">
      <div class="mb-3">
        <label class="form-label fw-bold">Nazwa drużyny głównej</label>
        <input type="text" name="gtk_name" class="form-control" value="{gtk_name}" placeholder="np. GTK, Kotwica, AZS...">
        <div class="form-text">Ta nazwa pojawi się we wszystkich raportach i statystykach.</div>
      </div>
      <div class="mb-3">
        <label class="form-label fw-bold">Aktualny sezon</label>
        <input type="text" name="current_season" class="form-control" value="{season}" placeholder="np. 2024/25">
        <div class="form-text">Domyślny sezon przy wgrywaniu nowych meczów.</div>
      </div>
      <button type="submit" class="btn btn-primary w-100">Zapisz ustawienia</button>
    </form>
  </div>

  <div class="card mt-3 p-3">
    <div class="section-hdr">Pobierz szablony</div>
    <div class="row g-2">
      <div class="col-6">
        <a href="/template/zapis" class="btn btn-outline-primary w-100 btn-sm">📝 Zapis meczu</a>
      </div>
      <div class="col-6">
        <a href="/template/szablon" class="btn btn-outline-success w-100 btn-sm">📋 Szablon raportu</a>
      </div>
    </div>
  </div>
</div></div>"""

    flash_html = ""
    return render_template_string(base(content, active="settings"))

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT (z match_id)
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/mecz/<int:match_id>/export/xlsx")
def export_match_xlsx(match_id):
    db = get_db(); cur = db.cursor()
    cur.execute("SELECT * FROM matches WHERE id=%s", (match_id,))
    m = cur.fetchone()
    if not m: return redirect(url_for("historia"))

    gtk_name = get_setting("gtk_name") or "GTK"
    name_opp = m["przeciwnik"]

    cur.execute("SELECT * FROM match_stats WHERE match_id=%s ORDER BY kwarta", (match_id,))
    all_stats = cur.fetchall()
    cur.execute("SELECT * FROM player_stats WHERE match_id=%s", (match_id,))
    all_players = cur.fetchall()
    cur.execute("SELECT * FROM timing_stats WHERE match_id=%s", (match_id,))
    all_timing = cur.fetchall()
    cur.close()

    wb = openpyxl.Workbook()
    HDR = PatternFill("solid", fgColor="1A2B4A")
    HDR_F = Font(color="FFFFFF", bold=True, size=10)
    CTR = Alignment(horizontal="center", vertical="center")
    BD = Border(bottom=Side(style="thin",color="CCCCCC"),right=Side(style="thin",color="CCCCCC"))

    def hdr_row(ws, row, labels):
        for i,l in enumerate(labels):
            c = ws.cell(row, i+1, l); c.fill=HDR; c.font=HDR_F; c.alignment=CTR

    # Ogólne
    ws1 = wb.active; ws1.title = "OGÓLNE"
    ws1.merge_cells("A1:M1")
    t=ws1["A1"]; t.value=f"RAPORT — {gtk_name} vs {name_opp}"; t.fill=HDR; t.font=Font(color="FFFFFF",bold=True,size=12); t.alignment=CTR
    hdr_row(ws1, 2, ["Q","PKT","POSS","2PM","2PA","2P%","3PM","3PA","3P%","FTM","FTA","BR","FD"])
    for team, label, col_offset in [(gtk_name,"gtk",0),(name_opp,"opp",14)]:
        for qn in [1,2,3,4]:
            qd = next((dict(r) for r in all_stats if r["druzyna"]==team[0:3].lower() and r["kwarta"]==qn), {})
            r = 3+qn-1+(col_offset//14)*5
            # simplified — just write gtk stats
        pass

    # Zawodnicy
    for team_label, druzyna in [(gtk_name,"gtk"),(name_opp,"opp")]:
        ws = wb.create_sheet(f"ZAW {team_label[:8]}")
        hdr_row(ws, 1, ["#","PTS","2PM","2PA","3PM","3PA","FTM","FTA","AST","OREB","DREB","BR","FIN"])
        players = [r for r in all_players if r["druzyna"]==druzyna]
        for i,p in enumerate(sorted(players, key=lambda x: x["pts"], reverse=True)):
            r=2+i
            for j,v in enumerate([p["nr"],p["pts"],p["p2m"],p["p2a"],p["p3m"],p["p3a"],p["ftm"],p["fta"],p["ast"],p["oreb"],p["dreb"],p["br"],p["finishes"]]):
                c=ws.cell(r,j+1,v); c.alignment=CTR; c.border=BD

    # Shot timing
    wst = wb.create_sheet("SHOT TIMING")
    hdr_row(wst, 1, ["Czas",f"{gtk_name} 2PT",f"{gtk_name} 3PT",f"{gtk_name} Eff%",f"{name_opp} 2PT",f"{name_opp} 3PT",f"{name_opp} Eff%"])
    for i,b in enumerate(BUCKETS):
        gd=next((r for r in all_timing if r["druzyna"]=="gtk" and r["bucket"]==b),{})
        od=next((r for r in all_timing if r["druzyna"]=="opp" and r["bucket"]==b),{})
        gm=gd.get("made2",0)+gd.get("made3",0); ga=gd.get("att2",0)+gd.get("att3",0)
        om=od.get("made2",0)+od.get("made3",0); oa=od.get("att2",0)+od.get("att3",0)
        row=[b,f"{gd.get('made2',0)}/{gd.get('att2',0)}",f"{gd.get('made3',0)}/{gd.get('att3',0)}",
             f"{gm/ga:.0%}" if ga else "-",
             f"{od.get('made2',0)}/{od.get('att2',0)}",f"{od.get('made3',0)}/{od.get('att3',0)}",
             f"{om/oa:.0%}" if oa else "-"]
        for j,v in enumerate(row): wst.cell(2+i,j+1,v).alignment=CTR

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"raport_{gtk_name}_vs_{name_opp}.xlsx".replace(" ","_"),
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route("/mecz/<int:match_id>/export/pdf")
def export_match_pdf(match_id):
    return redirect(url_for("mecz", match_id=match_id))


# ══════════════════════════════════════════════════════════════════════════════
# SZABLONY (identyczne jak w app_v2)
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/template/zapis")
def template_zapis():
    wb = openpyxl.Workbook()
    HDR=PatternFill("solid",fgColor="1A2B4A"); HDR_F=Font(color="FFFFFF",bold=True,size=10)
    YEL=PatternFill("solid",fgColor="FFF9C4"); GRN=PatternFill("solid",fgColor="E8F5E9")
    CTR=Alignment(horizontal="center",vertical="center",wrap_text=True)
    BORDER=Border(bottom=Side(style="thin",color="CCCCCC"),right=Side(style="thin",color="CCCCCC"),
                  left=Side(style="thin",color="CCCCCC"),top=Side(style="thin",color="CCCCCC"))

    KODY=[("2","Celny rzut za 2"),("0/2","Niecelny rzut za 2"),("3","Celny rzut za 3"),
          ("0/3","Niecelny rzut za 3"),("BR","Strata"),("P","Przewinienie"),("F","Faul"),
          ("2+1","2pkt + RW"),("2D","Tip-in celny"),("0/2D","Tip-in niecelny"),
          ("1/2W","1/2 RW"),("2/2W","2/2 RW"),("0/2W","0/2 RW"),("3+1","3pkt + RW"),
          ("1/3W","1/3 RW"),("2/3W","2/3 RW"),("3/3W","3/3 RW"),("0/3W","0/3 RW")]

    ws_k=wb.active; ws_k.title="KODY"
    ws_k.merge_cells("A1:B1"); t=ws_k["A1"]; t.value="KODY AKCJI"; t.fill=HDR; t.font=Font(color="FFFFFF",bold=True,size=12); t.alignment=CTR
    for col,lbl in [("A","KOD"),("B","OPIS")]:
        c=ws_k[f"{col}2"]; c.value=lbl; c.fill=HDR; c.font=HDR_F; c.alignment=CTR
    ws_k.column_dimensions["A"].width=12; ws_k.column_dimensions["B"].width=32
    for i,(k,o) in enumerate(KODY):
        r=3+i
        c1=ws_k.cell(r,1,k); c1.fill=YEL; c1.alignment=CTR; c1.border=BORDER; c1.font=Font(bold=True)
        c2=ws_k.cell(r,2,o); c2.border=BORDER
        if i%2==0: c2.fill=PatternFill("solid",fgColor="FAFAFA")

    COLS=[("A","Kwarta",7),("B","Czas",9),("C","Kod",12),("D","Strefa",8),
          ("E","Zaw 1",11),("F","Zaw 2",11),("G","Zaw 3",11),("H","Zaw 4",11),("I","Zaw 5",11),
          ("J","Timeout",8),("K","Kończący",11),("L","Asysta",9),("M","OREB",9),("N","DREB",9)]

    for team in ["drużyna_A","drużyna_B"]:
        ws=wb.create_sheet(team)
        ws.merge_cells("A1:N1"); t=ws["A1"]; t.value=f"ZAPIS MECZU — {team}"; t.fill=HDR; t.font=Font(color="FFFFFF",bold=True,size=11); t.alignment=CTR
        for col,lbl,w in COLS:
            ws.column_dimensions[col].width=w; c=ws[f"{col}2"]; c.value=lbl; c.fill=HDR; c.font=HDR_F; c.alignment=CTR; c.border=BORDER
        ws.row_dimensions[2].height=28
        for r in range(3,53):
            for col,_,_ in COLS:
                c=ws[f"{col}{r}"]; c.border=BORDER; c.alignment=Alignment(horizontal="center",vertical="center")
                if col in ["A","B","C","D"]: c.fill=YEL
                elif col in ["K","L","M","N"]: c.fill=GRN
        ws.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="ZAPIS_MECZU_szablon.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route("/template/szablon")
def template_szablon():
    wb = openpyxl.Workbook()
    HDR=PatternFill("solid",fgColor="1A2B4A"); HDR_F=Font(color="FFFFFF",bold=True,size=10)
    YEL=PatternFill("solid",fgColor="FFF8E1"); CTR=Alignment(horizontal="center",vertical="center")
    BORDER=Border(bottom=Side(style="thin",color="CCCCCC"),right=Side(style="thin",color="CCCCCC"),
                  left=Side(style="thin",color="CCCCCC"),top=Side(style="thin",color="CCCCCC"))

    def hdr(ws,row,col,val,w=10):
        c=ws.cell(row,col,val); c.fill=HDR; c.font=HDR_F; c.alignment=CTR; c.border=BORDER
        ws.column_dimensions[get_column_letter(col)].width=w

    sheets=[
        ("TEAM GENERAL",["Q","PKT","POSS","2PM","2PA","2P%","3PM","3PA","3P%","FTM","FTA","FT%","eFG%","TS%","ORtg","DRtg","NetRtg","PPP","TO%","FT Rate"]),
        ("PLAYERS",["#","MIN","PTS","2PM","2PA","2P%","3PM","3PA","3P%","FTM","FTA","FT%","eFG%","TS%","AST","OREB","DREB","BR","FD","FIN"]),
        ("LINEUPS",["Skład","POSS","PKT","PPP","eFG%","ORtg","DRtg","NetRtg","BR","FD"]),
        ("SHOT TIMING",["Czas","2PT Made","2PT Att","2PT%","3PT Made","3PT Att","3PT%","Eff% łącznie"]),
    ]
    ws_first=None
    for sheet_name, cols in sheets:
        if ws_first is None:
            ws=wb.active; ws.title=sheet_name; ws_first=ws
        else:
            ws=wb.create_sheet(sheet_name)
        ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
        t=ws["A1"]; t.value=sheet_name; t.fill=HDR; t.font=Font(color="FFFFFF",bold=True,size=12); t.alignment=CTR
        for i,c in enumerate(cols): hdr(ws,2,i+1,c)
        for r in range(3,18):
            for col in range(1,len(cols)+1):
                c=ws.cell(r,col); c.fill=YEL; c.alignment=CTR; c.border=BORDER
        ws.freeze_panes="B3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="SZABLON_MECZ.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == "__main__":
    app.run(debug=True)
