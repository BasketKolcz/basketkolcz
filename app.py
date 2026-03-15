from flask import Flask, render_template_string, request, redirect, url_for, flash
import openpyxl
from collections import defaultdict
import os, json

app = Flask(__name__)
app.secret_key = "basketkolcz2025"

# ── Parser ────────────────────────────────────────────────────────────────────

ACTION_2PM   = {"2","2+1","2+0","2D","2D+1"}
ACTION_2PA   = {"2","2+1","2+0","2D","2D+1","0/2","0/2D"}
ACTION_3PM   = {"3","3+1","3+0"}
ACTION_3PA   = {"3","3+1","3+0","0/3"}
ACTION_FTM   = set()  # obliczane z kodu
ACTION_BR    = {"BR"}
ACTION_P     = {"P"}
ACTION_F     = {"F"}
ACTION_TIP   = {"2D","0/2D","2D+1"}

def extract_ft(code):
    """Zwraca (made, attempted) dla rzutów wolnych z kodu."""
    import re
    m = re.match(r'^(\d+)/(\d+)W', code)
    if m:
        return int(m.group(1)), int(m.group(2))
    if re.match(r'^\d+D\d+/\d+W', code):
        parts = re.findall(r'(\d+)/(\d+)W', code)
        if parts:
            return int(parts[0][0]), int(parts[0][1])
    return 0, 0

def parse_team_sheet(ws):
    """Parsuje jeden arkusz drużyny i zwraca słownik statystyk."""
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
        "timing": defaultdict(lambda: defaultdict(lambda: {"made":0,"miss":0})),
        "total_finishes": 0,
    }

    current_q = 1
    current_lineup = []
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    for row in rows:
        if not any(v is not None for v in row[:11]):
            continue

        # Kwarta
        if row[0] is not None:
            try:
                current_q = int(str(row[0]).replace("*","").strip())
            except:
                current_q = 1

        # Zawodnicy na boisku
        for i in range(4, 9):
            if row[i] is not None:
                val = str(row[i]).strip()
                if ";" in val:
                    parts = val.split(";")
                    old, new = parts[0].strip(), parts[1].strip()
                    try:
                        idx = current_lineup.index(int(old))
                        current_lineup[idx] = int(new)
                    except (ValueError, IndexError):
                        try: current_lineup.append(int(new))
                        except: pass
                else:
                    try:
                        p = int(val)
                        if p not in current_lineup:
                            current_lineup.append(p)
                    except: pass

        if len(current_lineup) > 5:
            current_lineup = current_lineup[-5:]

        lineup_key = ";".join(str(p) for p in sorted(current_lineup))

        # Rozbij akcje po średniku
        raw_b = str(row[1]) if row[1] is not None else ""
        raw_c = str(row[2]) if row[2] is not None else ""
        raw_d = str(row[3]) if row[3] is not None else ""
        raw_k = str(row[10]) if len(row) > 10 and row[10] is not None else ""
        raw_l = str(row[11]) if len(row) > 11 and row[11] is not None else ""  # asysta
        raw_m = str(row[12]) if len(row) > 12 and row[12] is not None else ""  # OREB
        raw_n = str(row[13]) if len(row) > 13 and row[13] is not None else ""  # DREB

        codes   = [c.strip() for c in raw_c.split(";") if c.strip()]
        times   = [t.strip() for t in raw_b.split(",") if t.strip()]
        zones   = [z.strip() for z in raw_d.split(";") if z.strip()]
        finishers = [f.strip() for f in raw_k.split(";") if f.strip()]
        assists   = [a.strip() for a in raw_l.split(";") if a.strip()]
        orebs     = [o.strip() for o in raw_m.split(";") if o.strip()]
        drebs     = [d.strip() for d in raw_n.split(";") if d.strip()]

        q = stats["quarter"][current_q]
        stats["quarter"][current_q]["poss"] += 1
        stats["lineups"][lineup_key]["poss"] += 1

        for ai, code in enumerate(codes):
            q["acts"] += 1
            stats["lineups"][lineup_key]["acts"] += 1

            # Czas akcji → timing
            t_val = 0
            if ai < len(times):
                try: t_val = float(times[ai])
                except: pass
            stats["lineups"][lineup_key]["tempo"] += t_val

            bucket = "0s"
            if   t_val == 0:           bucket = "0s"
            elif t_val <= 4:           bucket = "1-4s"
            elif t_val <= 8:           bucket = "5-8s"
            elif t_val <= 12:          bucket = "9-12s"
            elif t_val <= 16:          bucket = "13-16s"
            elif t_val <= 20:          bucket = "17-20s"
            else:                      bucket = "21-24s"

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

            # OREB
            orebler = None
            if ai < len(orebs):
                try: orebler = int(orebs[ai])
                except: pass

            # DREB
            drebler = None
            if ai < len(drebs):
                try: drebler = int(drebs[ai])
                except: pass

            # Klasyfikacja kodu
            pts = 0
            is_shot = False

            if code in ACTION_2PM:
                q["2pm"] += 1; q["2pa"] += 1; pts = 2
                stats["lineups"][lineup_key]["2pm"] += 1
                stats["lineups"][lineup_key]["2pa"] += 1
                is_shot = True
                if zone: stats["zones"][zone]["made"] += 1
                if code in ACTION_TIP:
                    q["tip_made"] += 1
                if code in ("2+1","2D+1"):
                    q["and1_2"] += 1
                if finisher:
                    stats["players"][finisher]["2pm"] += 1
                    stats["players"][finisher]["2pa"] += 1
                    if assister:
                        stats["players"][assister]["ast"] += 1

            elif code in ("0/2","0/2D"):
                q["2pa"] += 1
                stats["lineups"][lineup_key]["2pa"] += 1
                is_shot = True
                if zone: stats["zones"][zone]["miss"] += 1
                if code == "0/2D": q["tip_miss"] += 1
                if finisher:
                    stats["players"][finisher]["2pa"] += 1
                if orebler:
                    stats["players"][orebler]["oreb"] += 1

            elif code in ACTION_3PM:
                q["3pm"] += 1; q["3pa"] += 1; pts = 3
                stats["lineups"][lineup_key]["3pm"] += 1
                stats["lineups"][lineup_key]["3pa"] += 1
                is_shot = True
                if zone: stats["zones"][zone]["made"] += 1
                if code in ("3+1",): q["and1_3"] += 1
                if finisher:
                    stats["players"][finisher]["3pm"] += 1
                    stats["players"][finisher]["3pa"] += 1
                    if assister:
                        stats["players"][assister]["ast"] += 1

            elif code == "0/3":
                q["3pa"] += 1
                stats["lineups"][lineup_key]["3pa"] += 1
                is_shot = True
                if zone: stats["zones"][zone]["miss"] += 1
                if finisher:
                    stats["players"][finisher]["3pa"] += 1
                if orebler:
                    stats["players"][orebler]["oreb"] += 1

            elif code in ACTION_BR:
                q["br"] += 1
                stats["lineups"][lineup_key]["br"] += 1
                if finisher:
                    stats["players"][finisher]["br"] += 1
                if drebler:
                    stats["players"][drebler]["dreb"] += 1

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

            # Faule wymuszone
            if fta > 0 or code in ("2+1","2+0","3+1","3+0","2D+1","F"):
                q["fd"] += 1 if fta > 0 else 0
                if finisher:
                    stats["players"][finisher]["fd"] += 1

            # Punkty
            q["pts"] += pts
            stats["lineups"][lineup_key]["pts"] += pts

            # Timing
            if is_shot:
                shot_type = "3PT" if code in ACTION_3PM | {"0/3"} else "2PT"
                if code in ACTION_3PM:
                    stats["timing"][bucket]["3PT"]["made"] += 1
                elif code == "0/3":
                    stats["timing"][bucket]["3PT"]["miss"] += 1
                elif code in ACTION_2PM:
                    stats["timing"][bucket]["2PT"]["made"] += 1
                elif code in ("0/2","0/2D"):
                    stats["timing"][bucket]["2PT"]["miss"] += 1

            # Wykończenia
            if finisher and code not in ACTION_P | ACTION_BR | ACTION_F:
                stats["players"][finisher]["finishes"] += 1
                stats["total_finishes"] += 1

    # Oblicz SUMA
    total = stats["quarter"]["SUMA"]
    for q_num in [1,2,3,4]:
        qd = stats["quarter"][q_num]
        for k in total:
            total[k] = total.get(k,0) + qd.get(k,0)

    return stats

def calc_kpi(d):
    """Oblicz metryki pochodne."""
    fga = d.get("2pa",0) + d.get("3pa",0)
    fgm = d.get("2pm",0) + d.get("3pm",0)
    pts = d.get("pts",0)
    poss = d.get("poss",1) or 1
    fta = d.get("fta",0)

    efg = (d.get("2pm",0) + 1.5*d.get("3pm",0)) / fga if fga else None
    ts  = pts / (2*(fga + 0.44*fta)) if (fga+fta) else None
    ortg = pts*100/poss if poss else None
    topct = d.get("br",0)/poss if poss else None
    ftr = fta/fga if fga else None
    ppp = pts/poss if poss else None
    fg_pct = fgm/fga if fga else None
    ft_pct = d.get("ftm",0)/fta if fta else None
    p2_pct = d.get("2pm",0)/d.get("2pa",1) if d.get("2pa") else None
    p3_pct = d.get("3pm",0)/d.get("3pa",1) if d.get("3pa") else None

    return {
        "efg":   f"{efg:.1%}" if efg is not None else "-",
        "ts":    f"{ts:.1%}"  if ts  is not None else "-",
        "ortg":  f"{ortg:.1f}" if ortg is not None else "-",
        "topct": f"{topct:.1%}" if topct is not None else "-",
        "ftr":   f"{ftr:.2f}"  if ftr  is not None else "-",
        "ppp":   f"{ppp:.2f}"  if ppp  is not None else "-",
        "fg_pct":f"{fg_pct:.1%}" if fg_pct is not None else "-",
        "ft_pct":f"{ft_pct:.1%}" if ft_pct is not None else "-",
        "p2_pct":f"{p2_pct:.1%}" if p2_pct is not None else "-",
        "p3_pct":f"{p3_pct:.1%}" if p3_pct is not None else "-",
    }

# ── HTML templates ─────────────────────────────────────────────────────────────

BASE_HTML = """<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Basket Kołcz — Analiza</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
  body{background:#f8f9fa;font-family:'Segoe UI',Arial,sans-serif}
  .navbar-brand{font-weight:700;font-size:1.3rem;letter-spacing:.5px}
  .navbar{background:#1a2b4a!important}
  .stat-card{background:#fff;border-radius:12px;border:1px solid #e9ecef;padding:1.25rem;text-align:center;margin-bottom:1rem}
  .stat-value{font-size:2rem;font-weight:700;color:#1a2b4a;line-height:1}
  .stat-label{font-size:.75rem;color:#888;text-transform:uppercase;letter-spacing:.5px;margin-top:.3rem}
  .kpi-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:.8rem;font-weight:600}
  .kpi-good{background:#e8f5e9;color:#1a6b3c}
  .kpi-mid{background:#fff8e1;color:#7d5a00}
  .kpi-low{background:#ffebee;color:#8b1a1a}
  .section-title{font-size:.7rem;text-transform:uppercase;letter-spacing:1px;color:#999;font-weight:600;margin-bottom:.75rem;padding-bottom:.4rem;border-bottom:1px solid #f0f0f0}
  .table th{background:#1a2b4a;color:#fff;font-size:.8rem;font-weight:600;border:none}
  .table td{font-size:.85rem;vertical-align:middle}
  .table-hover tbody tr:hover{background:#f0f4ff}
  .upload-zone{border:2px dashed #1a2b4a;border-radius:16px;padding:3rem 2rem;text-align:center;background:#fff;cursor:pointer;transition:.2s}
  .upload-zone:hover{background:#f0f4ff;border-color:#378add}
  .quarter-pill{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.75rem;font-weight:600;margin-right:4px}
  .q1{background:#e8f5e9;color:#1a6b3c}
  .q2{background:#e3f2fd;color:#0c447c}
  .q3{background:#fff3e0;color:#854f0b}
  .q4{background:#fce4ec;color:#8b1a1a}
  .qs{background:#1a2b4a;color:#fff}
  .hero{background:linear-gradient(135deg,#1a2b4a,#2e5090);color:#fff;border-radius:16px;padding:2rem;margin-bottom:2rem}
  .progress{height:6px;border-radius:3px}
  .tab-content{padding-top:1.5rem}
  .nav-tabs .nav-link.active{font-weight:600;color:#1a2b4a;border-bottom:2px solid #1a2b4a}
  @media(max-width:576px){.stat-value{font-size:1.5rem}}
</style>
</head>
<body>
<nav class="navbar navbar-dark mb-4">
  <div class="container">
    <a class="navbar-brand text-white" href="/">
      <span style="color:#EF9F27">&#9679;</span> Basket Kołcz
    </a>
    <span class="text-white-50" style="font-size:.85rem">Analiza Koszykówki</span>
  </div>
</nav>
<div class="container pb-5">
{% block content %}{% endblock %}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
{% block scripts %}{% endblock %}
</body>
</html>"""

INDEX_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
<div class="hero">
  <h1 class="fw-bold mb-1" style="font-size:1.8rem">Witaj w Basket Kołcz Stats</h1>
  <p class="mb-0 opacity-75">Wgraj plik zapis.xlsx i otrzymaj pełną analizę meczu w kilka sekund</p>
</div>

{% with messages = get_flashed_messages(with_categories=true) %}
  {% if messages %}
    {% for cat,msg in messages %}
      <div class="alert alert-{{ 'danger' if cat=='error' else 'info' }} alert-dismissible fade show">
        {{ msg }}<button type="button" class="btn-close" data-bs-dismiss="alert"></button>
      </div>
    {% endfor %}
  {% endif %}
{% endwith %}

<div class="row justify-content-center">
  <div class="col-lg-8">
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <div class="upload-zone" onclick="document.getElementById('file').click()">
        <div style="font-size:3rem;margin-bottom:1rem">📊</div>
        <h4 class="fw-bold" style="color:#1a2b4a">Wgraj plik zapis.xlsx</h4>
        <p class="text-muted mb-2">Kliknij lub przeciągnij plik tutaj</p>
        <p class="text-muted" style="font-size:.85rem">Format: plik z dwoma arkuszami (obie drużyny)</p>
        <input type="file" id="file" name="file" accept=".xlsx" class="d-none" onchange="this.form.submit()">
      </div>
    </form>
    <div class="mt-3 p-3 rounded" style="background:#fff;border:1px solid #e9ecef;font-size:.85rem;color:#666">
      <strong>Jak to działa:</strong>
      Wgraj plik zapis.xlsx z zakodowanym meczem → parser automatycznie wyliczy wszystkie statystyki → zobaczysz pełny raport z metrykami NBA/Euroleague
    </div>
  </div>
</div>
""").replace("{% block scripts %}{% endblock %}","")

REPORT_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <div>
    <h2 class="fw-bold mb-0" style="color:#1a2b4a">{{ meta.team_a }} vs {{ meta.team_b }}</h2>
    <p class="text-muted mb-0" style="font-size:.9rem">Raport meczowy</p>
  </div>
  <a href="/" class="btn btn-outline-secondary btn-sm">← Nowy mecz</a>
</div>

<!-- TABS DRUŻYN -->
<ul class="nav nav-tabs" id="teamTabs">
  <li class="nav-item">
    <button class="nav-link active" data-bs-toggle="tab" data-bs-target="#teamA">{{ meta.team_a }}</button>
  </li>
  <li class="nav-item">
    <button class="nav-link" data-bs-toggle="tab" data-bs-target="#teamB">{{ meta.team_b }}</button>
  </li>
  <li class="nav-item">
    <button class="nav-link" data-bs-toggle="tab" data-bs-target="#compare">Porównanie</button>
  </li>
</ul>

<div class="tab-content">

<!-- ═══ DRUŻYNA A ═══ -->
<div class="tab-pane fade show active" id="teamA">
  {% set s = stats_a %}
  {% set q = s.quarter %}
  {% include 'team_panel.html' %}
</div>

<!-- ═══ DRUŻYNA B ═══ -->
<div class="tab-pane fade" id="teamB">
  {% set s = stats_b %}
  {% set q = s.quarter %}
  {% include 'team_panel.html' %}
</div>

<!-- ═══ PORÓWNANIE ═══ -->
<div class="tab-pane fade" id="compare">
  <div class="row g-3">
    <div class="col-12">
      <div class="stat-card">
        <div class="section-title">Kluczowe metryki — porównanie</div>
        <div class="table-responsive">
          <table class="table table-hover mb-0">
            <thead><tr>
              <th>Metryka</th>
              <th class="text-center">{{ meta.team_a }}</th>
              <th class="text-center">{{ meta.team_b }}</th>
            </tr></thead>
            <tbody>
              {% for label, ka, kb in compare_rows %}
              <tr>
                <td class="fw-500">{{ label }}</td>
                <td class="text-center fw-bold" style="color:#1a6b3c">{{ ka }}</td>
                <td class="text-center fw-bold" style="color:#8b1a1a">{{ kb }}</td>
              </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="col-12">
      <div class="stat-card">
        <div class="section-title">Punkty per kwarta</div>
        <canvas id="quarterChart" height="120"></canvas>
      </div>
    </div>
  </div>
</div>

</div>
""")

TEAM_PANEL = """
{% set suma = s.quarter.get('SUMA', s.quarter.get(4,{})) %}
{% set kpi = kpis[loop_team] %}

<!-- KPI cards -->
<div class="row g-2 my-3">
  <div class="col-6 col-md-3">
    <div class="stat-card">
      <div class="stat-value">{{ suma.get('pts',0) }}</div>
      <div class="stat-label">Punkty</div>
    </div>
  </div>
  <div class="col-6 col-md-3">
    <div class="stat-card">
      <div class="stat-value" style="font-size:1.4rem">{{ kpi.efg }}</div>
      <div class="stat-label">eFG%</div>
    </div>
  </div>
  <div class="col-6 col-md-3">
    <div class="stat-card">
      <div class="stat-value" style="font-size:1.4rem">{{ kpi.ortg }}</div>
      <div class="stat-label">ORtg</div>
    </div>
  </div>
  <div class="col-6 col-md-3">
    <div class="stat-card">
      <div class="stat-value" style="font-size:1.4rem">{{ kpi.ppp }}</div>
      <div class="stat-label">Pkt/Posiadanie</div>
    </div>
  </div>
</div>

<!-- Per kwarta -->
<div class="stat-card mb-3">
  <div class="section-title">Statystyki per kwarta</div>
  <div class="table-responsive">
    <table class="table table-hover mb-0" style="font-size:.82rem">
      <thead><tr>
        <th>Kwarta</th><th>PKT</th><th>2PM/A</th><th>2P%</th>
        <th>3PM/A</th><th>3P%</th><th>FTM/A</th><th>BR</th>
        <th>POSS</th><th>PPP</th><th>eFG%</th>
      </tr></thead>
      <tbody>
        {% for qn in [1,2,3,4,'SUMA'] %}
        {% set qd = s.quarter.get(qn, {}) %}
        {% set qkpi = calc_kpi(qd) %}
        <tr {% if qn=='SUMA' %}style="font-weight:700;background:#f0f4ff"{% endif %}>
          <td>
            <span class="quarter-pill {{ 'q1' if qn==1 else 'q2' if qn==2 else 'q3' if qn==3 else 'q4' if qn==4 else 'qs' }}">
              {{ 'SUMA' if qn=='SUMA' else qn~'Q' }}
            </span>
          </td>
          <td>{{ qd.get('pts',0) }}</td>
          <td>{{ qd.get('2pm',0) }}/{{ qd.get('2pa',0) }}</td>
          <td>{{ qkpi.p2_pct }}</td>
          <td>{{ qd.get('3pm',0) }}/{{ qd.get('3pa',0) }}</td>
          <td>{{ qkpi.p3_pct }}</td>
          <td>{{ qd.get('ftm',0) }}/{{ qd.get('fta',0) }}</td>
          <td>{{ qd.get('br',0) }}</td>
          <td>{{ qd.get('poss',0) }}</td>
          <td>{{ qkpi.ppp }}</td>
          <td>{{ qkpi.efg }}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>

<!-- Zawodnicy -->
<div class="stat-card mb-3">
  <div class="section-title">Statystyki indywidualne</div>
  <div class="table-responsive">
    <table class="table table-hover mb-0" style="font-size:.82rem">
      <thead><tr>
        <th>#</th><th>2PM/A</th><th>3PM/A</th><th>FTM/A</th>
        <th>PTS</th><th>eFG%</th><th>AST</th><th>BR</th><th>FD</th>
      </tr></thead>
      <tbody>
        {% for pid, pd in s.players.items()|sort %}
        {% set player_kpi = calc_kpi(pd) %}
        {% set pts = pd.get('2pm',0)*2 + pd.get('3pm',0)*3 + pd.get('ftm',0) %}
        <tr>
          <td class="fw-bold">#{{ pid }}</td>
          <td>{{ pd.get('2pm',0) }}/{{ pd.get('2pa',0) }}</td>
          <td>{{ pd.get('3pm',0) }}/{{ pd.get('3pa',0) }}</td>
          <td>{{ pd.get('ftm',0) }}/{{ pd.get('fta',0) }}</td>
          <td class="fw-bold">{{ pts }}</td>
          <td>{{ player_kpi.efg }}</td>
          <td>{{ pd.get('ast',0) }}</td>
          <td>{{ pd.get('br',0) }}</td>
          <td>{{ pd.get('fd',0) }}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>

<!-- Top składy -->
<div class="stat-card">
  <div class="section-title">Najlepsze składy 5-osobowe (ofensywa)</div>
  <div class="table-responsive">
    <table class="table table-hover mb-0" style="font-size:.82rem">
      <thead><tr>
        <th>Skład</th><th>POSS</th><th>PKT</th><th>PPP</th><th>eFG%</th><th>BR</th>
      </tr></thead>
      <tbody>
        {% for lid, ld in s.lineups.items()|sort(attribute='1.pts',reverse=True) %}
        {% if ld.poss >= 2 %}
        {% set lkpi = calc_kpi(ld) %}
        <tr>
          <td style="font-family:monospace;font-size:.78rem">{{ lid }}</td>
          <td>{{ ld.poss }}</td>
          <td class="fw-bold">{{ ld.pts }}</td>
          <td>{{ lkpi.ppp }}</td>
          <td>{{ lkpi.efg }}</td>
          <td>{{ ld.br }}</td>
        </tr>
        {% endif %}
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
"""

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(INDEX_HTML)

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("Nie wybrano pliku", "error")
        return redirect(url_for("index"))

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        flash("Plik musi być w formacie .xlsx", "error")
        return redirect(url_for("index"))

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        data_sheets = [s for s in wb.sheetnames if s.upper() not in ("META","KODY")]

        if len(data_sheets) < 1:
            flash("Plik musi zawierać co najmniej jeden arkusz z danymi", "error")
            return redirect(url_for("index"))

        team_a = data_sheets[0]
        team_b = data_sheets[1] if len(data_sheets) > 1 else data_sheets[0]

        stats_a = parse_team_sheet(wb[team_a])
        stats_b = parse_team_sheet(wb[team_b])

        # KPI dla obu drużyn
        suma_a = stats_a["quarter"].get("SUMA", {})
        suma_b = stats_b["quarter"].get("SUMA", {})
        kpis = {
            "a": calc_kpi(suma_a),
            "b": calc_kpi(suma_b),
        }

        # Tabela porównawcza
        compare_rows = [
            ("Punkty",            suma_a.get("pts",0),      suma_b.get("pts",0)),
            ("Posiadania",        suma_a.get("poss",0),     suma_b.get("poss",0)),
            ("eFG%",              kpis["a"]["efg"],          kpis["b"]["efg"]),
            ("True Shooting%",    kpis["a"]["ts"],           kpis["b"]["ts"]),
            ("Offensive Rating",  kpis["a"]["ortg"],         kpis["b"]["ortg"]),
            ("Pkt/Posiadanie",    kpis["a"]["ppp"],          kpis["b"]["ppp"]),
            ("Turnover%",         kpis["a"]["topct"],        kpis["b"]["topct"]),
            ("FT Rate",           kpis["a"]["ftr"],          kpis["b"]["ftr"]),
            ("2PT%",              kpis["a"]["p2_pct"],       kpis["b"]["p2_pct"]),
            ("3PT%",              kpis["a"]["p3_pct"],       kpis["b"]["p3_pct"]),
            ("Straty (BR)",       suma_a.get("br",0),        suma_b.get("br",0)),
        ]

        meta = {"team_a": team_a, "team_b": team_b}

        # Punkty per kwarta (do wykresu)
        pts_a = [stats_a["quarter"].get(q,{}).get("pts",0) for q in [1,2,3,4]]
        pts_b = [stats_b["quarter"].get(q,{}).get("pts",0) for q in [1,2,3,4]]

        return render_template_string(
            build_report_html(),
            meta=meta,
            stats_a=stats_a,
            stats_b=stats_b,
            kpis=kpis,
            compare_rows=compare_rows,
            pts_a=pts_a,
            pts_b=pts_b,
            calc_kpi=calc_kpi,
        )

    except Exception as e:
        flash(f"Błąd podczas parsowania pliku: {str(e)}", "error")
        return redirect(url_for("index"))


def build_report_html():
    """Buduje pełny HTML raportu."""

    TEAM_TAB = """
<div class="tab-pane fade {{ 'show active' if loop_team == 'a' else '' }}" id="team{{ loop_team|upper }}">
{% set s = stats_a if loop_team == 'a' else stats_b %}
{% set kpi = kpis[loop_team] %}
{% set suma = {} %}
{% for qn in [1,2,3,4] %}
  {% set _ = suma.update({k: suma.get(k,0) + s.quarter.get(qn,{}).get(k,0) for k in ['pts','2pm','2pa','3pm','3pa','ftm','fta','br','poss','acts','fd','p']}) %}
{% endfor %}

<div class="row g-2 my-3">
  {% for val, lbl in [
    (s.quarter.get(1,{}).get('pts',0)+s.quarter.get(2,{}).get('pts',0)+s.quarter.get(3,{}).get('pts',0)+s.quarter.get(4,{}).get('pts',0), 'Punkty'),
    (kpi.efg, 'eFG%'),
    (kpi.ortg, 'ORtg'),
    (kpi.ppp, 'Pkt/Pos'),
    (kpi.p2_pct, '2PT%'),
    (kpi.p3_pct, '3PT%'),
  ] %}
  <div class="col-6 col-md-2">
    <div class="stat-card">
      <div class="stat-value" style="font-size:1.4rem">{{ val }}</div>
      <div class="stat-label">{{ lbl }}</div>
    </div>
  </div>
  {% endfor %}
</div>
</div>
"""

    report = BASE_HTML.replace("{% block content %}{% endblock %}", """
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <div>
    <h2 class="fw-bold mb-0" style="color:#1a2b4a">{{ meta.team_a }} vs {{ meta.team_b }}</h2>
    <p class="text-muted mb-0" style="font-size:.9rem">Raport meczowy</p>
  </div>
  <a href="/" class="btn btn-outline-secondary btn-sm">← Nowy mecz</a>
</div>

<ul class="nav nav-tabs mb-3" id="mainTabs">
  <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tabA">{{ meta.team_a }}</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabB">{{ meta.team_b }}</button></li>
  <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tabCmp">Porównanie</button></li>
</ul>

<div class="tab-content">

{% for loop_team, s_data in [('a', stats_a), ('b', stats_b)] %}
<div class="tab-pane fade {% if loop.first %}show active{% endif %}" id="tab{{ 'A' if loop.first else 'B' }}">
  {% set s = s_data %}
  {% set kpi = kpis[loop_team] %}

  <!-- KPI row -->
  <div class="row g-2 my-2">
    {% set pts_total = s.quarter.get(1,{}).get('pts',0)+s.quarter.get(2,{}).get('pts',0)+s.quarter.get(3,{}).get('pts',0)+s.quarter.get(4,{}).get('pts',0) %}
    {% for val, lbl in [(pts_total,'Punkty'),(kpi.efg,'eFG%'),(kpi.ts,'TS%'),(kpi.ortg,'ORtg'),(kpi.ppp,'Pkt/Pos'),(kpi.topct,'TO%')] %}
    <div class="col-6 col-md-2">
      <div class="stat-card">
        <div class="stat-value" style="font-size:1.3rem">{{ val }}</div>
        <div class="stat-label">{{ lbl }}</div>
      </div>
    </div>
    {% endfor %}
  </div>

  <!-- Per kwarta -->
  <div class="stat-card mb-3">
    <div class="section-title">Per kwarta</div>
    <div class="table-responsive">
      <table class="table table-sm table-hover mb-0">
        <thead><tr><th>Q</th><th>PKT</th><th>2PM/A</th><th>2P%</th><th>3PM/A</th><th>3P%</th><th>FT</th><th>BR</th><th>POSS</th><th>PPP</th><th>eFG%</th></tr></thead>
        <tbody>
          {% for qn in [1,2,3,4] %}
          {% set qd = s.quarter.get(qn, {}) %}
          {% set qk = calc_kpi(qd) %}
          <tr>
            <td><span class="quarter-pill q{{ qn }}">{{ qn }}Q</span></td>
            <td>{{ qd.get('pts',0) }}</td>
            <td>{{ qd.get('2pm',0) }}/{{ qd.get('2pa',0) }}</td>
            <td>{{ qk.p2_pct }}</td>
            <td>{{ qd.get('3pm',0) }}/{{ qd.get('3pa',0) }}</td>
            <td>{{ qk.p3_pct }}</td>
            <td>{{ qd.get('ftm',0) }}/{{ qd.get('fta',0) }}</td>
            <td>{{ qd.get('br',0) }}</td>
            <td>{{ qd.get('poss',0) }}</td>
            <td>{{ qk.ppp }}</td>
            <td>{{ qk.efg }}</td>
          </tr>
          {% endfor %}
          {% set sd = {'pts':0,'2pm':0,'2pa':0,'3pm':0,'3pa':0,'ftm':0,'fta':0,'br':0,'poss':0} %}
          {% for qn in [1,2,3,4] %}{% set qd=s.quarter.get(qn,{}) %}
            {% for k in sd %}{% set _=sd.update({k:sd[k]+qd.get(k,0)})%}{% endfor %}
          {% endfor %}
          {% set sk = calc_kpi(sd) %}
          <tr style="font-weight:700;background:#f0f4ff">
            <td><span class="quarter-pill qs">SUMA</span></td>
            <td>{{ sd.pts }}</td>
            <td>{{ sd['2pm'] }}/{{ sd['2pa'] }}</td><td>{{ sk.p2_pct }}</td>
            <td>{{ sd['3pm'] }}/{{ sd['3pa'] }}</td><td>{{ sk.p3_pct }}</td>
            <td>{{ sd.ftm }}/{{ sd.fta }}</td>
            <td>{{ sd.br }}</td><td>{{ sd.poss }}</td>
            <td>{{ sk.ppp }}</td><td>{{ sk.efg }}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>

  <!-- Zawodnicy -->
  <div class="stat-card mb-3">
    <div class="section-title">Zawodnicy</div>
    <div class="table-responsive">
      <table class="table table-sm table-hover mb-0">
        <thead><tr><th>#</th><th>PTS</th><th>2PM/A</th><th>3PM/A</th><th>FTM/A</th><th>eFG%</th><th>AST</th><th>OREB</th><th>DREB</th><th>BR</th><th>FD</th></tr></thead>
        <tbody>
          {% for pid, pd in s.players.items()|sort %}
          {% set pk = calc_kpi(pd) %}
          {% set pts = pd.get('2pm',0)*2+pd.get('3pm',0)*3+pd.get('ftm',0) %}
          <tr>
            <td class="fw-bold">#{{ pid }}</td>
            <td class="fw-bold">{{ pts }}</td>
            <td>{{ pd.get('2pm',0) }}/{{ pd.get('2pa',0) }}</td>
            <td>{{ pd.get('3pm',0) }}/{{ pd.get('3pa',0) }}</td>
            <td>{{ pd.get('ftm',0) }}/{{ pd.get('fta',0) }}</td>
            <td>{{ pk.efg }}</td>
            <td>{{ pd.get('ast',0) }}</td>
            <td>{{ pd.get('oreb',0) }}</td>
            <td>{{ pd.get('dreb',0) }}</td>
            <td>{{ pd.get('br',0) }}</td>
            <td>{{ pd.get('fd',0) }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

  <!-- Składy -->
  <div class="stat-card">
    <div class="section-title">Składy 5-osobowe (min. 2 posiadania)</div>
    <div class="table-responsive">
      <table class="table table-sm table-hover mb-0">
        <thead><tr><th>Skład</th><th>POSS</th><th>PKT</th><th>PPP</th><th>eFG%</th><th>ORtg</th><th>BR</th></tr></thead>
        <tbody>
          {% for lid, ld in s.lineups.items()|sort(attribute='1.pts',reverse=True) %}
          {% if ld.poss >= 2 %}
          {% set lk = calc_kpi(ld) %}
          <tr>
            <td style="font-family:monospace;font-size:.78rem">{{ lid }}</td>
            <td>{{ ld.poss }}</td>
            <td class="fw-bold">{{ ld.pts }}</td>
            <td>{{ lk.ppp }}</td>
            <td>{{ lk.efg }}</td>
            <td>{{ lk.ortg }}</td>
            <td>{{ ld.br }}</td>
          </tr>
          {% endif %}
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

</div>
{% endfor %}

<!-- Porównanie -->
<div class="tab-pane fade" id="tabCmp">
  <div class="row g-3 mt-1">
    <div class="col-lg-6">
      <div class="stat-card">
        <div class="section-title">Kluczowe metryki</div>
        <table class="table table-sm table-hover mb-0">
          <thead><tr><th>Metryka</th><th class="text-center" style="color:#1a6b3c">{{ meta.team_a }}</th><th class="text-center" style="color:#8b1a1a">{{ meta.team_b }}</th></tr></thead>
          <tbody>
            {% for label, va, vb in compare_rows %}
            <tr><td>{{ label }}</td>
              <td class="text-center fw-bold" style="color:#1a6b3c">{{ va }}</td>
              <td class="text-center fw-bold" style="color:#8b1a1a">{{ vb }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    <div class="col-lg-6">
      <div class="stat-card">
        <div class="section-title">Punkty per kwarta</div>
        <canvas id="quarterChart"></canvas>
      </div>
    </div>
  </div>
</div>

</div>
""").replace("{% block scripts %}{% endblock %}", """
<script>
const ctx = document.getElementById('quarterChart');
if(ctx){
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['1Q','2Q','3Q','4Q'],
      datasets: [
        {label:'{{ meta.team_a }}', data:{{ pts_a }}, backgroundColor:'#1a6b3c88', borderColor:'#1a6b3c', borderWidth:2, borderRadius:6},
        {label:'{{ meta.team_b }}', data:{{ pts_b }}, backgroundColor:'#8b1a1a88', borderColor:'#8b1a1a', borderWidth:2, borderRadius:6}
      ]
    },
    options:{responsive:true, plugins:{legend:{position:'top'}}, scales:{y:{beginAtZero:true,grid:{color:'#f0f0f0'}}}}
  });
}
</script>
""")
    return report


if __name__ == "__main__":
    app.run(debug=True)
