#!/usr/bin/env python3
"""
generate_stats_report.py — genererer en flot, selvstændig HTML5-statistikrapport
for det danske backdrop-sideprojekt (danish_backdrops.py).

KUN BACKDROPS (rettet 2026-08-07): "Fundet"-tallet i denne rapport tæller nu
UDELUKKENDE titler med et ægte dansk BACKDROP - postere indgår ikke længere
(se danish_backdrops.py / export_danish_artwork_review.py for baggrund:
UHF viste postere forkert beskåret, da de er portræt-format i en 16:9-ramme).

Læser:
    data/danish_artwork_cache.json          (alle TMDb-opslag: fundet/ikke fundet)
    data/danish_artwork_review.xlsx         (godkendelses-status, hvis den findes)
    data/danish_backdrops_run_log.json      (historik over tidligere kørsler)

Skriver:
    output/danish_backdrops_report.html     (ÉN fil, ingen eksterne afhængigheder -
                                              alle grafer er indlejret SVG, ingen CDN,
                                              virker garanteret offline for evigt)

BRUG
    python3 scripts/generate_stats_report.py
"""
from __future__ import annotations

import json
import math
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

DANISH_ARTWORK_CACHE_FILE = DATA_DIR / "danish_artwork_cache.json"
DANISH_ARTWORK_REVIEW_FILE = DATA_DIR / "danish_artwork_review.xlsx"
DANISH_BACKDROPS_RUN_LOG_FILE = DATA_DIR / "danish_backdrops_run_log.json"
REPORT_FILE = OUTPUT_DIR / "danish_backdrops_report.html"

COLORS = {
    "found": "#22c55e",
    "not_found": "#e2e8f0",
    "approved": "#3b82f6",
    "pending": "#f59e0b",
    "sport": "#a855f7",
    "grid": "#334155",
    "text_muted": "#94a3b8",
}


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return default
    return default


def load_review_status(path: Path) -> tuple[int, int, dict[str, str]]:
    if not path.exists():
        return 0, 0, {}
    try:
        from openpyxl import load_workbook
    except ImportError:
        return 0, 0, {}

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    try:
        key_col = headers.index("Nøgle (intern)")
        godkendt_col = headers.index("Godkendt (X)")
        note_col = headers.index("Note") if "Note" in headers else None
    except ValueError:
        return 0, 0, {}

    approved = 0
    flagged = 0
    notes: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2):
        key_val = row[key_col].value
        if not key_val:
            continue
        godkendt_val = row[godkendt_col].value
        note_val = row[note_col].value if note_col is not None else None
        if godkendt_val and str(godkendt_val).strip().upper() == "X":
            approved += 1
        elif note_val:
            flagged += 1
        if note_val:
            notes[str(key_val).strip()] = str(note_val)
    return approved, flagged, notes


def donut_chart_svg(segments: list[tuple[str, float, str]], size: int = 220, hole_ratio: float = 0.62) -> str:
    total = sum(v for _, v, _ in segments) or 1
    cx = cy = size / 2
    r_outer = size / 2 - 6
    r_inner = r_outer * hole_ratio

    def point(angle_deg, r):
        angle_rad = math.radians(angle_deg - 90)
        return cx + r * math.cos(angle_rad), cy + r * math.sin(angle_rad)

    paths = []
    angle = 0.0
    for label, value, color in segments:
        if value <= 0:
            continue
        sweep = (value / total) * 360
        end_angle = angle + sweep
        large_arc = 1 if sweep > 180 else 0

        x1o, y1o = point(angle, r_outer)
        x2o, y2o = point(end_angle, r_outer)
        x1i, y1i = point(end_angle, r_inner)
        x2i, y2i = point(angle, r_inner)

        path = (
            f'M {x1o:.2f},{y1o:.2f} '
            f'A {r_outer:.2f},{r_outer:.2f} 0 {large_arc} 1 {x2o:.2f},{y2o:.2f} '
            f'L {x1i:.2f},{y1i:.2f} '
            f'A {r_inner:.2f},{r_inner:.2f} 0 {large_arc} 0 {x2i:.2f},{y2i:.2f} Z'
        )
        title = f"{label}: {int(value):,} ({value/total*100:.1f}%)"
        paths.append(f'<path d="{path}" fill="{color}"><title>{title}</title></path>')
        angle = end_angle

    center_text = f'<text x="{cx}" y="{cy-6}" text-anchor="middle" class="donut-total">{int(total):,}</text>' \
                  f'<text x="{cx}" y="{cy+16}" text-anchor="middle" class="donut-sub">i alt</text>'

    return f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}">{"".join(paths)}{center_text}</svg>'


def bar_chart_svg(categories: list[str], series: list[tuple[str, list[float], str]],
                   width: int = 640, height: int = 320, y_label: str = "") -> str:
    pad_left, pad_right, pad_top, pad_bottom = 56, 20, 24, 64
    plot_w = width - pad_left - pad_right
    plot_h = height - pad_top - pad_bottom

    max_val = max((max(vals) if vals else 0) for _, vals, _ in series) or 1
    max_val = max_val * 1.15

    n_cat = len(categories)
    n_series = len(series)
    group_w = plot_w / n_cat if n_cat else plot_w
    bar_w = group_w / (n_series + 1)

    def y_scale(v):
        return pad_top + plot_h - (v / max_val) * plot_h

    svg_parts = []

    grid_lines = 5
    for i in range(grid_lines + 1):
        v = max_val / grid_lines * i
        y = y_scale(v)
        svg_parts.append(
            f'<line x1="{pad_left}" y1="{y:.1f}" x2="{width-pad_right}" y2="{y:.1f}" '
            f'stroke="{COLORS["grid"]}" stroke-width="1" opacity="0.4"/>'
        )
        svg_parts.append(
            f'<text x="{pad_left-10}" y="{y+4:.1f}" text-anchor="end" class="axis-label">{int(v):,}</text>'
        )

    for cat_idx, cat in enumerate(categories):
        group_x = pad_left + cat_idx * group_w
        for s_idx, (s_name, vals, color) in enumerate(series):
            val = vals[cat_idx] if cat_idx < len(vals) else 0
            bar_x = group_x + (s_idx + 0.5) * bar_w
            bar_y = y_scale(val)
            bar_h = pad_top + plot_h - bar_y
            svg_parts.append(
                f'<rect x="{bar_x:.1f}" y="{bar_y:.1f}" width="{bar_w*0.82:.1f}" height="{bar_h:.1f}" '
                f'rx="3" fill="{color}"><title>{cat} — {s_name}: {int(val):,}</title></rect>'
            )
        label_x = group_x + group_w / 2
        svg_parts.append(
            f'<text x="{label_x:.1f}" y="{height-pad_bottom+20}" text-anchor="middle" '
            f'class="axis-label" transform="rotate(-20 {label_x:.1f} {height-pad_bottom+20})">{cat}</text>'
        )

    legend_x = pad_left
    legend_y = height - 18
    for s_name, _, color in series:
        svg_parts.append(f'<rect x="{legend_x}" y="{legend_y-10}" width="12" height="12" rx="2" fill="{color}"/>')
        svg_parts.append(f'<text x="{legend_x+18}" y="{legend_y}" class="legend-label">{s_name}</text>')
        legend_x += 18 + len(s_name) * 7 + 24

    return f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}">{"".join(svg_parts)}</svg>'


def line_chart_svg(x_labels: list[str], series: list[tuple[str, list[float], str]],
                    width: int = 640, height: int = 260) -> str:
    pad_left, pad_right, pad_top, pad_bottom = 56, 20, 20, 50
    plot_w = width - pad_left - pad_right
    plot_h = height - pad_top - pad_bottom

    max_val = max((max(vals) if vals else 0) for _, vals, _ in series) or 1
    max_val = max_val * 1.15
    n = len(x_labels)

    def x_scale(i):
        return pad_left + (i / max(n - 1, 1)) * plot_w

    def y_scale(v):
        return pad_top + plot_h - (v / max_val) * plot_h

    svg_parts = []
    grid_lines = 4
    for i in range(grid_lines + 1):
        v = max_val / grid_lines * i
        y = y_scale(v)
        svg_parts.append(
            f'<line x1="{pad_left}" y1="{y:.1f}" x2="{width-pad_right}" y2="{y:.1f}" '
            f'stroke="{COLORS["grid"]}" stroke-width="1" opacity="0.4"/>'
        )
        svg_parts.append(f'<text x="{pad_left-10}" y="{y+4:.1f}" text-anchor="end" class="axis-label">{int(v):,}</text>')

    step = max(1, n // 8)
    for i, label in enumerate(x_labels):
        if i % step == 0 or i == n - 1:
            x = x_scale(i)
            svg_parts.append(
                f'<text x="{x:.1f}" y="{height-pad_bottom+18}" text-anchor="middle" class="axis-label">{label}</text>'
            )

    for s_name, vals, color in series:
        points = " ".join(f"{x_scale(i):.1f},{y_scale(v):.1f}" for i, v in enumerate(vals))
        area_points = f"{pad_left},{pad_top+plot_h} {points} {x_scale(n-1):.1f},{pad_top+plot_h}"
        svg_parts.append(f'<polygon points="{area_points}" fill="{color}" opacity="0.12"/>')
        svg_parts.append(f'<polyline points="{points}" fill="none" stroke="{color}" stroke-width="2.5"/>')
        for i, v in enumerate(vals):
            svg_parts.append(
                f'<circle cx="{x_scale(i):.1f}" cy="{y_scale(v):.1f}" r="3.5" fill="{color}">'
                f'<title>{x_labels[i]}: {int(v):,}</title></circle>'
            )

    return f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}">{"".join(svg_parts)}</svg>'


CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif;
    background: linear-gradient(160deg, #0f172a 0%, #1e293b 100%);
    color: #e2e8f0;
    padding: 32px 24px 64px;
    min-height: 100vh;
}
.wrap { max-width: 1180px; margin: 0 auto; }
header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 32px; flex-wrap: wrap; gap: 12px;
}
h1 { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; }
h1 span { color: #22c55e; }
.timestamp { color: #94a3b8; font-size: 13px; }
.cards {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 16px; margin-bottom: 32px;
}
.card {
    background: rgba(30, 41, 59, 0.7); border: 1px solid rgba(148,163,184,0.15);
    border-radius: 14px; padding: 18px 20px; backdrop-filter: blur(6px);
}
.card .label { font-size: 12px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px; }
.card .value { font-size: 30px; font-weight: 700; }
.card .value.green { color: #22c55e; }
.card .value.blue { color: #3b82f6; }
.card .value.amber { color: #f59e0b; }
.card .value.purple { color: #a855f7; }
.grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }
@media (max-width: 860px) { .grid2 { grid-template-columns: 1fr; } }
.panel {
    background: rgba(30, 41, 59, 0.7); border: 1px solid rgba(148,163,184,0.15);
    border-radius: 16px; padding: 22px; margin-bottom: 24px;
}
.panel h2 { font-size: 16px; font-weight: 600; margin-bottom: 16px; color: #f1f5f9; }
.panel h2 .sub { font-weight: 400; color: #94a3b8; font-size: 13px; margin-left: 8px; }
.donut-row { display: flex; align-items: center; justify-content: center; gap: 28px; flex-wrap: wrap; }
.donut-legend { display: flex; flex-direction: column; gap: 10px; }
.legend-item { display: flex; align-items: center; gap: 8px; font-size: 14px; }
.legend-dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }
.legend-item b { margin-left: auto; padding-left: 16px; }
svg text.axis-label { fill: #94a3b8; font-size: 10px; }
svg text.legend-label { fill: #cbd5e1; font-size: 12px; }
svg text.donut-total { fill: #f1f5f9; font-size: 26px; font-weight: 700; }
svg text.donut-sub { fill: #94a3b8; font-size: 12px; }
table.notes { width: 100%; border-collapse: collapse; font-size: 13px; }
table.notes th { text-align: left; color: #94a3b8; font-weight: 600; padding: 8px 10px; border-bottom: 1px solid rgba(148,163,184,0.2); }
table.notes td { padding: 8px 10px; border-bottom: 1px solid rgba(148,163,184,0.08); }
.badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 11px; font-weight: 600; }
.badge.approved { background: rgba(34,197,94,0.15); color: #22c55e; }
.badge.pending { background: rgba(245,158,11,0.15); color: #f59e0b; }
.empty-state { color: #94a3b8; font-size: 14px; padding: 20px 0; text-align: center; }
footer { text-align: center; color: #64748b; font-size: 12px; margin-top: 32px; }
"""


def build_html(cache: dict, run_log: list, approved: int, flagged: int, notes: dict) -> str:
    total_titles = len(cache)
    # KUN backdrop tæller som "fundet" - postere ignoreres (se modulets docstring)
    found = sum(1 for v in cache.values() if v.get("backdrop"))
    not_found = total_titles - found
    pending = max(found - approved - flagged, 0)
    hit_rate = (found / total_titles * 100) if total_titles else 0

    donut_found = donut_chart_svg([
        ("Fundet (backdrop)", found, COLORS["found"]),
        ("Ikke fundet", not_found, COLORS["not_found"]),
    ])
    donut_review = donut_chart_svg([
        ("Godkendt", approved, COLORS["approved"]),
        ("Afventer", pending, COLORS["pending"]),
        ("Markeret forkert", flagged, "#ef4444"),
    ]) if found > 0 else '<div class="empty-state">Ingen fund endnu</div>'

    per_file_bar = ""
    if run_log:
        latest = run_log[-1]
        per_file = latest.get("per_file", {})
        categories = list(per_file.keys())
        checked_vals = [per_file[c].get("checked", 0) for c in categories]
        found_vals = [per_file[c].get("danish_found", 0) for c in categories]
        sport_vals = [per_file[c].get("already_had_artwork", 0) for c in categories]
        per_file_bar = bar_chart_svg(
            categories,
            [
                ("Sport (sprunget over)", sport_vals, COLORS["sport"]),
                ("Titler tjekket", checked_vals, "#38bdf8"),
                ("Dansk backdrop fundet", found_vals, COLORS["found"]),
            ],
        )
    else:
        per_file_bar = '<div class="empty-state">Ingen kørselshistorik endnu</div>'

    cache_growth_chart = ""
    if len(run_log) >= 2:
        x_labels = [time.strftime("%d/%m", time.localtime(r["timestamp"])) for r in run_log]
        cache_sizes = [r.get("cache_size_after", 0) for r in run_log]
        cache_growth_chart = line_chart_svg(
            x_labels,
            [("Unikke titler i cache", cache_sizes, "#38bdf8")],
        )
    else:
        cache_growth_chart = '<div class="empty-state">Kør scriptet flere gange for at se udvikling over tid</div>'

    notes_rows = ""
    if notes:
        for key, note in sorted(notes.items()):
            title = cache.get(key, {}).get("title", key)
            notes_rows += f'<tr><td>{title}</td><td>{note}</td></tr>'
    notes_table = (
        f'<table class="notes"><thead><tr><th>Titel</th><th>Note</th></tr></thead>'
        f'<tbody>{notes_rows}</tbody></table>'
        if notes_rows else '<div class="empty-state">Ingen noter endnu</div>'
    )

    generated_at = time.strftime("%d. %B %Y kl. %H:%M")
    last_run_str = time.strftime("%d/%m/%Y %H:%M", time.localtime(run_log[-1]["timestamp"])) if run_log else "—"

    html = f"""<!DOCTYPE html>
<html lang="da">
<head>
<meta charset="UTF-8">
<title>Danske Backdrops — Statistik</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
<header>
    <h1>🇩🇰 Danske <span>Backdrops</span> — Statistikrapport</h1>
    <div class="timestamp">Genereret {generated_at} · Seneste kørsel: {last_run_str} · Kun backdrops (ingen postere)</div>
</header>

<div class="cards">
    <div class="card"><div class="label">Unikke titler</div><div class="value">{total_titles:,}</div></div>
    <div class="card"><div class="label">Dansk backdrop fundet</div><div class="value green">{found:,}</div></div>
    <div class="card"><div class="label">Godkendt (X)</div><div class="value blue">{approved:,}</div></div>
    <div class="card"><div class="label">Afventer godkendelse</div><div class="value amber">{pending:,}</div></div>
    <div class="card"><div class="label">Hit-rate</div><div class="value purple">{hit_rate:.1f}%</div></div>
</div>

<div class="grid2">
    <div class="panel">
        <h2>Dansk backdrop: fundet vs. ikke fundet</h2>
        <div class="donut-row">
            {donut_found}
            <div class="donut-legend">
                <div class="legend-item"><span class="legend-dot" style="background:{COLORS['found']}"></span>Fundet<b>{found:,}</b></div>
                <div class="legend-item"><span class="legend-dot" style="background:{COLORS['not_found']}"></span>Ikke fundet<b>{not_found:,}</b></div>
            </div>
        </div>
    </div>
    <div class="panel">
        <h2>Godkendelses-status <span class="sub">(af {found:,} fund)</span></h2>
        <div class="donut-row">
            {donut_review}
            <div class="donut-legend">
                <div class="legend-item"><span class="legend-dot" style="background:{COLORS['approved']}"></span>Godkendt<b>{approved:,}</b></div>
                <div class="legend-item"><span class="legend-dot" style="background:{COLORS['pending']}"></span>Afventer<b>{pending:,}</b></div>
                <div class="legend-item"><span class="legend-dot" style="background:#ef4444"></span>Markeret forkert<b>{flagged:,}</b></div>
            </div>
        </div>
    </div>
</div>

<div class="panel">
    <h2>Seneste kørsel — pr. fil <span class="sub">(sport-programmer er altid sprunget over)</span></h2>
    {per_file_bar}
</div>

<div class="panel">
    <h2>Cache-udvikling over tid <span class="sub">(antal kørsler: {len(run_log)})</span></h2>
    {cache_growth_chart}
</div>

<div class="panel">
    <h2>Noter fra godkendelsesfilen</h2>
    {notes_table}
</div>

<footer>epgoal · danish_backdrops sideprojekt (kun backdrops) · rapporten opdateres hver gang du kører generate_stats_report.py</footer>
</div>
</body>
</html>"""
    return html


def main() -> None:
    cache = load_json(DANISH_ARTWORK_CACHE_FILE, {})
    if not cache:
        sys.exit(f"❌ {DANISH_ARTWORK_CACHE_FILE} findes ikke eller er tom - kør danish_backdrops.py først.")

    run_log = load_json(DANISH_BACKDROPS_RUN_LOG_FILE, [])
    approved, flagged, notes = load_review_status(DANISH_ARTWORK_REVIEW_FILE)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    html = build_html(cache, run_log, approved, flagged, notes)
    REPORT_FILE.write_text(html, encoding="utf-8")

    print("=== Statistikrapport genereret (kun backdrops) ===")
    print(f"Fil: {REPORT_FILE}")
    print(f"Åbn den i din browser (dobbeltklik filen, eller 'start {REPORT_FILE.name}' i PowerShell).")


if __name__ == "__main__":
    main()
