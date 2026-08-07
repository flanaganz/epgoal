#!/usr/bin/env python3
"""
export_danish_artwork_review.py — eksporterer danske TMDb-fund til Excel til manuel godkendelse

FORMÅL
    Læser data/danish_artwork_cache.json (fra danish_backdrops.py) og skriver
    alle titler MED et fundet dansk backdrop/poster til
    data/danish_artwork_review.xlsx, så du kan gennemgå dem og markere "X" i
    kolonnen "Godkendt (X)" for de fund, der er korrekte og skal bruges.

    Kør dette script igen, når du har fundet flere danske backdrops over tid
    (fx efter flere kørsler af danish_backdrops.py) - eksisterende
    X-markeringer og noter BEVARES automatisk for titler, der allerede findes
    i filen. Kun NYE titler tilføjes som friske, tomme rækker.

WORKFLOW
    1) python3 scripts/danish_backdrops.py           (finder danske billeder)
    2) python3 scripts/export_danish_artwork_review.py  (eksporter til Excel)
    3) Åbn data/danish_artwork_review.xlsx, markér "X", gem filen
    4) python3 scripts/danish_backdrops.py           (injicerer de godkendte)

BRUG
    python3 scripts/export_danish_artwork_review.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"

DANISH_ARTWORK_CACHE_FILE = DATA_DIR / "danish_artwork_cache.json"
DANISH_ARTWORK_REVIEW_FILE = DATA_DIR / "danish_artwork_review.xlsx"

HEADERS = ["Nøgle (intern)", "Titel", "Dansk Backdrop", "Dansk Poster", "Godkendt (X)", "Note"]


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return default
    return default


def load_existing_review(path: Path) -> dict[str, dict]:
    """Læser eksisterende Excel-fil (hvis den findes) og returnerer
    {nøgle: {"godkendt": ..., "note": ...}} så vi kan bevare brugerens
    tidligere markeringer ved re-eksport."""
    existing: dict[str, dict] = {}
    if not path.exists():
        return existing

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    try:
        key_col = headers.index("Nøgle (intern)")
        godkendt_col = headers.index("Godkendt (X)")
        note_col = headers.index("Note") if "Note" in headers else None
    except ValueError:
        print(f"⚠️  Eksisterende {path.name} har uventet format - starter forfra uden at bevare markeringer.",
              file=sys.stderr)
        return existing

    for row in ws.iter_rows(min_row=2):
        key_val = row[key_col].value
        if not key_val:
            continue
        existing[str(key_val).strip()] = {
            "godkendt": row[godkendt_col].value,
            "note": row[note_col].value if note_col is not None else None,
        }
    return existing


def main() -> None:
    cache = load_json(DANISH_ARTWORK_CACHE_FILE, {})
    if not cache:
        sys.exit(f"❌ {DANISH_ARTWORK_CACHE_FILE} findes ikke eller er tom - kør danish_backdrops.py først.")

    candidates = {
        key: entry for key, entry in cache.items()
        if entry.get("backdrop") or entry.get("poster")
    }

    if not candidates:
        sys.exit("❌ Ingen danske backdrops/postere fundet endnu i cachen - intet at eksportere.")

    existing_review = load_existing_review(DANISH_ARTWORK_REVIEW_FILE)
    new_count = sum(1 for key in candidates if key not in existing_review)
    preserved_count = sum(1 for key in candidates if key in existing_review and existing_review[key].get("godkendt"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Danske backdrops"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2E5395")
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F1"

    sorted_keys = sorted(candidates.keys(), key=lambda k: candidates[k].get("title", k).lower())

    for row_idx, key in enumerate(sorted_keys, start=2):
        entry = candidates[key]
        title = entry.get("title") or key.title()
        backdrop_url = entry.get("backdrop")
        poster_url = entry.get("poster")
        prior = existing_review.get(key, {})

        ws.cell(row=row_idx, column=1, value=key)
        ws.cell(row=row_idx, column=2, value=title)

        if backdrop_url:
            c = ws.cell(row=row_idx, column=3, value="Åbn backdrop")
            c.hyperlink = backdrop_url
            c.style = "Hyperlink"
        if poster_url:
            c = ws.cell(row=row_idx, column=4, value="Åbn poster")
            c.hyperlink = poster_url
            c.style = "Hyperlink"

        godkendt_val = prior.get("godkendt")
        if godkendt_val:
            ws.cell(row=row_idx, column=5, value=godkendt_val)
        note_val = prior.get("note")
        if note_val:
            ws.cell(row=row_idx, column=6, value=note_val)

    last_row = len(sorted_keys) + 1
    dv = DataValidation(type="list", formula1='"X"', allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"E2:E{last_row}")

    widths = {"A": 32, "B": 45, "C": 16, "D": 16, "E": 14, "F": 35}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(DANISH_ARTWORK_REVIEW_FILE)

    print("=== Danske backdrops eksporteret til Excel ===")
    print(f"Fil: {DANISH_ARTWORK_REVIEW_FILE}")
    print(f"Titler i alt (med fundet dansk billede): {len(candidates):,}")
    print(f"  - Nye rækker tilføjet denne gang     : {new_count:,}")
    print(f"  - Eksisterende X-markeringer bevaret : {preserved_count:,}")
    print()
    print("Åbn filen, markér 'X' i kolonnen 'Godkendt (X)' for de rigtige fund, gem filen,")
    print("og kør derefter 'python3 scripts/danish_backdrops.py' igen for at injicere dem.")


if __name__ == "__main__":
    main()
