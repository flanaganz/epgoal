#!/usr/bin/env python3
"""
openepg_coverage_check.py

Henter de samme 6 OpenEPG-kilder som produktionen bruger (fra config.json),
og viser for HVER kanal i channel_priority.xlsx (dem der er markeret med
Følg (X)):
  - hvilke af de sammenlagte kanal-ID-varianter der rent faktisk findes i
    OpenEPG,
  - hvor mange programmer hver variant har,
  - første og sidste program-dato (dvs. hvor langt EPG-vinduet rækker).

Formål: bekræfte/afkræfte observationerne i EPG Mangler.xlsx (kolonne F) -
dvs. om OpenEPG reelt dækker de kanaler/varianter hvor EPGShare01 enten
mangler data helt, eller kun har kort/ujævn dækning på nogle af
HD/FHD-varianterne.

Kør:
    python scripts/openepg_coverage_check.py
"""

from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime
import json
import sys

import requests
from openpyxl import load_workbook
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
CONFIG_FILE = ROOT / "config.json"
CHANNEL_PRIORITY = DATA_DIR / "channel_priority.xlsx"

OUTPUT_DIR = ROOT / "output_epgshare"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CACHE_XML_DIR = OUTPUT_DIR / "openepg_raw"
CACHE_XML_DIR.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def load_channel_priority():
    """Returnerer liste af (kanonisk_navn, [alias1, alias2, ...], følg_bool)."""
    wb = load_workbook(CHANNEL_PRIORITY, data_only=True)
    ws = wb.active

    headers = [c.value for c in ws[1]]
    kanal_col = headers.index("Kanal")
    aliases_col = headers.index("Sammenlagte kanal-ID'er")
    follow_col = headers.index("Følg (X)")

    rows = []
    for row in ws.iter_rows(min_row=2):
        kanal = row[kanal_col].value
        if not kanal:
            continue
        follow = str(row[follow_col].value or "").strip().upper() == "X"
        if not follow:
            continue

        aliases_raw = row[aliases_col].value or ""
        aliases = [a.strip() for a in str(aliases_raw).split(",") if a.strip()]
        rows.append((str(kanal).strip(), aliases))

    return rows


def fetch_openepg_sources():
    """Henter (eller genbruger cachede) OpenEPG XML-filer fra config.json."""
    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    if not sources:
        sys.exit("❌ Ingen kilder fundet i config.json.")

    all_programmes = []  # liste af (channel_id, start_str, title)
    all_channel_ids = set()

    for source in sources:
        name = source["name"]
        url = source["url"]

        cache_path = CACHE_XML_DIR / f"{name}.xml"

        if cache_path.exists():
            print(f"📄 Genbruger cachet fil: {cache_path.name}")
            content = cache_path.read_bytes()
        else:
            print(f"📥 Henter {name} fra {url} ...")
            resp = requests.get(url, timeout=90)
            resp.raise_for_status()
            content = resp.content
            cache_path.write_bytes(content)
            print(f"✅ Downloadet og cachet: {len(content)//1024} KB")

        root = ET.fromstring(content)

        for ch in root.findall("channel"):
            cid = ch.get("id", "")
            if cid:
                all_channel_ids.add(cid)

        for p in root.findall("programme"):
            cid = p.get("channel", "")
            start = p.get("start", "")
            title_el = p.find("title")
            title = title_el.text if title_el is not None else ""
            all_programmes.append((cid, start, title))

    return all_channel_ids, all_programmes


def parse_start(start_str):
    try:
        return datetime.strptime(start_str[:14], "%Y%m%d%H%M%S")
    except (ValueError, IndexError):
        return None


def main():
    print("=== OpenEPG dæknings-tjek (produktionens 6 kilder) ===\n")

    channel_rows = load_channel_priority()
    all_channel_ids, all_programmes = fetch_openepg_sources()

    print(f"\nTotal antal <channel> i OpenEPG: {len(all_channel_ids):,}")
    print(f"Total antal <programme> i OpenEPG: {len(all_programmes):,}\n")

    # Byg statistik pr. kanal-id
    prog_count = Counter()
    first_seen = {}
    last_seen = {}

    for cid, start_str, _title in all_programmes:
        dt = parse_start(start_str)
        if dt is None:
            continue
        prog_count[cid] += 1
        if cid not in first_seen or dt < first_seen[cid]:
            first_seen[cid] = dt
        if cid not in last_seen or dt > last_seen[cid]:
            last_seen[cid] = dt

    print("=== Dækning pr. kanal (fra channel_priority.xlsx, Følg=X) ===")
    print("-" * 100)

    for kanonisk_navn, aliases in channel_rows:
        print(f"\n{kanonisk_navn}")

        any_found = False
        for alias in aliases:
            # Match både eksakt og case-insensitive, da OpenEPG kan have
            # andre skrivemåder end aliaslisten.
            matched_id = None
            for cid in all_channel_ids:
                if cid.lower() == alias.lower():
                    matched_id = cid
                    break

            if matched_id is None:
                continue

            any_found = True
            count = prog_count.get(matched_id, 0)
            if count == 0:
                print(f"   [{matched_id}] -> <channel> findes, men 0 programmer")
                continue

            first_dt = first_seen.get(matched_id)
            last_dt = last_seen.get(matched_id)
            days = (last_dt - first_dt).days if first_dt and last_dt else 0

            print(
                f"   [{matched_id}] -> {count:,} programmer | "
                f"{first_dt} -> {last_dt} (~{days} dage)"
            )

        if not any_found:
            print("   (Ingen af aliaserne findes i OpenEPG's <channel>-liste)")

    print("\n" + "-" * 100)
    print("Færdig. Sammenhold output med noterne i EPG Mangler.xlsx (kolonne F).")


if __name__ == "__main__":
    main()
