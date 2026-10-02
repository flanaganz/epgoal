#!/usr/bin/env python3
"""
epgshare_merge_openepg.py — EPGShare POC: supplerer EPGShare med OpenEPG
for de kanaler hvor EPGShare01 har for kort eller intet EPG-vindue.

Input:
    output_epgshare/epgshare_filtered.xml   (EPGShare, allerede kanalfiltreret)
    config.json "sources"                   (de 6 OpenEPG-kilder produktionen bruger)
    data/channel_priority.xlsx              (Følg=X + Sammenlagte kanal-ID'er)

Output:
    output_epgshare/epgshare_merged.xml     (EPGShare + OpenEPG-udfyldning)
    data/epgshare_merge_log.json            (hvilke kanaler blev suppleret og hvorfor)

Strategi pr. Følg(X)-kanal:
    1. Find alle <channel id> i epgshare_filtered.xml der matcher kanalens
       aliaser (eksakt, case-insensitive - samme logik som epgshare_filter.py).
    2. Mål EPGShare-vinduet: samlet antal <programme> og span i timer
       (sidste start - første start) for disse id'er.
    3. Hvis span < MIN_WINDOW_HOURS (eller 0 programmer i alt):
         - Find den OpenEPG-kanal (blandt samme aliasliste) med flest
           programmer.
         - Fjern EPGShare's sparsomme <programme>-elementer for denne kanal
           og erstat dem med en kopi af OpenEPG's <programme>-elementer,
           omdøbt til EPGShare's kanal-ID (eller - hvis EPGShare slet ikke
           havde kanalen - et nyt <channel> oprettes med første alias som id).
       Ellers beholdes EPGShare's data uændret.

Kør:
    python scripts/epgshare_merge_openepg.py
"""

import copy
import json
import sys
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
CONFIG_FILE = ROOT / "config.json"
CHANNEL_PRIORITY = DATA_DIR / "channel_priority.xlsx"

OUTPUT_DIR = ROOT / "output_epgshare"
INPUT_XML = OUTPUT_DIR / "epgshare_filtered.xml"
OUTPUT_XML = OUTPUT_DIR / "epgshare_merged.xml"
MERGE_LOG_FILE = DATA_DIR / "epgshare_merge_log.json"

# Hvor mange timers EPG-vindue en EPGShare-kanal MINDST skal have, for at vi
# IKKE supplerer/erstatter den med OpenEPG. Juster denne hvis du finder at
# grænsen rammer forkert (fx hvis en kanal med 20 timer reelt er fin).
MIN_WINDOW_HOURS = 24


def load_json(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def save_json(path: Path, data) -> None:
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )


def load_followed_rows():
    """Returnerer liste af (kanonisk_navn, [alias1, alias2, ...]) for
    kanaler markeret med Følg (X)."""
    wb = load_workbook(CHANNEL_PRIORITY, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]

    kanal_col = headers.index("Kanal")
    aliases_col = headers.index("Sammenlagte kanal-ID'er")
    follow_col = headers.index("Følg (X)")

    rows = []
    for row in ws.iter_rows(min_row=2):
        if str(row[follow_col].value or "").strip().upper() != "X":
            continue
        kanal = row[kanal_col].value
        if not kanal:
            continue
        aliases_raw = row[aliases_col].value or ""
        aliases = [a.strip() for a in str(aliases_raw).split(",") if a.strip()]
        # Kanonisk navn er selv også et brugbart alias.
        aliases.append(str(kanal).strip())
        rows.append((str(kanal).strip(), aliases))

    return rows


def parse_start(start_str):
    try:
        return datetime.strptime(start_str[:14], "%Y%m%d%H%M%S")
    except (ValueError, IndexError, TypeError):
        return None


def fetch_openepg():
    """Henter ALLE 6 OpenEPG-kilder friskt (ingen cache - skal køre som en
    del af den daglige automatiske opdatering). Returnerer (root_by_source,
    combined_programmes) hvor combined_programmes er en liste af
    (channel_id, programme_element)."""
    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    if not sources:
        sys.exit("❌ Ingen kilder fundet i config.json.")

    combined_programmes = []

    for source in sources:
        name = source["name"]
        url = source["url"]
        print(f"📥 Henter OpenEPG-kilde {name} ...")
        resp = requests.get(url, timeout=90)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
        for p in root.findall("programme"):
            combined_programmes.append((p.get("channel", ""), p))

    print(f"✅ OpenEPG: {len(combined_programmes):,} programmer hentet i alt.\n")
    return combined_programmes


def main():
    if not INPUT_XML.exists():
        sys.exit(f"❌ {INPUT_XML} findes ikke - kør epgshare_filter.py først.")

    print(f"📄 Læser EPGShare (filtreret): {INPUT_XML}")
    tree = ET.parse(INPUT_XML)
    root = tree.getroot()

    # --- EPGShare: byg statistik pr. channel-id ---
    epgshare_channel_ids = {ch.get("id", "") for ch in root.findall("channel")}

    epgshare_count = {}
    epgshare_first = {}
    epgshare_last = {}
    epgshare_programmes_by_channel = {}  # channel_id -> liste af <programme> elementer

    for p in root.findall("programme"):
        cid = p.get("channel", "")
        epgshare_programmes_by_channel.setdefault(cid, []).append(p)
        dt = parse_start(p.get("start", ""))
        if dt is None:
            continue
        epgshare_count[cid] = epgshare_count.get(cid, 0) + 1
        if cid not in epgshare_first or dt < epgshare_first[cid]:
            epgshare_first[cid] = dt
        if cid not in epgshare_last or dt > epgshare_last[cid]:
            epgshare_last[cid] = dt

    # --- OpenEPG: hent frisk ---
    openepg_programmes = fetch_openepg()

    openepg_count = {}
    for cid, p in openepg_programmes:
        openepg_count[cid] = openepg_count.get(cid, 0) + 1

    # --- Gennemgå hver Følg(X)-kanal ---
    followed_rows = load_followed_rows()

    report = {
        "kept_as_is": [],
        "supplemented": [],
        "no_fallback_available": [],
    }

    for kanonisk_navn, aliases in followed_rows:
        alias_lower_set = {a.lower() for a in aliases}

        # Hvilke EPGShare channel-id'er matcher denne kanal?
        matched_epgshare_ids = [
            cid for cid in epgshare_channel_ids
            if cid.lower() in alias_lower_set
        ]

        total_count = sum(epgshare_count.get(cid, 0) for cid in matched_epgshare_ids)

        span_hours = 0
        if matched_epgshare_ids:
            firsts = [epgshare_first[c] for c in matched_epgshare_ids if c in epgshare_first]
            lasts = [epgshare_last[c] for c in matched_epgshare_ids if c in epgshare_last]
            if firsts and lasts:
                span_hours = (max(lasts) - min(firsts)).total_seconds() / 3600

        needs_supplement = (total_count == 0) or (span_hours < MIN_WINDOW_HOURS)

        if not needs_supplement:
            report["kept_as_is"].append({
                "kanal": kanonisk_navn,
                "epgshare_ids": matched_epgshare_ids,
                "programmer": total_count,
                "vindue_timer": round(span_hours, 1),
            })
            continue

        # Find bedste OpenEPG-kanal blandt samme aliasliste.
        candidate_ids = [
            cid for cid in openepg_count
            if cid.lower() in alias_lower_set
        ]

        if not candidate_ids:
            report["no_fallback_available"].append({
                "kanal": kanonisk_navn,
                "epgshare_programmer": total_count,
                "epgshare_vindue_timer": round(span_hours, 1),
            })
            continue

        best_openepg_id = max(candidate_ids, key=lambda c: openepg_count[c])
        best_count = openepg_count[best_openepg_id]

        # Vælg primær EPGShare-id: den med flest programmer blandt de
        # matchede (hvis nogen findes), ellers opret et nyt <channel> med
        # første alias som id.
        if matched_epgshare_ids:
            primary_id = max(
                matched_epgshare_ids,
                key=lambda c: epgshare_count.get(c, 0),
            )
        else:
            primary_id = aliases[0]
            new_channel_el = ET.Element("channel", {"id": primary_id})
            dn = ET.SubElement(new_channel_el, "display-name")
            dn.text = kanonisk_navn
            root.append(new_channel_el)

        # Fjern EPGShare's sparsomme programmer for alle matchede id'er.
        for cid in matched_epgshare_ids:
            for p in epgshare_programmes_by_channel.get(cid, []):
                root.remove(p)

        # Klon OpenEPG's programmer og omdøb channel-attributten.
        inserted = 0
        for cid, p in openepg_programmes:
            if cid != best_openepg_id:
                continue
            clone = copy.deepcopy(p)
            clone.set("channel", primary_id)
            root.append(clone)
            inserted += 1

        report["supplemented"].append({
            "kanal": kanonisk_navn,
            "epgshare_programmer_foer": total_count,
            "epgshare_vindue_timer_foer": round(span_hours, 1),
            "openepg_id_brugt": best_openepg_id,
            "openepg_programmer_indsat": inserted,
            "primaer_channel_id": primary_id,
        })

    # --- Gem resultat ---
    tree.write(OUTPUT_XML, encoding="utf-8", xml_declaration=True)
    save_json(MERGE_LOG_FILE, report)

    print("=== EPGShare + OpenEPG merge ===")
    print(f"Output : {OUTPUT_XML}")
    print(f"Log    : {MERGE_LOG_FILE}\n")

    print(f"✅ Beholdt EPGShare uændret : {len(report['kept_as_is']):,} kanaler")
    print(f"🔄 Suppleret med OpenEPG    : {len(report['supplemented']):,} kanaler")
    print(f"❌ Intet fallback muligt    : {len(report['no_fallback_available']):,} kanaler")

    if report["supplemented"]:
        print("\nSupplerede kanaler:")
        print("-" * 80)
        for item in report["supplemented"]:
            print(
                f"  {item['kanal']:30s} "
                f"EPGShare: {item['epgshare_programmer_foer']:4d} prog. "
                f"({item['epgshare_vindue_timer_foer']:.1f}t) -> "
                f"OpenEPG: {item['openepg_programmer_indsat']:4d} prog. "
                f"[{item['openepg_id_brugt']}]"
            )

    if report["no_fallback_available"]:
        print("\nKanaler UDEN noget fallback (hverken EPGShare eller OpenEPG har data):")
        print("-" * 80)
        for item in report["no_fallback_available"]:
            print(
                f"  {item['kanal']:30s} "
                f"EPGShare: {item['epgshare_programmer']} prog. "
                f"({item['epgshare_vindue_timer']:.1f}t)"
            )

    print("\n=== Færdig. ===")


if __name__ == "__main__":
    main()
