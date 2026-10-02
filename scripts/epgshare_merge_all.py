#!/usr/bin/env python3
"""
epgshare_merge_all.py — EPGShare POC: samler EPGShare01 + OpenEPG + BSS
(ultratv.one) til ÉN samlet XML.

Prioritet pr. Følg(X)-kanal (fra data/channel_priority.xlsx):
    1. EPGShare01 - hvis vinduet er >= MIN_WINDOW_HOURS, beholdes den uændret.
    2. Ellers: sammenlign OpenEPG (6 kilder fra config.json) og BSS
       (ultratv.one) for samme kanal, og brug den kilde der har det
       LÆNGSTE vindue / flest programmer. EPGShare's sparsomme data for
       kanalen fjernes og erstattes.
    3. Hvis INGEN af de tre kilder har data -> logges som "intet fallback".

BSS-URL'en indeholder login og hentes derfor IKKE fra denne fil, men fra
miljøvariablen BSS_XMLTV_URL i .env (samme mønster som TMDB_API_KEY) - så
den ALDRIG committes til GitHub sammen med resten af output_epgshare/.

Output:
    output_epgshare/epgshare_merged.xml   (EPGShare + OpenEPG/BSS-udfyldning)
    data/epgshare_merge_log.json          (hvilken kilde blev brugt pr. kanal, og hvorfor)

Kør:
    python scripts/epgshare_merge_all.py
"""

import copy
import json
import os
import re
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

# Hvor mange timers EPG-vindue en kilde MINDST skal have, for at vi
# IKKE leder videre efter en bedre kilde. Juster hvis grænsen rammer forkert.
MIN_WINDOW_HOURS = 24

ENV_FILE = ROOT / ".env"
if ENV_FILE.exists():
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

BSS_XMLTV_URL = os.environ.get("BSS_XMLTV_URL")


def load_json(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def save_json(path: Path, data) -> None:
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )


def normalize_id(s: str) -> str:
    """Normaliserer et kanal-id/alias til ren a-z0-9 (ingen punktummer,
    mellemrum, apostroffer, '.dk'-endelser osv.) - gør OpenEPG's
    "Kanal 4.dk", BSS's "kanal4.dk" og EPGShare's "Kanal.4.dk" sammenlignelige."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def load_followed_rows():
    """Returnerer liste af (kanonisk_navn, normaliserede_tokens-set)."""
    wb = load_workbook(CHANNEL_PRIORITY, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]

    kanal_col = headers.index("Kanal")
    key_col = headers.index("Gruppe-nøgle (intern)")
    aliases_col = headers.index("Sammenlagte kanal-ID'er")
    follow_col = headers.index("Følg (X)")

    rows = []
    for row in ws.iter_rows(min_row=2):
        if str(row[follow_col].value or "").strip().upper() != "X":
            continue
        kanal = row[kanal_col].value
        if not kanal:
            continue

        tokens = set()
        key = row[key_col].value
        if key:
            tokens.add(normalize_id(str(key)))

        aliases_raw = row[aliases_col].value or ""
        for a in str(aliases_raw).split(","):
            a = a.strip()
            if a:
                tokens.add(normalize_id(a))

        tokens.add(normalize_id(str(kanal)))
        tokens.discard("")

        rows.append((str(kanal).strip(), tokens))

    return rows


def parse_start(start_str):
    try:
        return datetime.strptime(start_str[:14], "%Y%m%d%H%M%S")
    except (ValueError, IndexError, TypeError):
        return None


def build_channel_stats(programmes_by_channel):
    """programmes_by_channel: dict channel_id -> list af <programme>.
    Returnerer dict channel_id -> {"count": int, "span_hours": float}."""
    stats = {}
    for cid, plist in programmes_by_channel.items():
        dates = [parse_start(p.get("start", "")) for p in plist]
        dates = [d for d in dates if d is not None]
        count = len(dates)
        span = (max(dates) - min(dates)).total_seconds() / 3600 if dates else 0
        stats[cid] = {"count": count, "span_hours": span}
    return stats


def fetch_openepg():
    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    if not sources:
        print("⚠️  Ingen kilder fundet i config.json - springer OpenEPG over.")
        return {}

    programmes_by_channel: dict[str, list] = {}

    for source in sources:
        name = source["name"]
        url = source["url"]
        print(f"📥 Henter OpenEPG-kilde {name} ...")
        try:
            resp = requests.get(url, timeout=90)
            resp.raise_for_status()
        except requests.RequestException as exc:
            print(f"⚠️  Kunne ikke hente {name}: {exc}", file=sys.stderr)
            continue
        root = ET.fromstring(resp.content)
        for p in root.findall("programme"):
            cid = p.get("channel", "")
            programmes_by_channel.setdefault(cid, []).append(p)

    total = sum(len(v) for v in programmes_by_channel.values())
    print(f"✅ OpenEPG: {total:,} programmer hentet i alt.\n")
    return programmes_by_channel


def fetch_bss():
    if not BSS_XMLTV_URL:
        print("ℹ️  BSS_XMLTV_URL er ikke sat i .env - springer BSS over.\n")
        return {}

    print("📥 Henter BSS (ultratv.one) ...")
    try:
        resp = requests.get(BSS_XMLTV_URL, timeout=300)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"⚠️  Kunne ikke hente BSS: {exc}", file=sys.stderr)
        return {}

    root = ET.fromstring(resp.content)
    programmes_by_channel: dict[str, list] = {}
    for p in root.findall("programme"):
        cid = p.get("channel", "")
        programmes_by_channel.setdefault(cid, []).append(p)

    total = sum(len(v) for v in programmes_by_channel.values())
    print(f"✅ BSS: {total:,} programmer hentet i alt.\n")
    return programmes_by_channel


def find_best_candidate(tokens: set, programmes_by_channel: dict, stats: dict):
    """Finder det channel-id i programmes_by_channel hvis normaliserede id
    matcher et af tokens, med flest programmer (ved lige antal: længst span)."""
    best_id = None
    best_stat = {"count": 0, "span_hours": 0}

    for cid in programmes_by_channel:
        if normalize_id(cid) in tokens:
            cand_stat = stats.get(cid, {"count": 0, "span_hours": 0})
            if (cand_stat["count"], cand_stat["span_hours"]) > (best_stat["count"], best_stat["span_hours"]):
                best_id = cid
                best_stat = cand_stat

    return best_id, best_stat


def strip_artwork(programme_el) -> None:
    for tag in ("icon", "backdrop"):
        for old in programme_el.findall(tag):
            programme_el.remove(old)


def main():
    if not INPUT_XML.exists():
        sys.exit(f"❌ {INPUT_XML} findes ikke - kør epgshare_filter.py først.")

    print(f"📄 Læser EPGShare (filtreret): {INPUT_XML}")
    tree = ET.parse(INPUT_XML)
    root = tree.getroot()

    epgshare_channel_ids = {ch.get("id", "") for ch in root.findall("channel")}

    epgshare_programmes_by_channel: dict[str, list] = {}
    for p in root.findall("programme"):
        cid = p.get("channel", "")
        epgshare_programmes_by_channel.setdefault(cid, []).append(p)

    epgshare_stats = build_channel_stats(epgshare_programmes_by_channel)

    openepg_programmes_by_channel = fetch_openepg()
    openepg_stats = build_channel_stats(openepg_programmes_by_channel)

    bss_programmes_by_channel = fetch_bss()
    bss_stats = build_channel_stats(bss_programmes_by_channel)

    followed_rows = load_followed_rows()

    report = {
        "kept_as_is": [],
        "supplemented": [],
        "no_fallback_available": [],
    }

    for kanonisk_navn, tokens in followed_rows:
        # --- EPGShare: find matchende id'er og deres samlede dækning ---
        matched_epgshare_ids = [
            cid for cid in epgshare_channel_ids if normalize_id(cid) in tokens
        ]
        epgshare_count = sum(epgshare_stats.get(c, {}).get("count", 0) for c in matched_epgshare_ids)
        epgshare_span = max(
            (epgshare_stats.get(c, {}).get("span_hours", 0) for c in matched_epgshare_ids),
            default=0,
        )

        if epgshare_count > 0 and epgshare_span >= MIN_WINDOW_HOURS:
            report["kept_as_is"].append({
                "kanal": kanonisk_navn,
                "kilde": "epgshare",
                "programmer": epgshare_count,
                "vindue_timer": round(epgshare_span, 1),
            })
            continue

        # --- Find bedste alternativ blandt OpenEPG og BSS ---
        openepg_id, openepg_stat = find_best_candidate(tokens, openepg_programmes_by_channel, openepg_stats)
        bss_id, bss_stat = find_best_candidate(tokens, bss_programmes_by_channel, bss_stats)

        candidates = []
        if openepg_id and openepg_stat["count"] > 0:
            candidates.append(("openepg", openepg_id, openepg_stat, openepg_programmes_by_channel))
        if bss_id and bss_stat["count"] > 0:
            candidates.append(("bss", bss_id, bss_stat, bss_programmes_by_channel))

        if not candidates:
            report["no_fallback_available"].append({
                "kanal": kanonisk_navn,
                "epgshare_programmer": epgshare_count,
                "epgshare_vindue_timer": round(epgshare_span, 1),
            })
            continue

        # Vælg kilden med længst vindue (ved lige: flest programmer).
        candidates.sort(key=lambda c: (c[2]["span_hours"], c[2]["count"]), reverse=True)
        best_source, best_id, best_stat, best_programmes_map = candidates[0]

        # Hvis EPGShare rent faktisk var bedre end BÅDE OpenEPG og BSS (men
        # stadig under MIN_WINDOW_HOURS), så behold EPGShare alligevel -
        # ingen grund til at erstatte med noget dårligere.
        if epgshare_span >= best_stat["span_hours"] and epgshare_count > 0:
            report["kept_as_is"].append({
                "kanal": kanonisk_navn,
                "kilde": "epgshare (bedst tilgængelige, selvom < min. vindue)",
                "programmer": epgshare_count,
                "vindue_timer": round(epgshare_span, 1),
            })
            continue

        # Vælg/opret primær channel-id.
        if matched_epgshare_ids:
            primary_id = max(
                matched_epgshare_ids,
                key=lambda c: epgshare_stats.get(c, {}).get("count", 0),
            )
        else:
            primary_id = f"{normalize_id(kanonisk_navn)}.dk"
            new_channel_el = ET.Element("channel", {"id": primary_id})
            dn = ET.SubElement(new_channel_el, "display-name")
            dn.text = kanonisk_navn
            root.append(new_channel_el)

        # Fjern EPGShare's sparsomme programmer for alle matchede id'er.
        for cid in matched_epgshare_ids:
            for p in epgshare_programmes_by_channel.get(cid, []):
                root.remove(p)

        # Indsæt den valgte kildes programmer under primary_id.
        inserted = 0
        for p in best_programmes_map.get(best_id, []):
            clone = copy.deepcopy(p)
            clone.set("channel", primary_id)
            strip_artwork(clone)
            root.append(clone)
            inserted += 1

        report["supplemented"].append({
            "kanal": kanonisk_navn,
            "epgshare_programmer_foer": epgshare_count,
            "epgshare_vindue_timer_foer": round(epgshare_span, 1),
            "kilde_brugt": best_source,
            "kilde_channel_id": best_id,
            "programmer_indsat": inserted,
            "vindue_timer_efter": round(best_stat["span_hours"], 1),
            "primaer_channel_id": primary_id,
        })

    tree.write(OUTPUT_XML, encoding="utf-8", xml_declaration=True)
    save_json(MERGE_LOG_FILE, report)

    print("=== EPGShare + OpenEPG + BSS merge ===")
    print(f"Output : {OUTPUT_XML}")
    print(f"Log    : {MERGE_LOG_FILE}\n")

    print(f"✅ Beholdt EPGShare (eller bedst tilgængelige) : {len(report['kept_as_is']):,} kanaler")
    print(f"🔄 Suppleret med OpenEPG/BSS                   : {len(report['supplemented']):,} kanaler")
    print(f"❌ Intet fallback muligt                       : {len(report['no_fallback_available']):,} kanaler")

    if report["supplemented"]:
        print("\nSupplerede kanaler:")
        print("-" * 90)
        for item in report["supplemented"]:
            print(
                f"  {item['kanal']:30s} "
                f"EPGShare: {item['epgshare_programmer_foer']:4d} prog. "
                f"({item['epgshare_vindue_timer_foer']:.1f}t) -> "
                f"{item['kilde_brugt'].upper():8s}: {item['programmer_indsat']:4d} prog. "
                f"({item['vindue_timer_efter']:.1f}t) [{item['kilde_channel_id']}]"
            )

    if report["no_fallback_available"]:
        print("\nKanaler UDEN noget fallback (ingen af de 3 kilder har brugbar dækning):")
        print("-" * 90)
        for item in report["no_fallback_available"]:
            print(
                f"  {item['kanal']:30s} "
                f"EPGShare: {item['epgshare_programmer']} prog. "
                f"({item['epgshare_vindue_timer']:.1f}t)"
            )

    print("\n=== Færdig. ===")


if __name__ == "__main__":
    main()
