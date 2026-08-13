#!/usr/bin/env python3
"""
channel_health.py — scanner de FÆRDIGE output/denmarkX.xml-filer og måler
reel artwork-dækning pr. kanal.

Inspireret af brugerens tidligere GoldEPG-rapport (VIP Channel Health-tabel +
Top-10-kanaler-med-manglende-artwork), men RETTET til at sammenlægge samme
kanal på tværs af alle 6 kildefiler (det gamle eksempel viste hver kanal to
gange med identiske tal - et tegn på at kanaler ikke blev dedupliceret på
tværs af filer).

VIGTIGT: Dette script kører EFTER både enrich_epg.py OG danish_backdrops.py,
da det måler den FÆRDIGE tilstand (uanset om artwork kom fra sport-scriptet,
danske TMDb-backdrops, eller manuelle overrides - det tæller bare om
programmet har fået et <icon> eller <backdrop> sat, uden at skelne mellem
kilderne).

Skriver:
    data/channel_health.json - seneste snapshot (bruges af generate_stats_report.py)

BRUG
    python3 scripts/channel_health.py
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
CONFIG_FILE = ROOT / "config.json"
CHANNEL_HEALTH_FILE = DATA_DIR / "channel_health.json"


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return default
    return default


def save_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def scan_file(xml_path: Path, channel_stats: dict[str, dict]) -> int:
    """Opdaterer channel_stats IN-PLACE (nøgle = kanal-id). Returnerer antal
    programmer behandlet i denne fil."""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    display_names: dict[str, str] = {}
    for ch in root.findall("channel"):
        cid = ch.get("id", "")
        dn_el = ch.find("display-name")
        display_names[cid] = dn_el.text.strip() if dn_el is not None and dn_el.text else cid

    count = 0
    for programme in root.findall("programme"):
        count += 1
        chan_id = programme.get("channel", "")
        if chan_id not in channel_stats:
            channel_stats[chan_id] = {
                "display_name": display_names.get(chan_id, chan_id),
                "programmes": 0,
                "with_artwork": 0,
                "with_desc": 0,
            }
        stats = channel_stats[chan_id]
        stats["programmes"] += 1

        has_artwork = programme.find("icon") is not None or programme.find("backdrop") is not None
        if has_artwork:
            stats["with_artwork"] += 1

        desc_el = programme.find("desc")
        if desc_el is not None and desc_el.text and desc_el.text.strip():
            stats["with_desc"] += 1

    return count


def main() -> None:
    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    source_names = [s["name"] for s in sources] if sources else None

    if source_names is None:
        # Fallback: scan alle denmark*.xml i output/ hvis config.json ikke findes/er tom
        xml_files = sorted(OUTPUT_DIR.glob("denmark*.xml"))
    else:
        xml_files = [OUTPUT_DIR / f"{name}.xml" for name in source_names]
        xml_files = [p for p in xml_files if p.exists()]

    if not xml_files:
        sys.exit(f"❌ Ingen output/denmarkX.xml-filer fundet i {OUTPUT_DIR} - kør enrich_epg.py først.")

    print("=== Kanal-sundhedstjek (scanner FÆRDIGE output-filer) ===")
    print(f"Behandler {len(xml_files)} fil(er): {', '.join(p.name for p in xml_files)}")

    channel_stats: dict[str, dict] = {}
    total_programmes = 0
    for xml_path in xml_files:
        count = scan_file(xml_path, channel_stats)
        total_programmes += count
        print(f"   {xml_path.name}: {count:,} programmer")

    total_with_artwork = sum(c["with_artwork"] for c in channel_stats.values())
    overall_pct = (total_with_artwork / total_programmes * 100) if total_programmes else 0

    snapshot = {
        "timestamp": time.time(),
        "date_str": time.strftime("%Y-%m-%d %H:%M:%S"),
        "total_programmes": total_programmes,
        "total_with_artwork": total_with_artwork,
        "overall_artwork_pct": round(overall_pct, 2),
        "channels": channel_stats,
    }
    save_json(CHANNEL_HEALTH_FILE, snapshot)

    print(f"\n📊 SAMLET: {total_programmes:,} programmer på tværs af {len(channel_stats):,} kanaler")
    print(f"   Artwork-dækning i alt: {total_with_artwork:,} ({overall_pct:.1f}%)")

    worst = sorted(channel_stats.items(), key=lambda kv: kv[1]["programmes"] - kv[1]["with_artwork"], reverse=True)[:10]
    if worst:
        print("\n   Top 10 kanaler med flest MANGLENDE artwork (absolut antal):")
        for cid, s in worst:
            missing = s["programmes"] - s["with_artwork"]
            if missing == 0:
                continue
            pct = (s["with_artwork"] / s["programmes"] * 100) if s["programmes"] else 0
            print(f"     - {s['display_name']:40s} {missing:>5,} mangler  ({pct:.1f}% dækning, {s['programmes']:,} i alt)")

    print(f"\nGemt: {CHANNEL_HEALTH_FILE}")
    print("=== Færdig. ===")


if __name__ == "__main__":
    main()
