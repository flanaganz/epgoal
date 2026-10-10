#!/usr/bin/env python3
"""
epgshare_sport_channel_check.py

Diagnoseværktøj: viser for HVER kanal i epgshare_filtered.xml om
SportMatcher (data/sport_channels.json) genkender den som sport-kanal
eller ej - og hvor mange programmer kanalen har.

Bruges til at afsløre om sport_channels.json's "match"-mønstre (bygget til
OpenEPG's kanal-ID-format) rammer EPGShare's kanal-ID-format eller ej.
"""

from pathlib import Path
from collections import Counter
import json
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"

INPUT_XML = ROOT / "output_epgshare" / "epgshare_filtered.xml"
SPORT_CHANNELS_FILE = DATA_DIR / "sport_channels.json"


def load_json(path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def main():
    channels_raw = load_json(SPORT_CHANNELS_FILE, {})

    if isinstance(channels_raw, list):
        exclude_patterns = []
        channels = channels_raw
    else:
        exclude_patterns = [e["match"].lower() for e in channels_raw.get("exclude", [])]
        channels = channels_raw.get("channels", [])

    print("=== sport_channels.json indhold ===")
    print(f"Antal 'channels'-regler : {len(channels)}")
    print(f"Exclude-mønstre         : {exclude_patterns or '(ingen)'}")
    print()
    print("Alle match-mønstre og roller:")
    print("-" * 80)
    for entry in channels:
        print(f"  match={entry.get('match')!r:40s}  role={entry.get('role')}")
    print()

    tree = ET.parse(INPUT_XML)
    root = tree.getroot()

    programme_counts = Counter()
    for programme in root.findall("programme"):
        ch = (programme.get("channel") or "").strip()
        programme_counts[ch] += 1

    def match_channel(channel_id):
        low = (channel_id or "").lower()
        for pattern in exclude_patterns:
            if pattern in low:
                return None
        for entry in channels:
            if entry["match"].lower() in low:
                return entry
        return None

    print("=== Kanaler i epgshare_filtered.xml og deres sport-rolle ===")
    print("-" * 80)

    for channel_id in sorted(programme_counts.keys()):
        count = programme_counts[channel_id]
        role_entry = match_channel(channel_id)
        role = role_entry.get("role") if role_entry else "(ingen sport-rolle)"
        print(f"{count:5d}  {channel_id:30s}  ->  {role}")


if __name__ == "__main__":
    main()
