#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
from collections import Counter
import xml.etree.ElementTree as ET

BASE = Path(r"C:\EPGoal\output_epgshare")

FILES = {
    "EPGShare Raw": BASE / "epgshare_dk1.xml",
    "EPGShare Filtered": BASE / "epgshare_filtered.xml",
    "EPGShare Merged": BASE / "epgshare_merged.xml",
    "EPGoal": BASE / "epgoal.xml",
}

TV3_PATTERNS = [
    "tv3",
    "tv3+",
]

def inspect_xml(path):
    channels = {}
    programme_count = Counter()

    if not path.exists():
        return channels, programme_count

    tree = ET.parse(path)
    root = tree.getroot()

    for ch in root.findall("channel"):
        cid = (ch.get("id") or "").strip()

        names = [
            (dn.text or "").strip()
            for dn in ch.findall("display-name")
            if (dn.text or "").strip()
        ]

        channels[cid] = names

    for p in root.findall("programme"):
        cid = (p.get("channel") or "").strip()
        programme_count[cid] += 1

    return channels, programme_count


def is_tv3_related(channel_id, names):
    text = channel_id.lower()

    for name in names:
        text += " " + name.lower()

    return any(x in text for x in TV3_PATTERNS)


def print_stage(name, path):

    print()
    print("=" * 70)
    print(name)
    print("=" * 70)

    channels, programme_count = inspect_xml(path)

    tv3_channels = []

    for cid, names in channels.items():

        if is_tv3_related(cid, names):
            tv3_channels.append((cid, names))

    if not tv3_channels:
        print("Ingen TV3-relaterede kanaler fundet.")
        return

    for cid, names in sorted(tv3_channels):

        print(f"Channel ID : {cid}")
        print(f"Navne      : {names}")
        print(f"Programmer : {programme_count[cid]}")
        print()

    print("Opsummering:")
    print(f"TV3-kanaler fundet : {len(tv3_channels)}")
    print(
        f"TV3-programmer i alt : "
        f"{sum(programme_count[cid] for cid, _ in tv3_channels)}"
    )


def main():

    print()
    print("=== TV3 PIPELINE ANALYSE ===")
    print()

    for stage_name, xml_file in FILES.items():
        print_stage(stage_name, xml_file)


if __name__ == "__main__":
    main()