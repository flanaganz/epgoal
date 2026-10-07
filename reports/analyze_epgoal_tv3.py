#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
from collections import Counter, defaultdict
import xml.etree.ElementTree as ET

EPGOAL = Path(r"C:\EPGoal\output_epgshare\epgoal.xml")

TV3_IDS = {
    "TV3.dk",
    "tv3.dk",
    "TV3+.dk",
    "tv3plus.dk",
}

def main():

    print("=== ANALYSE AF EPGOAL.XML ===")
    print(f"Fil: {EPGOAL}")
    print()

    tree = ET.parse(EPGOAL)
    root = tree.getroot()

    channels = {}
    programme_count = Counter()

    for ch in root.findall("channel"):
        cid = ch.get("id", "").strip()

        names = []
        for dn in ch.findall("display-name"):
            if dn.text:
                names.append(dn.text.strip())

        channels[cid] = names

    for p in root.findall("programme"):
        cid = p.get("channel", "").strip()
        programme_count[cid] += 1

    print("=== TV3 RELATEREDE CHANNELS ===")
    print()

    for cid, names in sorted(channels.items()):

        if "tv3" in cid.lower():

            print(f"Channel ID : {cid}")
            print(f"Navne      : {names}")
            print(f"Programmer : {programme_count[cid]}")
            print()

    print("=== CHANNELS UDEN PROGRAMMER ===")
    print()

    for cid in sorted(channels):

        if programme_count[cid] == 0:

            print(cid)

    print()
    print("=== PROGRAMMER UDEN CHANNEL ===")
    print()

    channel_ids = set(channels)

    orphans = {}

    for p in root.findall("programme"):

        cid = p.get("channel", "").strip()

        if cid not in channel_ids:

            orphans[cid] = orphans.get(cid, 0) + 1

    for cid, count in sorted(orphans.items()):

        print(f"{cid}: {count}")

    print()
    print("=== CASE-ANALYSE ===")
    print()

    groups = defaultdict(list)

    for cid in channels:

        groups[cid.lower()].append(cid)

    for lower, variants in sorted(groups.items()):

        if len(variants) > 1:

            print(lower)
            for v in variants:
                print(f"   {v}")
            print()

    print()
    print("=== TV3 SPECIFIKT ===")
    print()

    for cid in sorted(TV3_IDS):

        print(
            f"{cid:15} "
            f"channel={'JA' if cid in channels else 'NEJ'} "
            f"programmer={programme_count[cid]}"
        )

if __name__ == "__main__":
    main()