#!/usr/bin/env python3

from pathlib import Path
import xml.etree.ElementTree as ET
from collections import Counter

ROOT = Path(__file__).resolve().parent.parent

INPUT_XML = ROOT / "output_epgshare" / "epgshare_dk1.xml"
OUTPUT_TXT = ROOT / "output_epgshare" / "all_channels.txt"


def main():

    print(f"Indlæser: {INPUT_XML}")

    tree = ET.parse(INPUT_XML)
    root = tree.getroot()

    channel_names = {}
    programme_counts = Counter()

    for channel in root.findall("channel"):

        channel_id = (channel.get("id") or "").strip()

        display_names = [
            (dn.text or "").strip()
            for dn in channel.findall("display-name")
            if dn.text
        ]

        channel_names[channel_id] = display_names

    for programme in root.findall("programme"):

        channel_id = (programme.get("channel") or "").strip()

        if channel_id:
            programme_counts[channel_id] += 1

    rows = []

    for channel_id in sorted(channel_names):

        display = ", ".join(channel_names[channel_id])

        rows.append(
            (
                channel_id,
                programme_counts[channel_id],
                display
            )
        )

    with open(OUTPUT_TXT, "w", encoding="utf-8") as f:

        f.write("KANAL-ID | PROGRAMMER | DISPLAY-NAME\n")
        f.write("=" * 120 + "\n")

        for channel_id, count, display in rows:
            line = f"{channel_id} | {count:,} | {display}"
            f.write(line + "\n")

    print()
    print("=== EPGShare kanalanalyse ===")
    print(f"Unikke kanaler : {len(rows):,}")
    print(f"Output         : {OUTPUT_TXT}")
    print()

    print("Top 50 kanaler:")
    print("-" * 80)

    for channel_id, count, display in sorted(
        rows,
        key=lambda x: x[1],
        reverse=True
    )[:50]:

        print(f"{count:5,d}  {channel_id}")

    print("-" * 80)


if __name__ == "__main__":
    main()