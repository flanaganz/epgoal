#!/usr/bin/env python3

from pathlib import Path
from collections import Counter
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent

INPUT_XML = ROOT / "output_epgshare" / "epgshare_filtered.xml"

tree = ET.parse(INPUT_XML)
root = tree.getroot()

titles = Counter()

for programme in root.findall("programme"):

    title = programme.findtext("title")

    if title:
        titles[title.strip()] += 1

print()
print("===== EPGShare POC =====")
print()

print(f"Programmer       : {sum(titles.values()):,}")
print(f"Unikke titler    : {len(titles):,}")

print()
print("Top 100 kandidat-titler")
print("-" * 80)

for title, count in titles.most_common(100):
    print(f"{count:4d}  {title}")