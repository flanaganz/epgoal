#!/usr/bin/env python3
"""
epgshare_download.py — EPGShare POC: henter frisk EPGShare01 DK1 XML.

Downloader https://epgshare01.online/epgshare01/epg_ripper_DK1.xml.gz,
pakker den ud (gzip) og gemmer som output_epgshare/epgshare_dk1.xml -
samme sti som epgshare_filter.py forventer som input.

Opdateres af EPGShare01 ca. én gang i døgnet.
"""

import gzip
import sys
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "output_epgshare"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

URL = "https://epgshare01.online/epgshare01/epg_ripper_DK1.xml.gz"
OUTPUT_FILE = OUTPUT_DIR / "epgshare_dk1.xml"


def main() -> None:
    print(f"📥 Henter EPGShare01 DK1 fra {URL} ...")

    try:
        resp = requests.get(URL, timeout=300)
        resp.raise_for_status()
    except requests.RequestException as exc:
        sys.exit(f"❌ Download fejlede: {exc}")

    print(f"✅ Downloadet {len(resp.content)//1024} KB (gzip)")
    print("📦 Udpakker gzip ...")

    try:
        xml_data = gzip.decompress(resp.content)
    except OSError as exc:
        sys.exit(f"❌ Kunne ikke udpakke gzip: {exc}")

    OUTPUT_FILE.write_bytes(xml_data)

    print(f"💾 Gemt: {OUTPUT_FILE}")
    print(f"   Størrelse (udpakket): {len(xml_data)//1024} KB")


if __name__ == "__main__":
    main()
