import gzip
import requests
from pathlib import Path

URL = "https://epgshare01.online/epgshare01/epg_ripper_DK1.xml.gz"

out_dir = Path("output_epgshare")
out_dir.mkdir(exist_ok=True)

xml_file = out_dir / "epgshare_dk1.xml"

print("Downloader EPGShare...")

r = requests.get(URL, timeout=300)
r.raise_for_status()

xml_data = gzip.decompress(r.content)

xml_file.write_bytes(xml_data)

print(f"Gemt: {xml_file}")
print(f"Størrelse: {xml_file.stat().st_size:,} bytes")