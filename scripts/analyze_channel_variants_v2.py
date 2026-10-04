#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Analyze all XMLTV/M3U/XLSX files in C:\\EPGoal\\output_epgshare.

Creates an Excel report comparing Normal, HD and FHD/UHD stream IDs against
all discovered XMLTV channel IDs. The script only reports and suggests matches;
it does not modify source files or epgoal.xml.
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional
import xml.etree.ElementTree as ET

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("FEJL: openpyxl mangler. Kør: py -m pip install openpyxl", file=sys.stderr)
    raise

DEFAULT_FOLDER = Path(r"C:\EPGoal\output_epgshare")
DEFAULT_REPORT = DEFAULT_FOLDER / "channel_variant_analysis_v2.xlsx"
CHECK_CANDIDATES = [
    Path(r"C:\EPGoal\EPGOAL_Channel_Ckeck.xlsx"),
    Path(r"C:\EPGoal\output_epgshare\EPGOAL_Channel_Ckeck.xlsx"),
    Path(r"C:\EPGoal\EPGOAL_Channel_Check.xlsx"),
]
VARIANT_ORDER = {"NORMAL": 0, "HD": 1, "FHD": 2, "UHD": 3}
VIDEO_WORDS = re.compile(r"(?i)(?:^|[\s._+()\-])(4k|uhd|fhd|full\s*hd|hd|sd)(?=$|[\s._+()\-])")
COUNTRY_WORDS = re.compile(r"(?i)(?:^|[\s._+()\-])(dk|da|denmark|danmark)(?=$|[\s._+()\-])")

@dataclass(frozen=True)
class XmlChannel:
    source: str
    channel_id: str
    display_name: str
    programme_count: int
    base_key: str
    variant: str

@dataclass(frozen=True)
class Stream:
    source: str
    tvg_id: str
    tvg_name: str
    group_title: str
    base_key: str
    variant: str


def clean_text(value: object) -> str:
    return " ".join(str(value or "").strip().split())


def ascii_fold(value: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", value) if not unicodedata.combining(c))


def detect_variant(*values: str) -> str:
    text = " ".join(clean_text(v) for v in values).lower()
    if re.search(r"(?:^|[\s._+()\-])(4k|uhd)(?=$|[\s._+()\-])", text): return "UHD"
    if re.search(r"(?:^|[\s._+()\-])(fhd|full\s*hd)(?=$|[\s._+()\-])", text): return "FHD"
    if re.search(r"(?:^|[\s._+()\-])hd(?=$|[\s._+()\-])", text): return "HD"
    return "NORMAL"


def base_key(*values: str) -> str:
    # Prefer the human-readable name, falling back to ID.
    text = next((clean_text(v) for v in values if clean_text(v)), "")
    text = ascii_fold(text).lower()
    text = VIDEO_WORDS.sub(" ", text)
    text = COUNTRY_WORDS.sub(" ", text)
    text = re.sub(r"\([^)]*(?:dk|da|denmark|danmark|hd|fhd|uhd|4k)[^)]*\)", " ", text, flags=re.I)
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def display_local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def iter_xml_files(folder: Path) -> Iterable[Path]:
    for p in folder.rglob("*"):
        if p.is_file() and p.suffix.lower() in {".xml", ".xmltv"}:
            yield p


def read_xmltv(path: Path) -> tuple[list[XmlChannel], Optional[str]]:
    try:
        counts: Counter[str] = Counter()
        meta: dict[str, str] = {}
        # Two passes keep memory low, also for very large BSS/OpenEPG files.
        for _, el in ET.iterparse(path, events=("end",)):
            tag = display_local(el.tag)
            if tag == "programme":
                cid = clean_text(el.attrib.get("channel"))
                if cid: counts[cid] += 1
            el.clear()
        for _, el in ET.iterparse(path, events=("end",)):
            if display_local(el.tag) == "channel":
                cid = clean_text(el.attrib.get("id"))
                names = [clean_text(c.text) for c in list(el) if display_local(c.tag) == "display-name" and clean_text(c.text)]
                if cid: meta[cid] = names[0] if names else cid
            el.clear()
        rows = []
        for cid, name in meta.items():
            rows.append(XmlChannel(path.name, cid, name, counts.get(cid, 0), base_key(name, cid), detect_variant(name, cid)))
        return rows, None
    except Exception as exc:
        return [], f"{path}: {type(exc).__name__}: {exc}"


ATTR_RE = re.compile(r'''([\w-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')''')
def read_m3u(path: Path) -> tuple[list[Stream], Optional[str]]:
    try:
        out, pending = [], None
        with path.open("r", encoding="utf-8-sig", errors="replace") as fh:
            for raw in fh:
                line = raw.strip()
                if line.startswith("#EXTINF"):
                    attrs = {m.group(1).lower(): (m.group(2) if m.group(2) is not None else m.group(3)) for m in ATTR_RE.finditer(line)}
                    name = clean_text(line.split(",", 1)[1] if "," in line else attrs.get("tvg-name"))
                    tid = clean_text(attrs.get("tvg-id"))
                    tname = clean_text(attrs.get("tvg-name") or name)
                    group = clean_text(attrs.get("group-title"))
                    pending = Stream(path.name, tid, tname, group, base_key(tname, name, tid), detect_variant(tname, name, tid))
                elif line and not line.startswith("#") and pending:
                    out.append(pending); pending = None
        return out, None
    except Exception as exc:
        return [], f"{path}: {type(exc).__name__}: {exc}"


def find_header(headers: list[str], terms: tuple[str, ...]) -> Optional[int]:
    normalized = [ascii_fold(h).lower().replace(" ", "") for h in headers]
    for i, h in enumerate(normalized):
        if any(t in h for t in terms): return i
    return None


def read_xlsx_streams(path: Path) -> tuple[list[Stream], list[tuple[str, str]], Optional[str]]:
    """Read likely mapping sheets and free-form channel check notes."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        streams, notes = [], []
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows: continue
            headers = [clean_text(v) for v in rows[0]]
            id_i = find_header(headers, ("tvg-id", "tvgid", "output/uhftvg-id", "uhftvg-id", "channel-id", "channelid"))
            name_i = find_header(headers, ("tvg-name", "tvgname", "kanal", "channel", "navn", "name"))
            var_i = find_header(headers, ("variant", "quality", "kvalitet"))
            if id_i is not None or name_i is not None:
                for row in rows[1:]:
                    cid = clean_text(row[id_i]) if id_i is not None and id_i < len(row) else ""
                    name = clean_text(row[name_i]) if name_i is not None and name_i < len(row) else ""
                    var = clean_text(row[var_i]) if var_i is not None and var_i < len(row) else ""
                    if cid or name:
                        streams.append(Stream(path.name, cid, name or cid, ws.title, base_key(name, cid), detect_variant(var, name, cid)))
            else:
                # The attached check file is two free-form columns without headers.
                for row in rows:
                    vals = [clean_text(v) for v in row if clean_text(v)]
                    if vals and re.search(r"(?i)\.dk$", vals[0]):
                        notes.append((base_key(vals[0]), " | ".join(vals[1:])))
        return streams, notes, None
    except Exception as exc:
        return [], [], f"{path}: {type(exc).__name__}: {exc}"


def score_candidate(stream: Stream, ch: XmlChannel) -> int:
    if not stream.base_key or stream.base_key != ch.base_key: return -1
    score = 100
    if stream.variant == ch.variant: score += 25
    if stream.tvg_id and stream.tvg_id.lower() == ch.channel_id.lower(): score += 100
    if stream.tvg_name and stream.tvg_name.lower() == ch.display_name.lower(): score += 30
    if ch.programme_count > 0: score += 10
    return score


def autosize(ws, max_width=60):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        width = min(max((len(clean_text(ws.cell(r, col).value)) for r in range(1, min(ws.max_row, 5000) + 1)), default=8) + 2, max_width)
        ws.column_dimensions[get_column_letter(col)].width = max(10, width)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def add_sheet(wb, title: str, headers: list[str], rows: list[list[object]]):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for row in rows: ws.append(row)
    autosize(ws)
    if ws.max_row > 1:
        tab = Table(displayName=re.sub(r"\W", "", title)[:25] + "Tbl", ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False)
        ws.add_table(tab)
    return ws


def main() -> int:
    ap = argparse.ArgumentParser(description="Sammenlign Normal/HD/FHD/UHD ID'er i hele output_epgshare")
    ap.add_argument("--folder", type=Path, default=DEFAULT_FOLDER)
    ap.add_argument("--output", type=Path, default=DEFAULT_REPORT)
    ap.add_argument("--check", type=Path, help="Valgfri EPGOAL_Channel_Ckeck.xlsx")
    args = ap.parse_args()
    folder = args.folder.resolve()
    if not folder.exists():
        print(f"FEJL: Mappen findes ikke: {folder}", file=sys.stderr); return 2

    xml_channels: list[XmlChannel] = []
    streams: list[Stream] = []
    notes: list[tuple[str, str]] = []
    errors: list[str] = []
    inventory: list[list[object]] = []

    xml_paths = sorted(iter_xml_files(folder))
    m3u_paths = sorted(p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() in {".m3u", ".m3u8"})
    xlsx_paths = sorted(p for p in folder.rglob("*.xlsx") if p.resolve() != args.output.resolve() and not p.name.startswith("~$"))
    check = args.check
    if not check:
        check = next((p for p in CHECK_CANDIDATES if p.exists()), None)
    if check and check.exists() and check.resolve() not in [p.resolve() for p in xlsx_paths]: xlsx_paths.append(check)

    for path in xml_paths:
        rows, err = read_xmltv(path); xml_channels.extend(rows)
        inventory.append([str(path.relative_to(folder)), "XMLTV", len(rows), sum(r.programme_count for r in rows), err or "OK"])
        if err: errors.append(err)
    for path in m3u_paths:
        rows, err = read_m3u(path); streams.extend(rows)
        inventory.append([str(path.relative_to(folder)), "M3U", len(rows), "", err or "OK"])
        if err: errors.append(err)
    for path in xlsx_paths:
        rows, n, err = read_xlsx_streams(path); streams.extend(rows); notes.extend(n)
        label = str(path.relative_to(folder)) if path.is_relative_to(folder) else str(path)
        inventory.append([label, "XLSX", len(rows), "", err or "OK"])
        if err: errors.append(err)

    by_base: dict[str, list[XmlChannel]] = defaultdict(list)
    for ch in xml_channels:
        if ch.base_key: by_base[ch.base_key].append(ch)
    note_map = defaultdict(list)
    for key, note in notes:
        if note: note_map[key].append(note)

    # De-duplicate stream rows across repeated mapping workbooks.
    unique_streams = list({(s.source, s.tvg_id, s.tvg_name, s.group_title, s.base_key, s.variant): s for s in streams}.values())
    comparison = []
    candidates_rows = []
    for s in sorted(unique_streams, key=lambda x: (x.base_key, VARIANT_ORDER.get(x.variant, 9), x.tvg_name, x.tvg_id)):
        candidates = sorted(((score_candidate(s, ch), ch) for ch in by_base.get(s.base_key, [])), key=lambda x: (-x[0], -x[1].programme_count, x[1].source, x[1].channel_id))
        best_score, best = candidates[0] if candidates else (-1, None)
        exact = next((ch for ch in xml_channels if s.tvg_id and ch.channel_id.lower() == s.tvg_id.lower()), None)
        matched = exact or best
        status = "MATCH" if exact else ("FORSLAG" if matched else "MANGLER")
        epg = "JA" if matched and matched.programme_count > 0 else "NEJ"
        comparison.append([
            s.base_key, s.variant, s.tvg_name, s.tvg_id, s.source, s.group_title,
            status, matched.channel_id if matched else "", matched.display_name if matched else "",
            matched.variant if matched else "", matched.source if matched else "",
            matched.programme_count if matched else 0, epg, best_score if matched else "",
            " || ".join(dict.fromkeys(note_map.get(s.base_key, [])))
        ])
        for score, ch in candidates[:10]:
            candidates_rows.append([s.base_key, s.variant, s.tvg_name, s.tvg_id, ch.channel_id, ch.display_name, ch.variant, ch.source, ch.programme_count, score])

    # Variant matrix: one row per normalized channel family.
    matrix = []
    all_keys = sorted(set(by_base) | {s.base_key for s in unique_streams if s.base_key})
    for key in all_keys:
        members = [r for r in comparison if r[0] == key]
        xmls = by_base.get(key, [])
        row = [key]
        for variant in ("NORMAL", "HD", "FHD", "UHD"):
            ss = [r for r in members if r[1] == variant]
            xs = [x for x in xmls if x.variant == variant]
            row += [" | ".join(dict.fromkeys(r[3] for r in ss if r[3])), " | ".join(dict.fromkeys(x.channel_id for x in xs)), sum(x.programme_count for x in xs)]
        expected = {r[1] for r in members}
        covered = {r[1] for r in members if r[12] == "JA"}
        missing = sorted(expected - covered, key=lambda v: VARIANT_ORDER.get(v, 9))
        matched_vars = sorted(expected & covered, key=lambda v: VARIANT_ORDER.get(v, 9))
        row += [", ".join(matched_vars), ", ".join(missing), " || ".join(dict.fromkeys(note_map.get(key, [])))]
        matrix.append(row)

    wb = Workbook(); wb.remove(wb.active)
    summary = wb.create_sheet("Oversigt")
    summary_rows = [
        ("Mappe", str(folder)), ("XML/XMLTV-filer", len(xml_paths)), ("M3U/M3U8-filer", len(m3u_paths)),
        ("Excel-filer", len(xlsx_paths)), ("XML channel-id'er", len(xml_channels)), ("Stream/mapping-rækker", len(unique_streams)),
        ("Match", sum(1 for r in comparison if r[6] == "MATCH")), ("Forslag", sum(1 for r in comparison if r[6] == "FORSLAG")),
        ("Mangler", sum(1 for r in comparison if r[6] == "MANGLER")), ("Streams uden EPG", sum(1 for r in comparison if r[12] == "NEJ")),
        ("Fejl under læsning", len(errors)),
    ]
    summary.append(["Nøgle", "Værdi"])
    for r in summary_rows: summary.append(r)
    autosize(summary)

    headers = ["Base key", "Variant", "Streamnavn", "Stream tvg-id", "Streamkilde", "Gruppe", "Status", "Bedste XML channel-id", "XML display-name", "XML variant", "XML-kilde", "Programmer", "EPG", "Matchscore", "Manuel note"]
    ws_all = add_sheet(wb, "Alle_ID_sammenligninger", headers, comparison)
    ws_missing = add_sheet(wb, "Manglende_IDer", headers, [r for r in comparison if r[6] != "MATCH" or r[12] == "NEJ"])
    ws_matching = add_sheet(wb, "Matchende_IDer", headers, [r for r in comparison if r[6] == "MATCH" and r[12] == "JA"])
    matrix_headers = ["Base key", "Normal stream-id", "Normal XML-id", "Normal programmer", "HD stream-id", "HD XML-id", "HD programmer", "FHD stream-id", "FHD XML-id", "FHD programmer", "UHD stream-id", "UHD XML-id", "UHD programmer", "Matchende varianter", "Manglende varianter", "Manuel note"]
    add_sheet(wb, "Variant_matrix", matrix_headers, matrix)
    add_sheet(wb, "Match_kandidater", ["Base key", "Stream variant", "Streamnavn", "Stream tvg-id", "XML channel-id", "XML display-name", "XML variant", "XML-kilde", "Programmer", "Score"], candidates_rows)
    add_sheet(wb, "XML_kanaler", ["Kilde", "Channel-id", "Display-name", "Programmer", "Base key", "Variant"], [[x.source, x.channel_id, x.display_name, x.programme_count, x.base_key, x.variant] for x in xml_channels])
    add_sheet(wb, "Filoversigt", ["Fil", "Type", "Rækker/kanaler", "Programmer", "Status"], inventory)
    add_sheet(wb, "Fejl", ["Fejl"], [[e] for e in errors] or [["Ingen fejl"]])

    # Visual flags.
    red = PatternFill("solid", fgColor="FFC7CE"); orange = PatternFill("solid", fgColor="FCE4D6"); green = PatternFill("solid", fgColor="C6EFCE")
    for ws in (ws_all, ws_missing, ws_matching):
        for row in range(2, ws.max_row + 1):
            status = ws.cell(row, 7).value; epg = ws.cell(row, 13).value
            fill = green if status == "MATCH" and epg == "JA" else (orange if status == "FORSLAG" else red)
            for col in range(1, ws.max_column + 1): ws.cell(row, col).fill = fill
    args.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    print("=== analyze_channel_variants_v2 ===")
    print(f"Mappe       : {folder}")
    print(f"XML-kanaler : {len(xml_channels):,}")
    print(f"Streams     : {len(unique_streams):,}")
    print(f"Match       : {sum(1 for r in comparison if r[6] == 'MATCH'):,}")
    print(f"Forslag     : {sum(1 for r in comparison if r[6] == 'FORSLAG'):,}")
    print(f"Mangler     : {sum(1 for r in comparison if r[6] == 'MANGLER'):,}")
    print(f"Rapport     : {args.output}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
