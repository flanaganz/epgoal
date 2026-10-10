#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sikker TV3/TV3+ mapping-fixer til channel_priority_v2.xlsx.

Standard er tørkørsel. Brug --apply for at skrive en ny rettet workbook.
Originalfilen overskrives aldrig.
"""
from __future__ import annotations
import argparse, json, re, shutil
from copy import copy
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_INPUT=Path(r"C:\EPGoal\data\channel_priority_v2.xlsx")
TARGETS={"tv3.dk","tv3+.dk"}

def split(v): return [x.strip() for x in re.split(r"[,;\n]+",str(v or "")) if x.strip()]
def is_tv3(name): return str(name or "").strip().lower()=="tv3.dk"
def is_tv3plus(name): return str(name or "").strip().lower()=="tv3+.dk"
def stream_is_plus(name): return "tv3+" in str(name or "").lower()
def stream_is_plain(name): return bool(re.search(r"\btv3\b",str(name or "").lower())) and not stream_is_plus(name) and not re.search(r"\b(max|puls|sport)\b",str(name or "").lower())
def join(vals): return ", ".join(vals)

def find_header(ws):
    for r in range(1,min(ws.max_row,30)+1):
        if str(ws.cell(r,1).value or "").strip()=="Kanal":return r
    raise ValueError("Kolonnen 'Kanal' blev ikke fundet")
def copy_style(src,dst):
    if src.has_style:
        dst._style=copy(src._style);dst.number_format=src.number_format
    if src.hyperlink:dst._hyperlink=copy(src.hyperlink)
    if src.comment:dst.comment=copy(src.comment)

def main():
    ap=argparse.ArgumentParser(description="Ret TV3 og TV3+ mappings sikkert")
    ap.add_argument("--input",type=Path,default=DEFAULT_INPUT)
    ap.add_argument("--output",type=Path,default=Path(r"C:\EPGoal\data\channel_priority_v2.xlsx"))
    ap.add_argument("--report-dir",type=Path,default=Path(r"C:\EPGoal\reports"))
    ap.add_argument("--apply",action="store_true",help="Skriv rettet kopi. Uden flag laves kun rapport.")
    args=ap.parse_args()
    if not args.input.exists(): raise SystemExit(f"FEJL: Filen findes ikke: {args.input}")
    args.report_dir.mkdir(parents=True,exist_ok=True)
    wb=load_workbook(args.input);ws=wb[wb.sheetnames[0]];hr=find_header(ws)
    headers={str(ws.cell(hr,c).value or "").strip():c for c in range(1,ws.max_column+1)}
    required=["Kanal","Output/UHF tvg-id","BSS M3U tvg-id'er","BSS M3U stream-navne"]
    missing=[x for x in required if x not in headers]
    if missing: raise SystemExit("FEJL: Manglende kolonner: "+", ".join(missing))
    rows=[]
    for r in range(hr+1,ws.max_row+1):
        name=str(ws.cell(r,headers["Kanal"]).value or "").strip()
        if name.lower() in TARGETS: rows.append((r,name))
    if len(rows)!=2: raise SystemExit(f"FEJL: Forventede præcis 2 rækker (TV3.dk og TV3+.dk), fandt {len(rows)}")

    changes=[];proposals=[]
    for r,name in rows:
        names=split(ws.cell(r,headers["BSS M3U stream-navne"]).value)
        ids=split(ws.cell(r,headers["BSS M3U tvg-id'er"]).value)
        wanted_names=[n for n in names if (stream_is_plain(n) if is_tv3(name) else stream_is_plus(n))]
        wanted_id="tv3.dk" if is_tv3(name) else "tv3plus.dk"
        proposed={
            "Output/UHF tvg-id":"TV3.dk" if is_tv3(name) else "TV3+.dk",
            "BSS M3U tvg-id'er":wanted_id,
            "BSS M3U stream-navne":join(wanted_names),
        }
        # Sørg for entydige kilde-ID'er, men behold alle øvrige kildevarianter.
        for col in ("EPGShare ID'er","OpenEPG ID'er","BSS XMLTV ID'er"):
            if col not in headers: continue
            old=split(ws.cell(r,headers[col]).value)
            if is_tv3(name): new=[x for x in old if "+" not in x and "plus" not in x.lower()]
            else: new=[x for x in old if ("+" in x or "plus" in x.lower())]
            # Undgå at tømme kolonnen hvis den kun indeholder et fælles/uklart ID.
            if new: proposed[col]=join(dict.fromkeys(new))
        for col,new in proposed.items():
            c=headers[col];old=ws.cell(r,c).value
            if str(old or "")!=str(new or ""):
                changes.append([name,col,old,new,"SKRIVES" if args.apply else "FORSLAG"])
                if args.apply: ws.cell(r,c).value=new
        proposals.append({"row":r,"channel":name,"proposed":proposed,"kept_streams":wanted_names,"removed_streams":[n for n in names if n not in wanted_names]})

    # Rapportworkbook
    rw=Workbook();rs=rw.active;rs.title="TV3-TV3plus Fix"
    hdr=["Kanal","Kolonne","Før","Efter","Status"];rs.append(hdr)
    for x in changes:rs.append(x)
    fill=PatternFill("solid",fgColor="1F4E78")
    for c in rs[1]:c.font=Font(color="FFFFFF",bold=True);c.fill=fill
    rs.freeze_panes="A2";rs.auto_filter.ref=rs.dimensions
    for c in range(1,6):rs.column_dimensions[get_column_letter(c)].width=[18,40,70,70,14][c-1]
    for row in rs.iter_rows(min_row=2):
        for c in row:c.alignment=Alignment(vertical="top",wrap_text=True)
    ss=rw.create_sheet("Summary")
    for row in [["Kørt",datetime.now().strftime("%Y-%m-%d %H:%M:%S")],["Input",str(args.input)],["Mode","APPLY" if args.apply else "DRY-RUN"],["Ændringer",len(changes)]]:ss.append(row)
    report=args.report_dir/"tv3_tv3plus_fix_report.xlsx";rw.save(report)
    json_path=args.report_dir/"tv3_tv3plus_fix_report.json"
    json_path.write_text(json.dumps({"generated":datetime.now().isoformat(timespec="seconds"),"mode":"apply" if args.apply else "dry-run","changes":[{"channel":x[0],"column":x[1],"before":x[2],"after":x[3]} for x in changes],"proposals":proposals},ensure_ascii=False,indent=2),encoding="utf-8")

    if args.apply:
        args.output.parent.mkdir(parents=True,exist_ok=True)
        backup = args.report_dir / (
            "channel_priority_v2_backup_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )   
        shutil.copy2(args.input,backup)
        wb.save(args.output)
        # Genåbn og kontrollér de to mål-rækker.
        check=load_workbook(args.output,data_only=False);cws=check[check.sheetnames[0]];chr=find_header(cws);ch={str(cws.cell(chr,c).value or "").strip():c for c in range(1,cws.max_column+1)}
        found={str(cws.cell(r,ch["Kanal"]).value or "").strip():str(cws.cell(r,ch["Output/UHF tvg-id"]).value or "").strip() for r in range(chr+1,cws.max_row+1) if str(cws.cell(r,ch["Kanal"]).value or "").strip().lower() in TARGETS}
        if found.get("TV3.dk")!="TV3.dk" or found.get("TV3+.dk")!="TV3+.dk": raise SystemExit("FEJL: Efterkontrol af output-ID'er fejlede")
        print(f"Rettet kopi : {args.output}\nBackup       : {backup}")
    else:
        print("DRY-RUN: Ingen kildefil blev ændret. Kør igen med --apply for at skrive en rettet kopi.")
    print(f"Rapport      : {report}\nJSON         : {json_path}\nÆndringer    : {len(changes)}")
if __name__=="__main__":main()
