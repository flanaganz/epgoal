#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Integrer EPGoal Console med produktionspipelinen sikkert."""
from pathlib import Path
from datetime import datetime
import ast, shutil

ROOT=Path(r"C:\EPGoal")
TARGET=ROOT/'scripts'/'opdater_epgoal.py'
RUNNER=ROOT/'EPGoalConsole'/'core'/'runner.py'
REPORTS=ROOT/'reports'

CONST_OLD='''GIT_COMMIT_PREFIX = "Auto-opdater EPGOAL"\n'''
CONST_NEW='''GIT_COMMIT_PREFIX = "Auto-opdater EPGOAL"\nCONSOLE_CONFIG_DIR = ROOT / "EPGoalConsole" / "config"\nSOURCE_OVERRIDES_FILE = CONSOLE_CONFIG_DIR / "source_overrides.json"\n'''

SPAN_BLOCK='''        epgshare_span = max(\n            (epgshare_stats.get(c, {}).get("span_hours", 0) for c in matched_epgshare_ids),\n            default=0,\n        )\n        if epgshare_count > 0 and epgshare_span >= MIN_WINDOW_HOURS:\n'''

OVERRIDE_BLOCK='''        epgshare_span = max(\n            (epgshare_stats.get(c, {}).get("span_hours", 0) for c in matched_epgshare_ids),\n            default=0,\n        )\n\n        # Kildevalg fra EPGoal Console. Uden en aktiv override fortsætter den\n        # eksisterende EPGShare -> OpenEPG -> BSS-prioritet uændret.\n        source_overrides = load_json(SOURCE_OVERRIDES_FILE, {})\n        override = source_overrides.get(kanonisk_navn, {})\n        override_enabled = bool(override.get("enabled", True))\n        override_source = str(override.get("source", "auto")).strip().lower()\n        override_id = str(override.get("source_channel_id", "")).strip()\n\n        if override_enabled and override_source == "exclude":\n            for cid in matched_epgshare_ids:\n                for p in list(epgshare_programmes_by_channel.get(cid, [])):\n                    if p in root:\n                        root.remove(p)\n                for ch in list(root.findall("channel")):\n                    if ch.get("id", "") == cid:\n                        root.remove(ch)\n            report.setdefault("excluded_by_console", []).append({\n                "kanal": kanonisk_navn, "kildeoverride": "exclude"\n            })\n            print(f"Console-override: udelader {kanonisk_navn}")\n            continue\n\n        forced = None\n        if override_enabled and override_source in {"epgshare", "openepg", "bss"}:\n            if override_source == "epgshare":\n                source_map, source_stats = epgshare_programmes_by_channel, epgshare_stats\n                candidate_ids = matched_epgshare_ids\n            elif override_source == "openepg":\n                source_map, source_stats = openepg_programmes_by_channel, openepg_stats\n                candidate_ids = list(openepg_programmes_by_channel)\n            else:\n                source_map, source_stats = bss_programmes_by_channel, bss_stats\n                candidate_ids = list(bss_programmes_by_channel)\n\n            selected_id = None\n            if override_id and override_id in source_map:\n                selected_id = override_id\n            elif override_id:\n                selected_id = next((cid for cid in candidate_ids if normalize_id(cid) == normalize_id(override_id)), None)\n            if selected_id is None and override_source == "epgshare" and matched_epgshare_ids:\n                selected_id = max(matched_epgshare_ids, key=lambda c: source_stats.get(c, {}).get("count", 0))\n            if selected_id is None and not override_id:\n                selected_id, _ = find_best_candidate(tokens, source_map, source_stats)\n\n            selected_stat = source_stats.get(selected_id, {"count": 0, "span_hours": 0}) if selected_id else {"count": 0, "span_hours": 0}\n            if selected_id and selected_stat["count"] > 0:\n                forced = (override_source, selected_id, selected_stat, source_map)\n            else:\n                print(\n                    f"ADVARSEL: Console-override for {kanonisk_navn} kunne ikke finde "\n                    f"{override_source}:{override_id or '(automatisk ID)'}. Bruger normal prioritet.",\n                    file=sys.stderr,\n                )\n\n        if forced:\n            best_source, best_id, best_stat, best_programmes_map = forced\n            if best_source == "epgshare":\n                report["kept_as_is"].append({\n                    "kanal": kanonisk_navn, "kilde": "epgshare",\n                    "programmer": best_stat["count"],\n                    "vindue_timer": round(best_stat["span_hours"], 1),\n                    "tvunget_kildeoverride": True, "kilde_channel_id": best_id,\n                })\n                print(f"Console-override: {kanonisk_navn} -> epgshare:{best_id}")\n                continue\n\n            if matched_epgshare_ids:\n                primary_id = max(matched_epgshare_ids, key=lambda c: epgshare_stats.get(c, {}).get("count", 0))\n            else:\n                primary_id = f"{normalize_id(kanonisk_navn)}.dk"\n                new_channel_el = ET.Element("channel", {"id": primary_id})\n                dn = ET.SubElement(new_channel_el, "display-name")\n                dn.text = kanonisk_navn\n                root.append(new_channel_el)\n            for cid in matched_epgshare_ids:\n                for p in list(epgshare_programmes_by_channel.get(cid, [])):\n                    if p in root:\n                        root.remove(p)\n            inserted = 0\n            for p in best_programmes_map.get(best_id, []):\n                clone = copy.deepcopy(p)\n                clone.set("channel", primary_id)\n                strip_artwork(clone)\n                root.append(clone)\n                inserted += 1\n            report["supplemented"].append({\n                "kanal": kanonisk_navn,\n                "epgshare_programmer_foer": epgshare_count,\n                "epgshare_vindue_timer_foer": round(epgshare_span, 1),\n                "kilde_brugt": best_source, "kilde_channel_id": best_id,\n                "programmer_indsat": inserted,\n                "vindue_timer_efter": round(best_stat["span_hours"], 1),\n                "primaer_channel_id": primary_id,\n                "tvunget_kildeoverride": True,\n            })\n            print(f"Console-override: {kanonisk_navn} -> {best_source}:{best_id} ({inserted:,} programmer)")\n            continue\n\n        if epgshare_count > 0 and epgshare_span >= MIN_WINDOW_HOURS:\n'''

GIT_OLD='''def step7_git_push() -> None:\n    section("TRIN 7/8: Git commit + push")\n    if not GIT_ENABLED:\n'''
GIT_NEW='''def step7_git_push() -> None:\n    section("TRIN 7/8: Git commit + push")\n    if os.environ.get("EPGOAL_SKIP_GIT", "").strip().lower() in {"1", "true", "yes"}:\n        print("Git springes over: kørslen blev startet som Byg og valider fra EPGoal Console.")\n        return\n    if not GIT_ENABLED:\n'''

OPEN_OLD='''def open_file_for_review(path: Path) -> None:\n    if not path.exists():\n'''
OPEN_NEW='''def open_file_for_review(path: Path) -> None:\n    if os.environ.get("EPGOAL_CONSOLE_MODE", "").strip().lower() in {"1", "true", "yes"}:\n        print(f"Console-tilstand: åbner ikke {path.name} automatisk.")\n        return\n    if not path.exists():\n'''

RUN_OLD='''    p=subprocess.Popen([sys.executable,str(target)],cwd=paths()['root'],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace',bufsize=1)'''
RUN_NEW='''    import os\n    env = os.environ.copy()\n    env["EPGOAL_CONSOLE_MODE"] = "1"\n    env["EPGOAL_SKIP_GIT"] = "1"\n    p=subprocess.Popen([sys.executable,str(target)],cwd=paths()['root'],env=env,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace',bufsize=1)'''

def replace_once(s,old,new,label):
    if new in s: return s,False
    if old not in s: raise SystemExit(f'FEJL: Kunne ikke finde blokken: {label}. Ingen filer blev ændret.')
    return s.replace(old,new,1),True

def main():
    if not TARGET.exists() or not RUNNER.exists(): raise SystemExit('FEJL: Produktionsscript eller Console runner mangler.')
    source=TARGET.read_text(encoding='utf-8'); runner=RUNNER.read_text(encoding='utf-8')
    changes=[]
    source,ch=replace_once(source,CONST_OLD,CONST_NEW,'konstanter');changes.append(ch)
    source,ch=replace_once(source,SPAN_BLOCK,OVERRIDE_BLOCK,'merge override');changes.append(ch)
    source,ch=replace_once(source,GIT_OLD,GIT_NEW,'Git skip');changes.append(ch)
    source,ch=replace_once(source,OPEN_OLD,OPEN_NEW,'review-filer');changes.append(ch)
    runner,ch=replace_once(runner,RUN_OLD,RUN_NEW,'Console runner miljø');changes.append(ch)
    ast.parse(source);ast.parse(runner)
    REPORTS.mkdir(parents=True,exist_ok=True);stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
    b1=REPORTS/f'opdater_epgoal_before_console_integration_{stamp}.py';b2=REPORTS/f'console_runner_before_integration_{stamp}.py'
    shutil.copy2(TARGET,b1);shutil.copy2(RUNNER,b2)
    TARGET.write_text(source,encoding='utf-8');RUNNER.write_text(runner,encoding='utf-8')
    ast.parse(TARGET.read_text(encoding='utf-8'));ast.parse(RUNNER.read_text(encoding='utf-8'))
    print('=== EPGOAL CONSOLE INTEGRATION GENNEMFØRT ===')
    print(f'Produktionsscript: {TARGET}')
    print(f'Console runner   : {RUNNER}')
    print(f'Backup 1         : {b1}')
    print(f'Backup 2         : {b2}')
    print('Console-kørsler bygger og validerer nu uden Git-push og uden at åbne reviewfiler.')
    print('Manuel kørsel af opdater_epgoal.py beholder Git-push og åbning af reviewfiler.')

if __name__=='__main__':main()
