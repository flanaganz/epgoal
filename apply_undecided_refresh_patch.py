#!/usr/bin/env python3
from pathlib import Path
import shutil, sys, ast
root=Path(__file__).resolve().parent
target=root/'scripts'/'danish_backdrops.py'
backup=root/'scripts'/'danish_backdrops.before_undecided_refresh.py'
if not target.exists(): sys.exit(f'Filen findes ikke: {target}')
s=target.read_text(encoding='utf-8')
shutil.copy2(target, backup)
old='    return approved\n\n\ndef tmdb_search(title: str) -> tuple[str, int] | None:'
new='    return approved\n\n\ndef load_undecided_review_keys(review_path: Path) -> dict[str, str]:\n    """Rækker uden X i både Godkendt og Ignorer; opfriskes frisk."""\n    if not review_path.exists():\n        return {}\n    try:\n        from openpyxl import load_workbook\n    except ImportError:\n        return {}\n    wb = load_workbook(review_path, data_only=True)\n    ws = wb.active\n    headers = [c.value for c in ws[1]]\n    try:\n        key_col = headers.index("Nøgle (intern)")\n        approved_col = headers.index("Godkendt (X)")\n    except ValueError:\n        return {}\n    ignored_col = headers.index("Ignorer (X)") if "Ignorer (X)" in headers else None\n    title_col = headers.index("Titel") if "Titel" in headers else None\n    result = {}\n    for row in ws.iter_rows(min_row=2):\n        key = row[key_col].value\n        if not key:\n            continue\n        approved = str(row[approved_col].value or "").strip().upper() == "X"\n        ignored = (str(row[ignored_col].value or "").strip().upper() == "X"\n                   if ignored_col is not None else False)\n        if approved or ignored:\n            continue\n        norm = normalize_title(str(key))\n        title = row[title_col].value if title_col is not None else key\n        result[norm] = str(title).strip()\n    return result\n\n\ndef tmdb_search(title: str) -> tuple[str, int] | None:'

if old not in s: sys.exit('Kunne ikke finde kodeblok 1; backup: '+str(backup))
s=s.replace(old,new,1)
old='                      refresh_titles_normalized: set[str],\n                      refresh_seen: set[str]) -> dict:'
new='                      refresh_titles_normalized: set[str],\n                      refresh_seen: set[str],\n                      force_refreshed_this_run: set[str]) -> dict:'

if old not in s: sys.exit('Kunne ikke finde kodeblok 2; backup: '+str(backup))
s=s.replace(old,new,1)
old='            force_refresh = norm in refresh_titles_normalized\n            was_cached_not_found = ('
new='            force_refresh = (norm in refresh_titles_normalized\n                             and norm not in force_refreshed_this_run)\n            was_cached_not_found = ('

if old not in s: sys.exit('Kunne ikke finde kodeblok 3; backup: '+str(backup))
s=s.replace(old,new,1)
old='            if force_refresh:\n                stats["force_refreshed"] += 1'
new='            if force_refresh:\n                force_refreshed_this_run.add(norm)\n                stats["force_refreshed"] += 1'

if old not in s: sys.exit('Kunne ikke finde kodeblok 4; backup: '+str(backup))
s=s.replace(old,new,1)
old='    refresh_titles_raw = load_refresh_titles(GENOPFRISK_TITLER_FILE)\n    refresh_titles_normalized = {normalize_title(t) for t in refresh_titles_raw}\n    refresh_seen: set[str] = set()'
new='    refresh_titles_raw = load_refresh_titles(GENOPFRISK_TITLER_FILE)\n    undecided_review_keys = load_undecided_review_keys(DANISH_ARTWORK_REVIEW_FILE)\n    refresh_titles_normalized = ({normalize_title(t) for t in refresh_titles_raw}\n                                 | set(undecided_review_keys))\n    refresh_seen: set[str] = set()\n    force_refreshed_this_run: set[str] = set()'

if old not in s: sys.exit('Kunne ikke finde kodeblok 5; backup: '+str(backup))
s=s.replace(old,new,1)
old='    if refresh_titles_raw:\n        print(f"🔄 Tvangsopfrisker {len(refresh_titles_raw):,} titel(r) fra {GENOPFRISK_TITLER_FILE.name} "\n              "(uanset cache-alder) ...")'
new='    if refresh_titles_raw:\n        print(f"🔄 Tvangsopfrisker {len(refresh_titles_raw):,} titel(r) fra {GENOPFRISK_TITLER_FILE.name} "\n              "(uanset cache-alder) ...")\n    if undecided_review_keys:\n        print(f"🔄 {len(undecided_review_keys):,} uafklarede titel(r) i "\n              f"{DANISH_ARTWORK_REVIEW_FILE.name} får friske TMDb-opslag ...")'

if old not in s: sys.exit('Kunne ikke finde kodeblok 6; backup: '+str(backup))
s=s.replace(old,new,1)
old='            manual_index, manual_titles_matched,\n            refresh_titles_normalized, refresh_seen,\n        )'
new='            manual_index, manual_titles_matched,\n            refresh_titles_normalized, refresh_seen, force_refreshed_this_run,\n        )'

if old not in s: sys.exit('Kunne ikke finde kodeblok 7; backup: '+str(backup))
s=s.replace(old,new,1)
old='        if stats["force_refreshed"]:\n            print(f"   🔄 Tvangsopfrisket (fra {GENOPFRISK_TITLER_FILE.name}): {stats[\'force_refreshed\']:,}")'
new='        if stats["force_refreshed"]:\n            print(f"   🔄 Tvangsopfrisket (tekstliste/review): {stats[\'force_refreshed\']:,}")'

if old not in s: sys.exit('Kunne ikke finde kodeblok 8; backup: '+str(backup))
s=s.replace(old,new,1)
ast.parse(s)
target.write_text(s,encoding='utf-8')
print('OK:',target)
print('Backup:',backup)
