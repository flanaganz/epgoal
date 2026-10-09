#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Patch opdater_epgoal.py med en fast OpenEPG-override for Discovery Science.

- Tager timestampet backup.
- Ændrer kun kildevalget i trin 3.
- Validerer det ændrede script med ast.parse.
- Kan køres flere gange uden at indsætte rettelsen igen.
"""
from __future__ import annotations

import ast
import shutil
from datetime import datetime
from pathlib import Path

TARGET = Path(r"C:\EPGoal\scripts\opdater_epgoal.py")
BACKUP_DIR = Path(r"C:\EPGoal\reports")

CONSTANT_MARKER = 'MIN_WINDOW_HOURS = 24\n'
CONSTANT_BLOCK = '''MIN_WINDOW_HOURS = 24

# Kanaler som altid skal bruge en bestemt fallback-kilde, selv når EPGShare
# teknisk set har et tilstrækkeligt tidsvindue. Nøglen er værdien i kolonnen
# "Kanal" i channel_priority_v2.xlsx.
SOURCE_OVERRIDES = {
    "Discovery Science.dk": {
        "source": "openepg",
        "channel_id": "Discovery Science.dk",
    },
}
'''

LOOP_MARKER = '''        epgshare_span = max(
            (epgshare_stats.get(c, {}).get("span_hours", 0) for c in matched_epgshare_ids),
            default=0,
        )
        if epgshare_count > 0 and epgshare_span >= MIN_WINDOW_HOURS:
'''

LOOP_REPLACEMENT = '''        epgshare_span = max(
            (epgshare_stats.get(c, {}).get("span_hours", 0) for c in matched_epgshare_ids),
            default=0,
        )

        # Fast kildeoverride. Bruges til kendte tilfælde hvor EPGShare har et
        # langt vindue, men næsten intet reelt programindhold.
        override = SOURCE_OVERRIDES.get(kanonisk_navn)
        if override:
            override_source = override.get("source")
            override_channel_id = override.get("channel_id")

            if override_source == "openepg":
                override_programmes_map = openepg_programmes_by_channel
                override_stats_map = openepg_stats
            elif override_source == "bss":
                override_programmes_map = bss_programmes_by_channel
                override_stats_map = bss_stats
            else:
                raise ValueError(
                    f"Ukendt source override for {kanonisk_navn}: {override_source}"
                )

            override_stat = override_stats_map.get(
                override_channel_id,
                {"count": 0, "span_hours": 0},
            )

            if override_stat["count"] <= 0:
                print(
                    f"ADVARSEL: Kildeoverride for {kanonisk_navn} fandt ingen "
                    f"programmer på {override_source}:{override_channel_id}. "
                    "Fortsætter med normal kildeprioritet.",
                    file=sys.stderr,
                )
            else:
                if matched_epgshare_ids:
                    primary_id = max(
                        matched_epgshare_ids,
                        key=lambda c: epgshare_stats.get(c, {}).get("count", 0),
                    )
                else:
                    primary_id = f"{normalize_id(kanonisk_navn)}.dk"
                    new_channel_el = ET.Element("channel", {"id": primary_id})
                    dn = ET.SubElement(new_channel_el, "display-name")
                    dn.text = kanonisk_navn
                    root.append(new_channel_el)

                for cid in matched_epgshare_ids:
                    for p in epgshare_programmes_by_channel.get(cid, []):
                        root.remove(p)

                inserted = 0
                for p in override_programmes_map.get(override_channel_id, []):
                    clone = copy.deepcopy(p)
                    clone.set("channel", primary_id)
                    strip_artwork(clone)
                    root.append(clone)
                    inserted += 1

                report["supplemented"].append({
                    "kanal": kanonisk_navn,
                    "epgshare_programmer_foer": epgshare_count,
                    "epgshare_vindue_timer_foer": round(epgshare_span, 1),
                    "kilde_brugt": override_source,
                    "kilde_channel_id": override_channel_id,
                    "programmer_indsat": inserted,
                    "vindue_timer_efter": round(override_stat["span_hours"], 1),
                    "primaer_channel_id": primary_id,
                    "tvunget_kildeoverride": True,
                })
                print(
                    f"Kildeoverride: {kanonisk_navn} -> "
                    f"{override_source}:{override_channel_id} "
                    f"({inserted:,} programmer)"
                )
                continue

        if epgshare_count > 0 and epgshare_span >= MIN_WINDOW_HOURS:
'''


def main() -> None:
    if not TARGET.exists():
        raise SystemExit(f"FEJL: Filen findes ikke: {TARGET}")

    source = TARGET.read_text(encoding="utf-8")

    if "SOURCE_OVERRIDES = {" not in source:
        if CONSTANT_MARKER not in source:
            raise SystemExit("FEJL: Kunne ikke finde MIN_WINDOW_HOURS = 24")
        source = source.replace(CONSTANT_MARKER, CONSTANT_BLOCK, 1)

    if '"tvunget_kildeoverride": True' not in source:
        if LOOP_MARKER not in source:
            raise SystemExit(
                "FEJL: Kunne ikke finde merge-blokken. Scriptet kan være ændret. "
                "Ingen fil blev skrevet."
            )
        source = source.replace(LOOP_MARKER, LOOP_REPLACEMENT, 1)

    # Syntakskontrol før originalen røres.
    ast.parse(source, filename=str(TARGET))

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = BACKUP_DIR / f"opdater_epgoal_before_discovery_override_{stamp}.py"
    shutil.copy2(TARGET, backup)
    TARGET.write_text(source, encoding="utf-8")

    # Kontrol efter skrivning.
    written = TARGET.read_text(encoding="utf-8")
    ast.parse(written, filename=str(TARGET))
    required = [
        '"Discovery Science.dk": {',
        '"source": "openepg"',
        '"channel_id": "Discovery Science.dk"',
        '"tvunget_kildeoverride": True',
    ]
    missing = [value for value in required if value not in written]
    if missing:
        shutil.copy2(backup, TARGET)
        raise SystemExit("FEJL: Efterkontrol fejlede. Originalen er gendannet.")

    print("=== DISCOVERY SCIENCE OVERRIDE TILFØJET ===")
    print(f"Opdateret : {TARGET}")
    print(f"Backup    : {backup}")
    print("Override  : Discovery Science.dk -> OpenEPG:Discovery Science.dk")
    print()
    print("Kør nu:")
    print(r"  cd C:\EPGoal\scripts")
    print("  python opdater_epgoal.py")


if __name__ == "__main__":
    main()
