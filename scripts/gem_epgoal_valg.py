#!/usr/bin/env python3
"""
gem_epgoal_valg.py — EPGOAL-modstykke til det gamle gem_mine_valg.bat.

Kør denne EFTER du har udfyldt og GEMT BEGGE:
    data/sport_artwork_review.xlsx
    data/danish_artwork_review.xlsx

Den skriver dine valg tilbage til systemet og bager dem ind i epgoal.xml:
    1) import_sport_review.py   - skriver dine sport-billedvalg til
                                   sport_program_overrides.json (din
                                   EKSISTERENDE, uændrede fil)
    2) Sport-berigelse           - genkører sport-matchning på
                                   epgshare_merged.xml med de nye valg
                                   (samme logik som opdater_epgoal.py trin 4)
    3) Danske TMDb-backdrops     - injicerer NYE godkendte (X) danske
                                   backdrops fra danish_artwork_review.xlsx
                                   (samme logik som opdater_epgoal.py trin 5)
    4) Normalisering             - bygger epgoal.xml på ny fra den opdaterede
                                   epgshare_merged.xml (samme logik som
                                   opdater_epgoal.py trin 6)
    5) Git commit + push

VIGTIGT: Dette script downloader IKKE frisk EPG og laver IKKE en ny merge -
det genbruger den eksisterende output_epgshare/epgshare_merged.xml fra
sidste fulde kørsel af opdater_epgoal.py. Kør opdater_epgoal.py igen i
stedet, hvis du også vil have frisk EPG-data.

Kør:
    python scripts/gem_epgoal_valg.py
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.etree import ElementTree as ET

# Genbruger ALLE funktioner og konstanter fra opdater_epgoal.py, så der ikke
# findes to kopier af sport-/TMDb-/normaliserings-logikken der kan løbe fra
# hinanden over tid.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import opdater_epgoal as core  # noqa: E402


def run_import_sport_review() -> bool:
    print()
    print("=" * 70)
    print("TRIN 1/5: Gem sport-billedvalg (import_sport_review.py)")
    print("=" * 70)
    ok = core.run_script("import_sport_review.py")
    if not ok:
        print("FEJL: import_sport_review.py fejlede - dine sport-valg er IKKE gemt endnu.", file=sys.stderr)
    return ok


def run_sport_pass() -> None:
    print()
    print("=" * 70)
    print("TRIN 2/5: Sport-berigelse med nye valg (epgshare_merged.xml)")
    print("=" * 70)

    if not core.EPGSHARE_MERGED_FILE.exists():
        sys.exit(
            f"FEJL: {core.EPGSHARE_MERGED_FILE} findes ikke.\n"
            "Kør 'python scripts/opdater_epgoal.py' først for at bygge grundlaget."
        )

    matcher = core.SportMatcher(core.SPORT_IMAGE_BASE_URL)
    tmdb_overrides = core.load_json(core.TMDB_OVERRIDES_FILE, {})
    cache = core.load_json(core.CACHE_FILE, {})
    cache_before = len(cache)
    sport_tmdb_fallback_enabled = bool(core.TMDB_API_KEY)

    tree = ET.parse(core.EPGSHARE_MERGED_FILE)
    root = tree.getroot()

    stats, (fallback_log, partial_log, all_sport_log) = core.run_sport_enrichment_on_tree(
        root, matcher, tmdb_overrides, cache, sport_tmdb_fallback_enabled
    )

    tree.write(core.EPGSHARE_MERGED_FILE, encoding="utf-8", xml_declaration=True)
    core.save_json(core.CACHE_FILE, cache)
    core.save_json(core.FALLBACK_LOG_FILE, fallback_log)
    core.save_json(core.PARTIAL_SPORT_LOG_FILE, partial_log)
    core.save_json(core.ALL_SPORT_TITLES_LOG_FILE, all_sport_log)

    print(f"Programmer i alt        : {stats['programmes']:,}")
    print(f"Sport - specifikt match : {stats['sport_matched']:,}")
    print(
        f"Sport - TMDb-match      : {stats['sport_tmdb_matched']:,} "
        f"(cache: {stats['sport_tmdb_cache_hit']:,} / friske: {stats['sport_tmdb_fresh_call']:,})"
    )
    print(f"Sport - kanal-fallback  : {stats['sport_defaulted']:,}")
    print(f"Sport - sprunget over   : {stats['sport_skipped']:,}")
    print(f"Cache voksede fra {cache_before:,} til {len(cache):,} unikke titler")


def run_danish_backdrops_pass() -> None:
    print()
    print("=" * 70)
    print("TRIN 3/5: Danske backdrops - injicerer nye godkendelser (X)")
    print("=" * 70)

    tree = ET.parse(core.EPGSHARE_MERGED_FILE)
    root = tree.getroot()
    core.run_danish_backdrops_on_tree(root)
    tree.write(core.EPGSHARE_MERGED_FILE, encoding="utf-8", xml_declaration=True)


def run_normalize_pass() -> None:
    print()
    print("=" * 70)
    print("TRIN 4/5: Normaliserer til Output/UHF tvg-id (epgoal.xml)")
    print("=" * 70)
    core.step6_normalize_uhf_ids()


def run_git_pass() -> None:
    print()
    print("=" * 70)
    print("TRIN 5/5: Git commit + push")
    print("=" * 70)
    if not core.GIT_ENABLED:
        print("Git er deaktiveret - springer over.")
        return
    message = f"Gem mine valg (sport + danske backdrops) {core.time.strftime('%Y-%m-%d %H:%M:%S')}"
    core.git_push(core.ROOT, message)


def main() -> None:
    if not core.SPORT_ARTWORK_REVIEW_FILE.exists():
        sys.exit(f"FEJL: {core.SPORT_ARTWORK_REVIEW_FILE} findes ikke - kør opdater_epgoal.py først.")
    if not core.DANISH_ARTWORK_REVIEW_FILE.exists():
        sys.exit(f"FEJL: {core.DANISH_ARTWORK_REVIEW_FILE} findes ikke - kør opdater_epgoal.py først.")

    print("=== GEM EPGOAL VALG ===")
    print(f"Tidspunkt: {core.time.strftime('%Y-%m-%d %H:%M:%S')}")

    sport_import_ok = run_import_sport_review()
    if not sport_import_ok:
        print("\nFortsætter alligevel til sport-berigelse (uden nye sport-overrides).", file=sys.stderr)

    run_sport_pass()
    run_danish_backdrops_pass()
    run_normalize_pass()
    run_git_pass()

    print()
    print("=" * 70)
    print("FÆRDIG! Dine valg er gemt, og epgoal.xml er opdateret og pushet.")
    print("=" * 70)
    print("UHF-URL (uændret):")
    print("  https://raw.githubusercontent.com/flanaganz/epgoal/main/output_epgshare/epgoal.xml")


if __name__ == "__main__":
    main()
