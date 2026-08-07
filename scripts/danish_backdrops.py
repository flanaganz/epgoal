#!/usr/bin/env python3
"""
danish_backdrops.py — separat sideprojekt til danske TMDb-backdrops (Mac mini / always-on)

FORMÅL
    UHF henter selv billeder fra TMDb for programmer uden <icon>/<backdrop> i
    XML'en, men understøtter IKKE at vælge foretrukket sprog - den vælger
    altid det højest ratede billede og falder tilbage til US/UK-billeder,
    selvom du har uploadet danske backdrops til TMDb. Dette script retter det
    ved selv at slå titler op, tjekke om der findes et ÆGTE dansk backdrop
    (include_image_language=da, UDEN fallback til andre sprog), og i så fald
    indlejre det direkte i XML'en - så UHF bare bruger det, der allerede står
    der, i stedet for at lave sit eget (sprog-uafhængige) valg.

ADSKILLELSE FRA SPORTS-SCRIPTET (enrich_epg.py) - VIGTIGT
    - Dette script rører ALDRIG data/cache.json (sports-scriptets cache).
      Det har sin egen, helt separate cache: data/danish_artwork_cache.json.
    - Det rører ALDRIG sport_categories.json, sport_channels.json el.lign.
    - Det læser output/denmarkX.xml (dvs. FILERNE EFTER sports-scriptet er
      kørt), og springer AUTOMATISK alle programmer over, der allerede har
      et <icon> eller <backdrop> tag - det er præcis de programmer,
      sports-scriptet allerede har kurateret. Dette er den konkrete garanti
      for at sport-billeder aldrig bliver overskrevet af dette script.
    - Kør ALTID dette script EFTER enrich_epg.py, aldrig før eller i stedet
      for det.

BRUG
    python3 scripts/danish_backdrops.py
    python3 scripts/danish_backdrops.py --limit 50      (test på kun 50 unikke titler)
    python3 scripts/danish_backdrops.py --files denmark1 (kør kun på én fil)

CACHE
    data/danish_artwork_cache.json gemmer BÅDE positive resultater (fundet et
    dansk backdrop) OG negative resultater (bekræftet INGEN dansk backdrop
    findes), så scriptet ikke spilder TMDb-kald på at spørge om det samme
    igen og igen. Levetid styres af cache_max_age_days i config.json
    (samme indstilling som sports-scriptet bruger, men i en helt separat fil).
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
from pathlib import Path
from xml.etree import ElementTree as ET

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
CONFIG_FILE = ROOT / "config.json"

DANISH_ARTWORK_CACHE_FILE = DATA_DIR / "danish_artwork_cache.json"

ENV_FILE = ROOT / ".env"
if ENV_FILE.exists():
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

TMDB_API_KEY = os.environ.get("TMDB_API_KEY")
TMDB_BASE = "https://api.themoviedb.org/3"
IMAGE_BASE = "https://image.tmdb.org/t/p"

MATCH_SIMILARITY_MIN = 0.55
REQUEST_SLEEP_SECONDS = 0.05
SAVE_CACHE_EVERY_N_LOOKUPS = 25  # gem cachen løbende, så et afbrudt kør ikke mister alt arbejde

INVISIBLE_CHARS_PATTERN = re.compile(r"[\u200B\u200C\u200D\u2060\uFEFF\u00AD]")

SESSION = requests.Session()


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            print(f"⚠️  Kunne ikke læse {path}, bruger default", file=sys.stderr)
    return default


def save_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def normalize_title(title: str) -> str:
    """Samme normalisering som sports-scriptet, for konsistens (men bruges her
    kun som cache-nøgle i en helt separat cache-fil)."""
    t = title.strip()
    t = unicodedata.normalize("NFKC", t)
    t = INVISIBLE_CHARS_PATTERN.sub(" ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip().lower()


def tmdb_search(title: str) -> tuple[str, int] | None:
    """Finder det bedste tv/film-match til titlen. Samme similarity-logik som
    sports-scriptet, for at undgå forkerte matches på tvetydige titler."""
    resp = SESSION.get(
        f"{TMDB_BASE}/search/multi",
        params={"api_key": TMDB_API_KEY, "query": title, "language": "da-DK", "include_adult": "false"},
        timeout=15,
    )
    resp.raise_for_status()
    results = [r for r in resp.json().get("results", []) if r.get("media_type") in ("tv", "movie")]
    if not results:
        return None

    def score(r):
        name = r.get("name") or r.get("title") or ""
        similarity = difflib.SequenceMatcher(None, name.lower(), title.lower()).ratio()
        popularity_bonus = min(r.get("popularity", 0), 50) / 50
        return similarity * 2 + popularity_bonus

    results.sort(key=score, reverse=True)
    best = results[0]
    best_name = best.get("name") or best.get("title") or ""
    similarity = difflib.SequenceMatcher(None, best_name.lower(), title.lower()).ratio()
    if similarity < MATCH_SIMILARITY_MIN:
        return None
    return best["media_type"], best["id"]


def tmdb_danish_images(media_type: str, tmdb_id: int) -> tuple[str | None, str | None]:
    """Henter KUN danske billeder (include_image_language=da, INGEN fallback
    til andre sprog). Returnerer (backdrop_path, poster_path) - begge kan
    være None, hvis intet dansk billede af den type findes."""
    resp = SESSION.get(
        f"{TMDB_BASE}/{media_type}/{tmdb_id}/images",
        params={"api_key": TMDB_API_KEY, "include_image_language": "da"},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    def pick_best(items):
        if not items:
            return None
        items = sorted(items, key=lambda b: -b.get("vote_average", 0))
        return items[0]["file_path"]

    return pick_best(data.get("backdrops", [])), pick_best(data.get("posters", []))


def resolve_danish_artwork(raw_title: str, cache: dict, cache_max_age_days: int,
                           backdrop_size: str, poster_size: str) -> tuple[dict | None, bool]:
    """Returnerer (dict med backdrop/poster-URL'er ELLER None hvis intet
    dansk fundet, bool om resultatet kom fra cache)."""
    key = normalize_title(raw_title)

    cached = cache.get(key)
    if cached is not None and (time.time() - cached.get("ts", 0)) / 86400 < cache_max_age_days:
        if cached.get("backdrop") or cached.get("poster"):
            return {"backdrop": cached.get("backdrop"), "poster": cached.get("poster")}, True
        return None, True  # bekræftet: intet dansk billede findes

    backdrop_url = poster_url = None
    try:
        match = tmdb_search(raw_title)
        if match:
            media_type, tmdb_id = match
            b_path, p_path = tmdb_danish_images(media_type, tmdb_id)
            if b_path:
                backdrop_url = f"{IMAGE_BASE}/{backdrop_size}{b_path}"
            if p_path:
                poster_url = f"{IMAGE_BASE}/{poster_size}{p_path}"
    except requests.RequestException as exc:
        print(f"   TMDb-fejl for '{raw_title}': {exc}", file=sys.stderr)
        return None, False  # fejl -> spring over denne gang, cache IKKE et negativt resultat

    cache[key] = {"backdrop": backdrop_url, "poster": poster_url, "ts": time.time()}

    if backdrop_url or poster_url:
        return {"backdrop": backdrop_url, "poster": poster_url}, False
    return None, False


def process_xml_file(xml_path: Path, cache: dict, cache_max_age_days: int,
                      backdrop_size: str, poster_size: str, limit: int | None,
                      titles_processed_this_run: set) -> dict:
    """Læser en output/denmarkX.xml, tilføjer danske backdrops/postere til
    programmer der IKKE allerede har <icon>/<backdrop> (dvs. springer alt
    sport-kurateret indhold over), og gemmer filen igen (overskriver in-place)."""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    stats = {
        "programmes": 0, "already_had_artwork": 0, "checked": 0,
        "danish_found": 0, "danish_not_found": 0, "cache_hits": 0, "fresh_calls": 0,
    }

    resolved_this_file: dict[str, dict | None] = {}

    for programme in root.findall("programme"):
        stats["programmes"] += 1

        if programme.find("icon") is not None or programme.find("backdrop") is not None:
            stats["already_had_artwork"] += 1
            continue

        title_el = programme.find("title")
        if title_el is None or not title_el.text:
            continue
        title = title_el.text
        norm = normalize_title(title)

        if limit is not None and norm not in titles_processed_this_run and len(titles_processed_this_run) >= limit:
            continue  # testgrænse nået - spring resterende NYE titler over (allerede sete titler behandles stadig)

        if norm not in resolved_this_file:
            was_cached_before = norm in cache and (time.time() - cache[norm].get("ts", 0)) / 86400 < cache_max_age_days
            art, from_cache = resolve_danish_artwork(title, cache, cache_max_age_days, backdrop_size, poster_size)
            resolved_this_file[norm] = art
            titles_processed_this_run.add(norm)
            stats["checked"] += 1
            if from_cache:
                stats["cache_hits"] += 1
            else:
                stats["fresh_calls"] += 1
                time.sleep(REQUEST_SLEEP_SECONDS)
            if art:
                stats["danish_found"] += 1
            else:
                stats["danish_not_found"] += 1

        art = resolved_this_file[norm]
        if art:
            if art.get("poster"):
                el = ET.SubElement(programme, "icon")
                el.set("src", art["poster"])
            if art.get("backdrop"):
                el = ET.SubElement(programme, "backdrop")
                el.set("src", art["backdrop"])

    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    return stats


def git_push(repo_dir: Path, commit_message: str) -> None:
    print("\n⬆️  Committer og pusher til GitHub ...")
    try:
        subprocess.run(["git", "add", "-A"], cwd=repo_dir, check=False)
        result = subprocess.run(
            ["git", "commit", "-m", commit_message], cwd=repo_dir, check=False, capture_output=True, text=True
        )
        if "nothing to commit" in (result.stdout + result.stderr).lower():
            print("   Ingen ændringer at committe.")
            return
        subprocess.run(["git", "push", "origin", "main"], cwd=repo_dir, check=False)
        print("✅ Git push forsøgt (tjek GitHub for resultat).")
    except FileNotFoundError:
        print("⚠️  git blev ikke fundet i PATH — spring commit/push over.", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(description="Tilføj danske TMDb-backdrops til ikke-sport-programmer.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Maks antal NYE (ikke-cachede) unikke titler at slå op - til test af en mindre batch.")
    parser.add_argument("--files", nargs="*", default=None,
                        help="Kør kun på specifikke kildenavne (fx denmark1 denmark3). Default: alle kilder fra config.json.")
    args = parser.parse_args()

    if not TMDB_API_KEY:
        sys.exit("❌ TMDB_API_KEY er ikke sat i .env - kan ikke slå danske backdrops op.")

    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    if not sources:
        sys.exit("❌ Ingen kilder defineret i config.json.")

    source_names = [s["name"] for s in sources]
    if args.files:
        source_names = [n for n in source_names if n in args.files]
        if not source_names:
            sys.exit(f"❌ Ingen af de angivne --files matcher kilderne i config.json ({[s['name'] for s in sources]}).")

    backdrop_size = config.get("image", {}).get("backdrop_size", "w1280")
    poster_size = config.get("image", {}).get("poster_size", "w500")
    cache_max_age_days = config.get("cache_max_age_days", 30)
    git_cfg = config.get("git", {})

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    cache = load_json(DANISH_ARTWORK_CACHE_FILE, {})
    cache_size_before = len(cache)

    print("=== Danske TMDb-backdrops (separat sideprojekt) ===")
    print(f"Cache-fil: {DANISH_ARTWORK_CACHE_FILE.name} (separat fra sports-scriptets cache.json)")
    print(f"Cache indeholder {cache_size_before:,} tidligere opslag (levetid: {cache_max_age_days} dage)")
    if args.limit:
        print(f"⚠️  TEST-TILSTAND: maks {args.limit} NYE unikke titler slås op i denne kørsel")
    print(f"Behandler filer: {', '.join(source_names)}")

    titles_processed_this_run: set = set()
    grand_total = {
        "programmes": 0, "already_had_artwork": 0, "checked": 0,
        "danish_found": 0, "danish_not_found": 0, "cache_hits": 0, "fresh_calls": 0,
    }

    for name in source_names:
        xml_path = OUTPUT_DIR / f"{name}.xml"
        if not xml_path.exists():
            print(f"\n⚠️  {xml_path} findes ikke - har du kørt enrich_epg.py først? Springer over.")
            continue

        print(f"\n📄 Behandler {xml_path.name} ...")
        stats = process_xml_file(
            xml_path, cache, cache_max_age_days, backdrop_size, poster_size,
            args.limit, titles_processed_this_run,
        )
        save_json(DANISH_ARTWORK_CACHE_FILE, cache)  # løbende gemning efter hver fil

        print(f"   Programmer i alt              : {stats['programmes']:,}")
        print(f"   Sprunget over (havde allerede billede - sport): {stats['already_had_artwork']:,}")
        print(f"   Titler tjekket (unikke)        : {stats['checked']:,}")
        print(f"   Dansk backdrop/poster fundet   : {stats['danish_found']:,}")
        print(f"   Bekræftet intet dansk billede  : {stats['danish_not_found']:,}")
        print(f"   (cache: {stats['cache_hits']:,} / friske TMDb-kald: {stats['fresh_calls']:,})")

        for k in grand_total:
            grand_total[k] += stats[k]

    save_json(DANISH_ARTWORK_CACHE_FILE, cache)
    cache_size_after = len(cache)

    print("\n📊 SAMLET RAPPORT")
    print("--------------------------------")
    print(f"Programmer i alt              : {grand_total['programmes']:,}")
    print(f"Sprunget over (sport)          : {grand_total['already_had_artwork']:,}")
    print(f"Titler tjekket (unikke)        : {grand_total['checked']:,}")
    print(f"Dansk backdrop/poster tilføjet : {grand_total['danish_found']:,}")
    print(f"Bekræftet intet dansk billede  : {grand_total['danish_not_found']:,}")
    print(f"Cache voksede fra {cache_size_before:,} til {cache_size_after:,} unikke titler")
    print("--------------------------------")

    if git_cfg.get("enabled", True):
        prefix = "Auto-sync EPG (danske backdrops)"
        git_push(ROOT, f"{prefix} {time.strftime('%Y-%m-%d %H:%M:%S')}")

    print("\n=== Færdig. Danske backdrops tilføjet til output/-filerne. ===")


if __name__ == "__main__":
    main()
