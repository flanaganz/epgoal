#!/usr/bin/env python3
"""
enrich_epg.py — sport-first lokal udgave (Mac mini / always-on)

Matching-strategi pr. programme:
    -1) sport_channels.json["exclude"] -> kanalen behandles som IKKE-sport
    0) sport_skip_titles.json          -> fjern evt. icon/backdrop, rør intet andet
    0.5) sport_prefer_tmdb_titles.json -> SPRING lokal kategori/nøgleords-matching
         over for disse titler, gå direkte til TMDb-opslag (trin 4).
    1) sport_program_overrides.json  -> eksakt FULD titel-match
    2) sport_categories.json[keywords]  -> nøgleord (priority-kategorier først,
       derefter almindelige kategorier længst-match-først)
    3) sport_categories.json[prefix]    -> "Kategori: Begivenhed"-syntaksen
       (SIDSTE UDVEJ blandt de lokale trin)
    4) TMDb-opslag -> kun for "always_sport"-kanaler, kun hvis aktiveret
    5) kanalens default_backdrop/-poster (sidste udvej, kun "always_sport")

    "partial_sport"-kanaler (fx DR1, DR2, TV 2 hovedkanal) rammes KUN via
    trin 0.5-4 ovenfor - de har INGEN kanal-fallback-billede (trin 5), da
    langt størstedelen af deres programmer ikke er sport. Programmer der
    korrekt matches på disse kanaler logges separat i
    data/partial_sport_matches_log.json (se nedenfor), så det er muligt at
    følge med i, hvad der reelt bliver fanget af landskampe/OL/VM m.v.

Titel-normalisering bruger unicodedata.normalize("NFKC", ...) samt eksplicit
erstatning af usynlige tegn (nulbredde-mellemrum, blødt bindestreg, BOM m.fl.)
med et almindeligt mellemrum.

Brug:
    python3 scripts/enrich_epg.py
"""
from __future__ import annotations

import difflib
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import quote
from xml.etree import ElementTree as ET

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
CONFIG_FILE = ROOT / "config.json"

CACHE_FILE = DATA_DIR / "cache.json"
TMDB_OVERRIDES_FILE = DATA_DIR / "overrides.json"

SPORT_CHANNELS_FILE = DATA_DIR / "sport_channels.json"
SPORT_CATEGORIES_FILE = DATA_DIR / "sport_categories.json"
SPORT_PROGRAM_OVERRIDES_FILE = DATA_DIR / "sport_program_overrides.json"
SPORT_SKIP_TITLES_FILE = DATA_DIR / "sport_skip_titles.json"
SPORT_PREFER_TMDB_TITLES_FILE = DATA_DIR / "sport_prefer_tmdb_titles.json"

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
    t = title.strip()
    t = unicodedata.normalize("NFKC", t)
    t = INVISIBLE_CHARS_PATTERN.sub(" ", t)
    t = re.sub(r"\s+", " ", t)
    t = t.strip()
    return t.lower()


class SportMatcher:
    def __init__(self, image_base_url: str):
        self.image_base_url = image_base_url.rstrip("/") + "/"

        channels_raw = load_json(SPORT_CHANNELS_FILE, {})
        if isinstance(channels_raw, list):
            self.exclude_patterns: list[str] = []
            self.channels = channels_raw
        else:
            self.exclude_patterns = [e["match"].lower() for e in channels_raw.get("exclude", [])]
            self.channels = channels_raw.get("channels", [])

        self.categories = load_json(SPORT_CATEGORIES_FILE, [])
        self.program_overrides = load_json(SPORT_PROGRAM_OVERRIDES_FILE, {})
        self.skip_titles = set(load_json(SPORT_SKIP_TITLES_FILE, []))

        prefer_tmdb_raw = load_json(SPORT_PREFER_TMDB_TITLES_FILE, {"titles": []})
        self.prefer_tmdb_titles = set(t.strip().lower() for t in prefer_tmdb_raw.get("titles", []))

        self.prefix_lookup: dict[str, dict] = {}
        for cat in self.categories:
            for prefix in cat.get("prefix", []) or []:
                self.prefix_lookup[prefix.strip().lower()] = cat

        priority_keywords: list[tuple[str, dict]] = []
        normal_keywords: list[tuple[str, dict]] = []
        for cat in self.categories:
            target = priority_keywords if cat.get("priority") else normal_keywords
            for keyword in cat.get("keywords", []) or []:
                target.append((keyword.strip().lower(), cat))
        normal_keywords.sort(key=lambda pair: len(pair[0]), reverse=True)
        self.keyword_lookup: list[tuple[str, dict]] = priority_keywords + normal_keywords

    def match_channel(self, channel_id: str) -> dict | None:
        low = (channel_id or "").lower()

        for pattern in self.exclude_patterns:
            if pattern in low:
                return None

        for entry in self.channels:
            if entry["match"].lower() in low:
                return entry
        return None

    def _image_urls(self, backdrop_filename: str | None, poster_filename: str | None) -> dict:
        def build(filename: str | None) -> str | None:
            if not filename:
                return None
            return self.image_base_url + quote(filename)

        return {"backdrop": build(backdrop_filename), "poster": build(poster_filename)}

    def _real_match(self, backdrop_filename: str | None, poster_filename: str | None) -> dict | None:
        if not backdrop_filename and not poster_filename:
            return None
        return self._image_urls(backdrop_filename, poster_filename)

    def prefers_tmdb(self, raw_title: str) -> bool:
        return normalize_title(raw_title) in self.prefer_tmdb_titles

    def resolve_local(self, raw_title: str) -> dict | None:
        norm = normalize_title(raw_title)

        if norm in self.skip_titles:
            return {"skip": True}

        if norm in self.prefer_tmdb_titles:
            return None

        override = self.program_overrides.get(norm)
        if override:
            match = self._real_match(override.get("backdrop"), override.get("poster"))
            if match:
                return match

        for keyword, cat in self.keyword_lookup:
            if keyword in norm:
                match = self._real_match(cat.get("backdrop"), cat.get("poster"))
                if match:
                    return match

        prefix = norm.split(":", 1)[0].strip()
        cat = self.prefix_lookup.get(prefix)
        if cat:
            match = self._real_match(cat.get("backdrop"), cat.get("poster"))
            if match:
                return match

        return None


def tmdb_search(title: str) -> tuple[str, int] | None:
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


def tmdb_images(media_type: str, tmdb_id: int) -> tuple[str | None, str | None]:
    resp = SESSION.get(
        f"{TMDB_BASE}/{media_type}/{tmdb_id}/images",
        params={"api_key": TMDB_API_KEY, "include_image_language": "da,en,null"},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    def pick(items):
        if not items:
            return None
        lang_rank = {"da": 0, "en": 1, None: 2}
        items = sorted(items, key=lambda b: (lang_rank.get(b.get("iso_639_1"), 3), -b.get("vote_average", 0)))
        return items[0]["file_path"]

    return pick(data.get("backdrops", [])), pick(data.get("posters", []))


def resolve_tmdb_artwork(raw_title: str, overrides: dict, cache: dict, cache_max_age_days: int,
                          backdrop_size: str, poster_size: str) -> tuple[dict, bool]:
    key = normalize_title(raw_title)

    override = overrides.get(key)
    if override:
        if override.get("backdrop_url") or override.get("poster_url"):
            return {"backdrop": override.get("backdrop_url"), "poster": override.get("poster_url")}, False
        if override.get("tmdb_id") and override.get("media_type"):
            try:
                b_path, p_path = tmdb_images(override["media_type"], override["tmdb_id"])
                return {
                    "backdrop": f"{IMAGE_BASE}/{backdrop_size}{b_path}" if b_path else None,
                    "poster": f"{IMAGE_BASE}/{poster_size}{p_path}" if p_path else None,
                }, False
            except requests.RequestException as exc:
                print(f"   TMDb-fejl (override) for '{raw_title}': {exc}", file=sys.stderr)

    cached = cache.get(key)
    if cached is not None and (time.time() - cached.get("ts", 0)) / 86400 < cache_max_age_days:
        return {"backdrop": cached.get("backdrop"), "poster": cached.get("poster")}, True

    backdrop_url = poster_url = None
    try:
        match = tmdb_search(raw_title)
        if match:
            media_type, tmdb_id = match
            b_path, p_path = tmdb_images(media_type, tmdb_id)
            if b_path:
                backdrop_url = f"{IMAGE_BASE}/{backdrop_size}{b_path}"
            if p_path:
                poster_url = f"{IMAGE_BASE}/{poster_size}{p_path}"
    except requests.RequestException as exc:
        print(f"   TMDb-fejl for '{raw_title}': {exc}", file=sys.stderr)

    cache[key] = {"backdrop": backdrop_url, "poster": poster_url, "ts": time.time()}
    return {"backdrop": backdrop_url, "poster": poster_url}, False


def set_artwork(programme: ET.Element, backdrop_url: str | None, poster_url: str | None) -> tuple[bool, bool]:
    for old in programme.findall("icon"):
        programme.remove(old)
    for old in programme.findall("backdrop"):
        programme.remove(old)

    added_poster = added_backdrop = False
    if poster_url:
        el = ET.SubElement(programme, "icon")
        el.set("src", poster_url)
        added_poster = True
    if backdrop_url:
        el = ET.SubElement(programme, "backdrop")
        el.set("src", backdrop_url)
        added_backdrop = True
    return added_poster, added_backdrop


def clear_artwork(programme: ET.Element) -> None:
    for old in programme.findall("icon"):
        programme.remove(old)
    for old in programme.findall("backdrop"):
        programme.remove(old)


def process_xml(xml_bytes: bytes, matcher: SportMatcher, source_cfg: dict,
                 tmdb_overrides: dict, cache: dict, cache_max_age_days: int,
                 backdrop_size: str, poster_size: str, sport_tmdb_fallback_enabled: bool,
                 fallback_titles_log: dict, partial_sport_matches_log: dict) -> tuple[bytes, dict]:
    root = ET.fromstring(xml_bytes)

    channel_role: dict[str, dict | None] = {}
    for ch in root.findall("channel"):
        cid = ch.get("id", "")
        channel_role[cid] = matcher.match_channel(cid)

    stats = {
        "programmes": 0, "sport_matched": 0,
        "sport_tmdb_matched": 0, "sport_tmdb_cache_hit": 0, "sport_tmdb_fresh_call": 0,
        "sport_defaulted": 0, "sport_skipped": 0, "sport_no_image_yet": 0,
        "tmdb_enriched": 0, "partial_sport_matched": 0,
    }
    tmdb_cache_this_run: dict[str, tuple[dict, bool]] = {}
    sport_cache_this_run: dict[str, dict | None] = {}

    do_tmdb_nonsport = bool(source_cfg.get("enrich_non_sport_with_tmdb", False)) and bool(TMDB_API_KEY)
    do_tmdb_sport = bool(sport_tmdb_fallback_enabled) and bool(TMDB_API_KEY)

    for programme in root.findall("programme"):
        stats["programmes"] += 1
        title_el = programme.find("title")
        if title_el is None or not title_el.text:
            continue
        title = title_el.text

        chan_id = programme.get("channel", "")
        role_entry = channel_role.get(chan_id)

        if role_entry is not None:
            if title not in sport_cache_this_run:
                sport_cache_this_run[title] = matcher.resolve_local(title)
            result = sport_cache_this_run[title]

            if result and result.get("skip"):
                clear_artwork(programme)
                stats["sport_skipped"] += 1
                continue

            if result:
                set_artwork(programme, result.get("backdrop"), result.get("poster"))
                stats["sport_matched"] += 1
                if role_entry.get("role") == "partial_sport":
                    stats["partial_sport_matched"] += 1
                    partial_sport_matches_log.setdefault(chan_id, {})
                    key = f"{title} -> {result.get('backdrop') or result.get('poster')}"
                    partial_sport_matches_log[chan_id][key] = partial_sport_matches_log[chan_id].get(key, 0) + 1
                continue

            if do_tmdb_sport and role_entry.get("role") in ("always_sport", "partial_sport"):
                if title not in tmdb_cache_this_run:
                    art, from_cache = resolve_tmdb_artwork(
                        title, tmdb_overrides, cache, cache_max_age_days, backdrop_size, poster_size
                    )
                    tmdb_cache_this_run[title] = (art, from_cache)
                    if not from_cache:
                        time.sleep(REQUEST_SLEEP_SECONDS)
                art, from_cache = tmdb_cache_this_run[title]
                if art.get("backdrop") or art.get("poster"):
                    set_artwork(programme, art.get("backdrop"), art.get("poster"))
                    stats["sport_tmdb_matched"] += 1
                    if from_cache:
                        stats["sport_tmdb_cache_hit"] += 1
                    else:
                        stats["sport_tmdb_fresh_call"] += 1
                    if role_entry.get("role") == "partial_sport":
                        stats["partial_sport_matched"] += 1
                        partial_sport_matches_log.setdefault(chan_id, {})
                        key = f"{title} -> TMDb"
                        partial_sport_matches_log[chan_id][key] = partial_sport_matches_log[chan_id].get(key, 0) + 1
                    continue

            if role_entry.get("role") == "always_sport":
                fallback_backdrop = role_entry.get("default_backdrop")
                fallback_poster = role_entry.get("default_poster")
                fallback_titles_log.setdefault(chan_id, {})
                fallback_titles_log[chan_id][title] = fallback_titles_log[chan_id].get(title, 0) + 1
                if fallback_backdrop or fallback_poster:
                    urls = matcher._image_urls(fallback_backdrop, fallback_poster)
                    set_artwork(programme, urls["backdrop"], urls["poster"])
                    stats["sport_defaulted"] += 1
                else:
                    stats["sport_no_image_yet"] += 1
            # "partial_sport" uden match: rører intet, tælles ikke - langt størstedelen
            # af programmerne på disse kanaler er bevidst ikke-sport.
            continue

        if do_tmdb_nonsport:
            if title not in tmdb_cache_this_run:
                art, from_cache = resolve_tmdb_artwork(
                    title, tmdb_overrides, cache, cache_max_age_days, backdrop_size, poster_size
                )
                tmdb_cache_this_run[title] = (art, from_cache)
                if not from_cache:
                    time.sleep(REQUEST_SLEEP_SECONDS)
            art, _ = tmdb_cache_this_run[title]
            if art.get("backdrop") or art.get("poster"):
                set_artwork(programme, art.get("backdrop"), art.get("poster"))
                stats["tmdb_enriched"] += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), stats


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
    config = load_json(CONFIG_FILE, {})
    sources = config.get("sources", [])
    if not sources:
        sys.exit("❌ Ingen kilder defineret i config.json.")

    sport_cfg = config.get("sport", {})
    image_base_url = sport_cfg.get("image_base_url")
    if not image_base_url:
        sys.exit("❌ config.json mangler sport.image_base_url.")

    sport_tmdb_fallback_enabled = bool(sport_cfg.get("tmdb_fallback_enabled", False))

    any_tmdb = any(s.get("enrich_non_sport_with_tmdb") for s in sources) or sport_tmdb_fallback_enabled
    if any_tmdb and not TMDB_API_KEY:
        sys.exit("❌ TMDB_API_KEY er ikke sat, men TMDb-berigelse er aktiveret.")

    backdrop_size = config.get("image", {}).get("backdrop_size", "w1280")
    poster_size = config.get("image", {}).get("poster_size", "w500")
    cache_max_age_days = config.get("cache_max_age_days", 30)
    git_cfg = config.get("git", {})

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    matcher = SportMatcher(image_base_url)
    tmdb_overrides = load_json(TMDB_OVERRIDES_FILE, {})
    cache = load_json(CACHE_FILE, {})
    cache_size_before = len(cache)

    print("=== EPG sport-berigelse (lokal kørsel) ===")
    print(f"Sport-billeder hentes fra: {matcher.image_base_url}")
    print(f"Udelukkede kanal-mønstre: {matcher.exclude_patterns or '(ingen)'}")
    print(f"Titler der foretrækker TMDb: {sorted(matcher.prefer_tmdb_titles) or '(ingen)'}")
    print(f"TMDb-fallback for sport-programmer: {'AKTIVERET' if sport_tmdb_fallback_enabled and TMDB_API_KEY else 'slået fra'}")
    print(f"Cache indeholder {cache_size_before:,} tidligere TMDb-opslag (levetid: {cache_max_age_days} dage)")

    grand_total = {
        "programmes": 0, "sport_matched": 0,
        "sport_tmdb_matched": 0, "sport_tmdb_cache_hit": 0, "sport_tmdb_fresh_call": 0,
        "sport_defaulted": 0, "sport_skipped": 0, "sport_no_image_yet": 0, "tmdb_enriched": 0,
        "partial_sport_matched": 0,
    }
    fallback_titles_log: dict[str, dict[str, int]] = {}
    partial_sport_matches_log: dict[str, dict[str, int]] = {}

    for source in sources:
        name, url = source["name"], source["url"]
        print(f"\n📥 Henter {name} fra {url} ...")
        resp = SESSION.get(url, timeout=90)
        resp.raise_for_status()
        print(f"✅ Downloadet {len(resp.content)//1024} KB")

        print("🖼️  Behandler (sport-kanaler beriges, resten passerer uændret) ...")
        enriched, stats = process_xml(
            resp.content, matcher, source, tmdb_overrides, cache, cache_max_age_days,
            backdrop_size, poster_size, sport_tmdb_fallback_enabled,
            fallback_titles_log, partial_sport_matches_log,
        )

        out_path = OUTPUT_DIR / f"{name}.xml"
        out_path.write_bytes(enriched)
        print(f"💾 Gemt: {out_path}")

        total = stats["programmes"] or 1
        print(f"   Programmer i alt           : {stats['programmes']:,}")
        print(f"   Sport - specifikt match     : {stats['sport_matched']:,}")
        print(f"   Sport - TMDb-match i alt    : {stats['sport_tmdb_matched']:,} "
              f"(cache: {stats['sport_tmdb_cache_hit']:,} / friske kald: {stats['sport_tmdb_fresh_call']:,})")
        print(f"   Sport - kanal-fallback      : {stats['sport_defaulted']:,}")
        print(f"   Sport - sprunget over       : {stats['sport_skipped']:,}")
        print(f"   Sport - mangler billede     : {stats['sport_no_image_yet']:,}")
        print(f"   Heraf på partial_sport-kanaler (DR1/DR2/TV2): {stats['partial_sport_matched']:,}")
        if source.get("enrich_non_sport_with_tmdb"):
            print(f"   Ikke-sport TMDb-beriget     : {stats['tmdb_enriched']:,} ({stats['tmdb_enriched']/total:.1%})")

        for k in grand_total:
            grand_total[k] += stats[k]

    save_json(CACHE_FILE, cache)
    cache_size_after = len(cache)

    fallback_log_path = DATA_DIR / "fallback_titles_log.json"
    save_json(fallback_log_path, fallback_titles_log)

    partial_sport_log_path = DATA_DIR / "partial_sport_matches_log.json"
    save_json(partial_sport_log_path, partial_sport_matches_log)

    print("\n📊 SAMLET RAPPORT (alle 6 filer)")
    print("--------------------------------")
    print(f"Programmer i alt            : {grand_total['programmes']:,}")
    print(f"Sport - specifikt match     : {grand_total['sport_matched']:,}")
    print(f"Sport - TMDb-match i alt    : {grand_total['sport_tmdb_matched']:,} "
          f"(cache: {grand_total['sport_tmdb_cache_hit']:,} / friske kald: {grand_total['sport_tmdb_fresh_call']:,})")
    print(f"Sport - kanal-fallback      : {grand_total['sport_defaulted']:,}")
    print(f"Sport - sprunget over       : {grand_total['sport_skipped']:,}")
    print(f"Sport - mangler billede     : {grand_total['sport_no_image_yet']:,}")
    print(f"Ikke-sport TMDb-beriget     : {grand_total['tmdb_enriched']:,}")
    print(f"Heraf på partial_sport-kanaler (DR1/DR2/TV2): {grand_total['partial_sport_matched']:,}")
    print(f"Cache-fil voksede fra {cache_size_before:,} til {cache_size_after:,} unikke titler")
    print(f"Se {fallback_log_path.name} for FULD liste over titler der endte i kanal-fallback (til finjustering)")
    print(f"Se {partial_sport_log_path.name} for FULD liste over titler der blev fanget på DR1/DR2/TV2")
    print("--------------------------------")

    if git_cfg.get("enabled", True):
        prefix = git_cfg.get("commit_message_prefix", "Auto-sync EPG (sport-billeder)")
        git_push(ROOT, f"{prefix} {time.strftime('%Y-%m-%d %H:%M:%S')}")

    print("\n=== Færdig. UHF-klar EPG opdateret. ===")


if __name__ == "__main__":
    main()
