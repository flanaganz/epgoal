#!/usr/bin/env python3
"""
Uafhængig test af LOKAL sport-match-logik (trin 0-3, ingen netværk krævet).
TMDb-fallback-trinnet (trin 4) testes IKKE her, da det kræver en rigtig
TMDB_API_KEY og netværksadgang - test det via en fuld kørsel af
enrich_epg.py i stedet, og tjek "Sport - TMDb-match (nyt)"-linjen i rapporten.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from enrich_epg import SportMatcher  # noqa: E402

TEST_CASES = [
    ("TV 2 Sport HD (D) (T).dk", "Fodbold: 3F Superligaen", "3F Superliga.jpg"),
    ("TV 2 Sport HD (D) (T).dk", "Atletik: Diamond League", "Atletik Diamond League.jpg"),
    ("TV3 Sport HD (D) (T).dk", "Superliga", "3F Superliga.jpg (TV3's faste navn - beholder eget billede)"),

    ("TV 2 Sport HD (D) (T).dk", "Et ukendt TV 2 Sport-program",
     "INGEN lokalt match -> falder til TMDb/kanal-fallback i den fulde pipeline (tv2sport-generic.jpg)"),
    ("TV 2 Sport X HD (D) (T).dk", "Et ukendt TV 2 Sport X-program",
     "INGEN lokalt match -> falder til TMDb/kanal-fallback (TV_2_Sport_X.jpg)"),

    ("TV 2 Sport HD (D) (T).dk", "ATP Tour Highlights", "SPECIFIK override -> ATP Tour Highlights.jpg"),

    ("TV 2 Sport HD (D) (T).dk", "Cykling: Tour de France Femmes - Etaper",
     "SPECIFIK -> Cykling Tour de France Femmes.jpg"),
    ("TV 2 Sport X HD (D) (T).dk", "NFL: Kansas City Chiefs @ Buffalo Bills", "SPECIFIK -> NFL.jpg"),

    ("TV3 Sport HD (D) (T).dk", "Onside",
     "RETTET: intet lokalt match mere (var Fodbold.jpg) -> prøver nu TMDb først i fuld pipeline"),
    ("TV3 Sport HD (D) (T).dk", "1. Division Magasinet",
     "RETTET: intet lokalt match mere -> prøver TMDb først"),
    ("TV3 Sport HD (D) (T).dk", "Var-Rummet",
     "RETTET: intet lokalt match mere -> prøver TMDb først"),
    ("TV3 Sport HD (D) (T).dk", "GP Confidential",
     "Uændret: intet lokalt match -> prøver TMDb først, ellers TV3-generic"),

    ("TV 2 Sport HD (D) (T).dk", "Sendeophold", "SKIP (ingen artwork)"),
    ("TV 2 HD (D) (T).dk", "TV Avisen", "TV2 hovedkanal - INGEN match -> skal IKKE røres"),
    ("BBC News (T).dk", "News Live", "ikke sport-kanal -> IGNORERES fuldstændig"),
]


def main() -> None:
    import json
    cfg = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text(encoding="utf-8"))
    image_base_url = cfg.get("sport", {}).get("image_base_url", "https://example.com/Sport/")

    matcher = SportMatcher(image_base_url)

    print("=== Sport-matching test (kun lokale trin 0-3) ===\n")
    for channel_id, title, expectation in TEST_CASES:
        role_entry = matcher.match_channel(channel_id)
        if role_entry is None:
            print(f"❌ IKKE en sport-kanal | kanal='{channel_id}' | titel='{title}'")
            print(f"   Forventning: {expectation}\n")
            continue

        result = matcher.resolve_local(title)

        if result and result.get("skip"):
            outcome = "SKIP (ingen artwork)"
        elif result:
            outcome = f"backdrop={result.get('backdrop')} | poster={result.get('poster')}"
        elif role_entry.get("role") == "always_sport":
            outcome = "INTET LOKALT MATCH -> vil i fuld pipeline prøve TMDb, derefter kanal-fallback"
        else:
            outcome = "INGEN MATCH -> rører intet (partial_sport uden match)"

        print(f"✅ kanal='{channel_id}' | titel='{title}'")
        print(f"   Forventning: {expectation}")
        print(f"   Resultat:    {outcome}\n")


if __name__ == "__main__":
    main()
