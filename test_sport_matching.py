#!/usr/bin/env python3
"""Uafhængig test af sport-match-logikken (ingen netværk/git/TMDb krævet)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from enrich_epg import SportMatcher  # noqa: E402

TEST_CASES = [
    # --- Eksisterende, allerede bekræftede matches ---
    ("TV 2 Sport HD (D) (T).dk", "Fodbold: 3F Superligaen", "3F Superliga.jpg"),
    ("TV 2 Sport HD (D) (T).dk", "Atletik: Diamond League", "Atletik Diamond League.jpg"),
    ("TV3 Sport HD (D) (T).dk", "Superliga", "3F Superliga.jpg (TV3's faste navn)"),

    # --- NYT: TV 2 Sport vs TV 2 Sport X skal nu have HVER SIN generic ---
    ("TV 2 Sport HD (D) (T).dk", "Et ukendt TV 2 Sport-program",
     "KANAL-FALLBACK -> tv2sport-generic.jpg (IKKE TV_2_Sport_X.jpg)"),
    ("TV 2 Sport X HD (D) (T).dk", "Et ukendt TV 2 Sport X-program",
     "KANAL-FALLBACK -> TV_2_Sport_X.jpg (IKKE tv2sport-generic.jpg)"),

    # --- NYT: ATP Tour Highlights - KUN dette specifikke program ---
    ("TV 2 Sport HD (D) (T).dk", "ATP Tour Highlights",
     "SPECIFIK override -> ATP Tour Highlights.jpg"),
    ("TV 2 Sport HD (D) (T).dk", "Tennis: Wimbledon - Kampe",
     "IKKE ATP-billedet - almindelig tennis har intet generisk billede -> falder til kanal-generic"),

    # --- NYT: Tour de France Femmes specifikt, almindelig cykling IKKE ---
    ("TV 2 Sport HD (D) (T).dk", "Cykling: Tour de France Femmes - Etaper",
     "SPECIFIK -> Cykling Tour de France Femmes.jpg"),
    ("TV 2 Sport HD (D) (T).dk", "Cykling: Giro d'Italia - Etaper",
     "IKKE TdF Femmes-billedet - almindelig cykling har intet generisk billede -> falder til kanal-generic"),

    # --- NYT: NFL ---
    ("TV 2 Sport X HD (D) (T).dk", "NFL: Kansas City Chiefs @ Buffalo Bills",
     "SPECIFIK -> NFL.jpg"),

    # --- Kontrol ---
    ("TV 2 Sport HD (D) (T).dk", "Sendeophold", "SKIP (ingen artwork)"),
    ("TV 2 HD (D) (T).dk", "TV Avisen", "TV2 hovedkanal - INGEN match -> skal IKKE røres"),
    ("BBC News (T).dk", "News Live", "ikke sport-kanal -> IGNORERES fuldstændig"),
]


def main() -> None:
    import json
    cfg = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text(encoding="utf-8"))
    image_base_url = cfg.get("sport", {}).get("image_base_url", "https://example.com/Sport/")

    matcher = SportMatcher(image_base_url)

    print("=== Sport-matching test ===\n")
    for channel_id, title, expectation in TEST_CASES:
        role_entry = matcher.match_channel(channel_id)
        if role_entry is None:
            print(f"❌ IKKE en sport-kanal | kanal='{channel_id}' | titel='{title}'")
            print(f"   Forventning: {expectation}\n")
            continue

        result = matcher.resolve(title)

        if result and result.get("skip"):
            outcome = "SKIP (ingen artwork)"
        elif result:
            outcome = f"backdrop={result.get('backdrop')} | poster={result.get('poster')}"
        elif role_entry.get("role") == "always_sport":
            fb, fp = role_entry.get("default_backdrop"), role_entry.get("default_poster")
            if fb or fp:
                fallback = matcher._image_urls(fb, fp)
                outcome = f"KANAL-FALLBACK backdrop={fallback['backdrop']} | poster={fallback['poster']}"
            else:
                outcome = "MANGLER BILLEDE"
        else:
            outcome = "INGEN MATCH -> rører intet (partial_sport uden match)"

        print(f"✅ kanal='{channel_id}' | titel='{title}'")
        print(f"   Forventning: {expectation}")
        print(f"   Resultat:    {outcome}\n")


if __name__ == "__main__":
    main()
