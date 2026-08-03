#!/usr/bin/env python3
"""Uafhængig test af sport-match-logikken (ingen netværk/git/TMDb krævet)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from enrich_epg import SportMatcher  # noqa: E402

TEST_CASES = [
    # --- Dine faktisk uploadede billeder ---
    ("TV 2 Sport HD (D) (T).dk", "Fodbold: 3F Superligaen",
     "SPECIFIK override -> 3F Superliga.jpg (ikke generisk Fodbold.jpg)"),
    ("TV 2 Sport HD (D) (T).dk", "Atletik: Diamond League",
     "SPECIFIK override -> Atletik Diamond League.jpg"),
    ("TV 2 Sport X HD (D) (T).dk", "UEFA Fodbold-EM U19",
     "Ingen eksakt override -> falder til generisk Fodbold.jpg via nøgleord"),
    ("TV3 Sport HD (D) (T).dk", "Superliga",
     "TV3's faste programnavn -> genbruger 3F Superliga.jpg"),

    # --- Kategorier UDEN billede endnu -> skal falde videre, ikke brække ---
    ("TV 2 Sport HD (D) (T).dk", "Cykling: Tour de France Femmes - Etaper",
     "Intet cykling-billede endnu -> INGEN match her (kanal-generic mangler også -> 'mangler billede')"),
    ("TV3 Sport HD (D) (T).dk", "GP Confidential",
     "Override findes men peger på null -> falder videre (kanal-generic mangler også -> 'mangler billede')"),

    # --- Skip / kontrol ---
    ("TV 2 Sport HD (D) (T).dk", "Sendeophold", "SKIP (ingen artwork)"),
    ("TV 2 HD (D) (T).dk", "Fodbold: 3F Superligaen", "TV2 hovedkanal - matcher via samme override"),
    ("TV 2 HD (D) (T).dk", "TV Avisen", "TV2 hovedkanal - INGEN match -> skal IKKE røres"),
    ("BBC News (T).dk", "News Live", "ikke sport-kanal -> IGNORERES fuldstændig"),
]


def main() -> None:
    import json
    cfg = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text(encoding="utf-8"))
    image_base_url = cfg.get("sport", {}).get("image_base_url", "https://example.com/sport-images/")

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
                outcome = "MANGLER BILLEDE (hverken specifikt match eller kanal-generic er uploadet endnu)"
        else:
            outcome = "INGEN MATCH -> rører intet (partial_sport uden match)"

        print(f"✅ kanal='{channel_id}' | titel='{title}'")
        print(f"   Forventning: {expectation}")
        print(f"   Resultat:    {outcome}\n")


if __name__ == "__main__":
    main()
