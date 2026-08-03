# epgoal — sport-first EPG-berigelse

Repo: **https://github.com/flanaganz/epgoal**

## Status på dine sport-billeder

Du har uploadet 3 billeder til `sport-images/`:

| Fil (som du navngav den) | Bruges til |
|---|---|
| `Fodbold.jpg` | Generisk fallback for fodbold-programmer der ikke matcher noget specifikt |
| `3F Superliga.jpg` | SPECIFIK til "Fodbold: 3F Superligaen" + TV3's "Superliga" |
| `Atletik Diamond League.jpg` | SPECIFIK til "Atletik: Diamond League" |

**Mellemrum og stort/lille bogstav i filnavne er OK** — scriptet
URL-encoder automatisk filnavnet (fx `3F Superliga.jpg` bliver til
`3F%20Superliga.jpg` i URL'en), og GitHub raw-URL'er er
case-sensitive, så det er vigtigt at filnavnet i JSON matcher **præcis**
det du har uploadet (stort F i "Fodbold.jpg" er allerede rettet ind
efter det).

**Alt andet mangler stadig billeder** (cykling, tennis, badminton,
golf, motorsport, kanal-generics osv.) — de er sat til `null` i JSON'en
lige nu, hvilket betyder scriptet **ikke** linker til ikke-eksisterende
filer. I stedet tælles de i rapportens "mangler billede"-linje, så du
kan se præcis hvor mange programmer der stadig venter på et billede.

## Sådan tilføjer du flere billeder

1. Upload filen til `sport-images/` på GitHub (browser: "Add file" →
   "Upload files", eller `git add sport-images/ && git commit && git push`
   lokalt)
2. Ret det tilsvarende `null` til dit filnavn i enten:
   - `data/sport_categories.json` (generisk kategori, fx alle cykling-programmer)
   - `data/sport_program_overrides.json` (specifik begivenhed/program)
   - `data/sport_channels.json` (`default_backdrop`/`default_poster` — kanalens sidste sikkerhedsnet)
3. Kør `python3 scripts/test_sport_matching.py` for at se ændringen slå igennem

**Anbefalet rækkefølge at udfylde i:** kanal-generics først
(`tv2sport-generic.jpg`, `tv3sport-generic.jpg` i `sport_channels.json`)
— så har ALT på de kanaler i det mindste et fornuftigt billede, mens du
løbende laver mere specifikke kategori-/event-billeder ovenpå.

## Arkitektur

**Fase 1 (denne udgave):** selv-hostede billeder til sportskanalerne.
**Fase 2 (senere):** TMDb-berigelse af alt ikke-sport, slås til pr.
kilde i `config.json`.

**Alle 6 Open-EPG-filer hentes og behandles hver kørsel**, fordi
sportskanalerne kan ligge i hvilken som helst af dem. Ikke-sport kører
urørt igennem.

```
open-epg.com/denmark1.xml … denmark6.xml
        ▼
scripts/enrich_epg.py
        │  sport-kanal? → match-kæde → <icon>/<backdrop>
        │  ellers       → urørt (fase 1) / TMDb (fase 2)
        ▼
output/denmark1.xml … denmark6.xml
        │  git add / commit / push
        ▼
raw.githubusercontent.com/flanaganz/epgoal/main/output/denmarkX.xml
        ▼
   UHF peger på ALLE 6 her i stedet for open-epg.com direkte
```

## Match-kæde for sport-kanaler

| Trin | Fil | Fanger |
|---|---|---|
| 0. Skip | `sport_skip_titles.json` | "Godnat", "Sendeophold" mv. |
| 1. Eksakt FULD titel | `sport_program_overrides.json` | Både TV3 Sports faste navne OG specifikke begivenheder der skal have et andet billede end deres kategoris generiske fallback |
| 2. Kategori-præfiks | `sport_categories.json` (prefix) | TV 2 Sport(X)'s "Kategori: Begivenhed" |
| 3. Nøgleord | `sport_categories.json` (keywords) | Fanger fx "UEFA Fodbold-EM U19" (længste match vinder) |
| 4. Kanal-fallback | `sport_channels.json` (default_backdrop) | Sidste sikkerhedsnet for `always_sport`-kanaler |

Et match der peger på `null` (billede ikke lavet endnu) tæller
**ikke** som "matched" — scriptet falder automatisk videre til næste
trin, så du aldrig får brækkede billed-links.

**TV2 hovedkanal** (`partial_sport`) rammes KUN hvis trin 0-3 matcher.

## Kom i gang

```bash
cd ~
git clone git@github.com:flanaganz/epgoal.git
cd epgoal
python3 -m venv .venv && source .venv/bin/activate
pip install -r scripts/requirements.txt
python3 scripts/test_sport_matching.py   # valider match-logik uden netværk
python3 scripts/enrich_epg.py            # fuld kørsel: henter, beriger, gemmer, pusher
```

### launchd (automatisk 2×/dag)
```bash
cp launchd/com.simon.epgenricher.plist ~/Library/LaunchAgents/
mkdir -p ~/epgoal/logs
```
Ret `DIT-BRUGERNAVN` i plist-filen, indlæs:
```bash
launchctl load ~/Library/LaunchAgents/com.simon.epgenricher.plist
launchctl start com.simon.epgenricher   # test manuelt
```

### Peg UHF på output-filerne
```
https://raw.githubusercontent.com/flanaganz/epgoal/main/output/denmark1.xml
... osv. til denmark6.xml
```

## Filer du selv vedligeholder

| Fil | Formål |
|---|---|
| `data/sport_channels.json` | Hvilke kanaler er sport, og hvor "strengt" |
| `data/sport_categories.json` | Kategori-præfiks + nøgleord → billedfil |
| `data/sport_program_overrides.json` | Specifikke faste/enkelte programmer → eget billede |
| `data/sport_skip_titles.json` | Titler der slet ikke skal have artwork |
| `sport-images/` | Dine egne billeder |
| `config.json` | Kilde-liste, TMDb-fase-2-toggle, billedstørrelser |
