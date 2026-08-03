# epgoal — sport-first EPG-berigelse (lokal Mac mini-udgave)

Repo: **https://github.com/flanaganz/epgoal**

**Fase 1 (denne udgave):** selv-hostede billeder til sportskanalerne
(TV 2 Sport, TV 2 Sport X, TV3 Sport, delvist TV2), da TMDb stort set
aldrig har poster til live sport.

**Fase 2 (senere, "bygge ovenpå"):** TMDb-berigelse af alt ikke-sport.
Slås til pr. kilde i `config.json`.

## Placering af dine sport-billeder

**Læg dem i `sport-images/` i roden af repoet:**

```
github.com/flanaganz/epgoal
└── sport-images/
    ├── fodbold.jpg
    ├── fodbold-poster.jpg
    ├── cykling.jpg
    ├── tv2sport-generic.jpg
    ├── tv3sport-generic.jpg
    └── ...
```

`config.json` peger allerede på:

```
https://raw.githubusercontent.com/flanaganz/epgoal/main/sport-images/
```

— så du skal IKKE rette noget, blot uploade billeder med de rigtige
filnavne (se `data/sport\_categories.json`, `data/sport\_channels.json`,
`data/sport\_program\_overrides.json` for den fulde liste af forventede
filnavne).

**Upload via browser:** Gå ind i `sport-images`-mappen på GitHub →
"Add file" → "Upload files" → træk billeder ind → commit.

**Upload via terminal (Mac mini):**

```bash
cd \~/epgoal
cp \~/Downloads/fodbold.jpg sport-images/
git add sport-images/
git commit -m "Tilføj sport-billeder"
git push
```

## Hvorfor ALLE 6 filer behandles hver gang

Sportskanalerne kan ligge i **hvilken som helst** af Open-EPGs 6 filer.
Derfor henter og skriver scriptet alle 6 filer hver kørsel — kun
sport-kanaler beriges, resten kører urørt igennem.

```
open-epg.com/denmark1.xml … denmark6.xml
        │  (scriptet henter ALLE 6, hver gang)
        ▼
scripts/enrich\_epg.py
        │  sport-kanal? → match-kæde → <icon>/<backdrop>
        │  ellers       → urørt (fase 1) / TMDb (fase 2)
        ▼
output/denmark1.xml … denmark6.xml
        │  git add / commit / push
        ▼
raw.githubusercontent.com/flanaganz/epgoal/main/output/denmarkX.xml
        │
        ▼
   UHF peger på ALLE 6 her i stedet for open-epg.com direkte
```

## Match-kæde for sport-kanaler

|Trin|Fil|Fanger|
|-|-|-|
|0. Skip|`sport\_skip\_titles.json`|"Godnat", "Sendeophold" mv.|
|1. Eksakt titel|`sport\_program\_overrides.json`|TV3 Sports faste navne: Superliga, Onside, GP Confidential osv.|
|2. Kategori-præfiks|`sport\_categories.json` (prefix)|TV 2 Sport(X)'s "Kategori: Begivenhed" (Cykling:, Badminton: osv.)|
|3. Nøgleord|`sport\_categories.json` (keywords)|Fanger fx "Masser af bordtennis" (længste match vinder, så "bordtennis" ikke fejlmatcher til "tennis")|
|4. Kanal-fallback|`sport\_channels.json` (default\_backdrop)|Kun for `always\_sport`-kanaler — garanterer intet mangler artwork|

**TV2 hovedkanal** (`partial\_sport`) rammes KUN hvis trin 0-3 matcher.

## Kom i gang

### 1\. Klon repoet på din Mac mini

```bash
cd \~
git clone git@github.com:flanaganz/epgoal.git
cd epgoal
```

(Alle projektfilerne du får her fra chatten skal lægges ind i denne
mappe og pushes op, hvis de ikke allerede er der.)

### 2\. Læg sport-billeder i `sport-images/` (se ovenfor)

### 3\. Installér og testkør

```bash
python3 -m venv .venv \&\& source .venv/bin/activate
pip install -r scripts/requirements.txt
python3 scripts/test\_sport\_matching.py   # valider match-logik uden netværk
python3 scripts/enrich\_epg.py            # fuld kørsel: henter, beriger, gemmer, pusher
```

`TMDB\_API\_KEY` er IKKE nødvendig i fase 1.

### 4\. launchd (automatisk 2×/dag)

```bash
cp launchd/com.simon.epgenricher.plist \~/Library/LaunchAgents/
mkdir -p \~/epgoal/logs
```

Ret stierne i den kopierede plist-fil (skift `DIT-BRUGERNAVN` ud), indlæs:

```bash
launchctl load \~/Library/LaunchAgents/com.simon.epgenricher.plist
launchctl start com.simon.epgenricher   # test manuelt
```

### 5\. Peg UHF på output-filerne

```
https://raw.githubusercontent.com/flanaganz/epgoal/main/output/denmark1.xml
https://raw.githubusercontent.com/flanaganz/epgoal/main/output/denmark2.xml
... osv. til denmark6.xml
```

(eller via jsDelivr: `https://cdn.jsdelivr.net/gh/flanaganz/epgoal@main/output/denmarkX.xml`)

## Filer du selv vedligeholder

|Fil|Formål|
|-|-|
|`data/sport\_channels.json`|Hvilke kanaler er sport, og hvor "strengt"|
|`data/sport\_categories.json`|Kategori-præfiks + nøgleord → billedfil|
|`data/sport\_program\_overrides.json`|Faste programnavne uden kategori-præfiks|
|`data/sport\_skip\_titles.json`|Titler der slet ikke skal have artwork|
|`sport-images/`|Dine egne billeder|
|`config.json`|Kilde-liste, TMDb-fase-2-toggle, billedstørrelser|

## Aktivér fase 2 senere (TMDb for ikke-sport)

Sæt `"enrich\_non\_sport\_with\_tmdb": true` på den/de kilder i
`config.json`, og sørg for `TMDB\_API\_KEY` er sat (`.env`-fil).

