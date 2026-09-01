@echo off
setlocal enabledelayedexpansion

rem === opdater_alt.bat ===
rem Henter frisk data, beriger sport + danske backdrops, og bygger/opdaterer
rem BEGGE review-Excel-filer - som til sidst aabnes automatisk til gennemgang.
rem
rem   1) enrich_epg.py                  - henter frisk EPG og beriger SPORT
rem                                        (skriver ogsaa logs som export_sport_review.py bruger)
rem   2) danish_backdrops.py             - tilfoejer danske TMDb-backdrops (RESTEN)
rem                                        + injicerer allerede GODKENDTE (X) valg
rem                                        fra danish_artwork_review.xlsx
rem   3) export_sport_review.py          - bygger/opdaterer sport_artwork_review.xlsx
rem   4) export_danish_artwork_review.py - bygger/opdaterer danish_artwork_review.xlsx
rem   5) Aabner BEGGE Excel-filer automatisk
rem
rem NAAR DU HAR UDFYLDT OG GEMT BEGGE FILER: koer "gem_mine_valg.bat"
rem for at skrive dine valg tilbage og bage dem ind i XML'erne.
rem
rem Raekkefoelgen 1->2 er vigtig: enrich_epg.py overskriver output/denmarkX.xml
rem med friske downloads hver gang, saa danish_backdrops.py SKAL koere bagefter,
rem ellers forsvinder de danske backdrops ved naeste sports-opdatering.
rem Raekkefoelgen 3 og 4 SKAL koere efter 1 og 2, da de laeser de logs/cache-filer
rem som enrich_epg.py og danish_backdrops.py netop har genereret/opdateret.
rem
rem Dobbeltklik denne fil, eller koer den fra en almindelig PowerShell/CMD-prompt.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo ================================================
echo  TRIN 1/4: Sport-berigelse (enrich_epg.py)
echo ================================================
echo.

python3 scripts\enrich_epg.py
if errorlevel 1 (
    echo.
    echo [FEJL] enrich_epg.py fejlede - stopper her.
    echo De oevrige trin bliver IKKE koert, da sport-data ikke er opdateret.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 2/4: Danske backdrops (danish_backdrops.py)
echo ================================================
echo.

python3 scripts\danish_backdrops.py
if errorlevel 1 (
    echo.
    echo [ADVARSEL] danish_backdrops.py fejlede eller blev afbrudt.
    echo Sport-data ER opdateret korrekt - kun de danske backdrops mangler.
    echo Du kan koere "python3 scripts\danish_backdrops.py" igen senere for at fortsaette.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 3/4: Byg sport-review (export_sport_review.py)
echo ================================================
echo.

python3 scripts\export_sport_review.py
if errorlevel 1 (
    echo.
    echo [ADVARSEL] export_sport_review.py fejlede - sport_artwork_review.xlsx
    echo er muligvis IKKE opdateret. Sport- og danske backdrops-data er dog
    echo allerede opdateret korrekt fra trin 1-2.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 4/4: Byg dansk-review (export_danish_artwork_review.py)
echo ================================================
echo.

python3 scripts\export_danish_artwork_review.py
if errorlevel 1 (
    echo.
    echo [ADVARSEL] export_danish_artwork_review.py fejlede - danish_artwork_review.xlsx
    echo er muligvis IKKE opdateret.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  FAERDIG! Aabner review-filerne til gennemgang ...
echo ================================================
echo.
echo Naeste skridt:
echo   1. Vaelg billeder / marker godkendt (X) i BEGGE filer
echo   2. GEM begge filer
echo   3. Koer "gem_mine_valg.bat" for at gemme valgene og opdatere XML'erne
echo.

start "" "%SCRIPT_DIR%data\sport_artwork_review.xlsx"
start "" "%SCRIPT_DIR%data\danish_artwork_review.xlsx"

pause
