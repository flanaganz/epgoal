@echo off
setlocal enabledelayedexpansion

rem === gem_mine_valg.bat ===
rem Koer denne EFTER du har udfyldt og GEMT BEGGE:
rem   data\sport_artwork_review.xlsx
rem   data\danish_artwork_review.xlsx
rem
rem Den skriver dine valg tilbage til systemet og baeger dem ind i XML'erne:
rem   1) import_sport_review.py - skriver sport-billedvalg til
rem                                sport_program_overrides.json
rem   2) enrich_epg.py           - henter frisk EPG + baeger sport-valgene ind
rem   3) danish_backdrops.py     - injicerer NYE godkendte (X) danske backdrops
rem                                (danish_artwork_review.xlsx's X-markeringer
rem                                laeses direkte af scriptet - der findes IKKE
rem                                noget separat "import"-script for dem)
rem
rem Dobbeltklik denne fil, eller koer den fra en almindelig PowerShell/CMD-prompt.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo ================================================
echo  TRIN 1/3: Gem sport-billedvalg (import_sport_review.py)
echo ================================================
echo.

python3 scripts\import_sport_review.py
if errorlevel 1 (
    echo.
    echo [FEJL] import_sport_review.py fejlede - stopper her.
    echo Dine sport-valg er IKKE gemt endnu.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 2/3: Sport-berigelse (enrich_epg.py)
echo ================================================
echo.

python3 scripts\enrich_epg.py
if errorlevel 1 (
    echo.
    echo [FEJL] enrich_epg.py fejlede - stopper her.
    echo Dine sport-valg ER gemt i sport_program_overrides.json,
    echo men er IKKE baget ind i XML-filerne endnu.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 3/3: Danske backdrops (danish_backdrops.py)
echo ================================================
echo.

python3 scripts\danish_backdrops.py
if errorlevel 1 (
    echo.
    echo [ADVARSEL] danish_backdrops.py fejlede eller blev afbrudt.
    echo Sport-valgene ER gemt og baget ind korrekt - kun de nye danske
    echo backdrop-godkendelser mangler. Koer "python3 scripts\danish_backdrops.py"
    echo igen senere for at fortsaette.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  FAERDIG! Dine valg er gemt, og XML'erne er opdateret.
echo ================================================
echo.
pause
