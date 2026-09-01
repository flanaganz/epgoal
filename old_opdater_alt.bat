@echo off
setlocal enabledelayedexpansion

rem === opdater_alt.bat ===
rem Kører den fulde EPG-opdatering i korrekt rækkefølge:
rem   1) enrich_epg.py       - henter frisk EPG fra Open-EPG og beriger SPORT
rem   2) danish_backdrops.py - tilføjer danske TMDb-backdrops til RESTEN
rem
rem Rækkefølgen er vigtig: enrich_epg.py overskriver output/denmarkX.xml med
rem friske downloads hver gang, så danish_backdrops.py SKAL køre bagefter,
rem ellers forsvinder de danske backdrops ved næste sports-opdatering.
rem
rem Dobbeltklik denne fil, eller kør den fra en almindelig PowerShell/CMD-prompt.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo ================================================
echo  TRIN 1/2: Sport-berigelse (enrich_epg.py)
echo ================================================
echo.

python3 scripts\enrich_epg.py
if errorlevel 1 (
    echo.
    echo [FEJL] enrich_epg.py fejlede - stopper her.
    echo Danish_backdrops.py bliver IKKE koert, da sport-data ikke er opdateret.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 2/2: Danske backdrops (danish_backdrops.py)
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
echo  FAERDIG! Sport + danske backdrops er opdateret.
echo ================================================
echo.
pause
