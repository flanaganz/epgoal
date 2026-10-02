@echo off
setlocal enabledelayedexpansion

rem === opdater_epgshare.bat ===
rem EPGShare01-POC - koerer hele kaeden isoleret fra produktionen (opdater_alt.bat).
rem
rem   1) epgshare_download.py         - henter frisk EPGShare01 DK1 (gzip, udpakkes)
rem   2) epgshare_filter.py           - filtrerer til kun dine X-markerede kanaler
rem                                     (data\channel_priority.xlsx) og rydder evt.
rem                                     icon/backdrop der allerede laa i kilden
rem   3) epgshare_merge_openepg.py    - supplerer/erstatter kanaler hvor EPGShare01
rem                                     har for kort eller intet EPG-vindue med
rem                                     frisk data fra OpenEPG (samme 6 kilder som
rem                                     produktionen bruger) - se
rem                                     data\epgshare_merge_log.json for detaljer
rem   4) enrich_epg_epgshare.py       - beriger SPORT (samme data/logik som produktion:
rem                                     sport_channels.json, sport_categories.json,
rem                                     sport_program_overrides.json osv.)
rem   5) danish_backdrops_epgshare.py - tilfoejer danske TMDb-backdrops (RESTEN)
rem                                     + injicerer allerede GODKENDTE (X) valg fra
rem                                     data\danish_artwork_review.xlsx (deles med
rem                                     produktionen - samme cache, samme godkendelser)
rem
rem Trin 4 og 5 committer og pusher til GitHub automatisk (git.enabled=true i
rem config-epgshare.json), saa output_epgshare\epgshare_merged.xml altid ligger
rem friskt tilgaengeligt paa GitHub bagefter.
rem
rem UHF-URL (indsaet som kilde i UHF):
rem   https://raw.githubusercontent.com/flanaganz/epgoal/main/output_epgshare/epgshare_merged.xml
rem
rem Dobbeltklik denne fil, eller koer den fra en almindelig PowerShell/CMD-prompt.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo ================================================
echo  TRIN 1/5: Henter EPGShare01 DK1 (epgshare_download.py)
echo ================================================
echo.

python scripts\epgshare_download.py
if errorlevel 1 (
    echo.
    echo [FEJL] epgshare_download.py fejlede - stopper her.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 2/5: Filtrerer til dine kanaler (epgshare_filter.py)
echo ================================================
echo.

python scripts\epgshare_filter.py
if errorlevel 1 (
    echo.
    echo [FEJL] epgshare_filter.py fejlede - stopper her.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 3/5: Supplerer huller med OpenEPG (epgshare_merge_openepg.py)
echo ================================================
echo.

python scripts\epgshare_merge_openepg.py
if errorlevel 1 (
    echo.
    echo [FEJL] epgshare_merge_openepg.py fejlede - stopper her.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 4/5: Sport-berigelse (enrich_epg_epgshare.py)
echo ================================================
echo.

python scripts\enrich_epg_epgshare.py config-epgshare.json
if errorlevel 1 (
    echo.
    echo [FEJL] enrich_epg_epgshare.py fejlede - stopper her.
    echo De oevrige trin bliver IKKE koert.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  TRIN 5/5: Danske backdrops (danish_backdrops_epgshare.py)
echo ================================================
echo.

python scripts\danish_backdrops_epgshare.py
if errorlevel 1 (
    echo.
    echo [ADVARSEL] danish_backdrops_epgshare.py fejlede eller blev afbrudt.
    echo Sport-data ER opdateret korrekt - kun de danske backdrops mangler.
    echo Du kan koere "python scripts\danish_backdrops_epgshare.py" igen senere.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  FAERDIG! EPGShare-POC opdateret og pushet til GitHub.
echo ================================================
echo.
echo UHF-URL (indsaet som kilde i UHF):
echo   https://raw.githubusercontent.com/flanaganz/epgoal/main/output_epgshare/epgshare_merged.xml
echo.
echo Se data\epgshare_merge_log.json for hvilke kanaler der blev suppleret med
echo OpenEPG i dag, og hvilke (hvis nogen) der stadig mangler data begge steder.
echo.

pause
