@echo off
REM ============================================================
REM kor_statistik.bat — statistik-kørsel for EPGoal
REM
REM Kører den POTENTIELLE artwork-dækning (enrich_non_sport_with_tmdb=true
REM for alle 6 kilder), UDEN at røre produktionens output/ eller
REM data/channel_health.json. Bruger config-stats.json og skriver til
REM output_stats/ + data/channel_health_stats.json.
REM
REM Rører ALDRIG:
REM   - output\ (produktion, det UHF rent faktisk bruger)
REM   - data\channel_health.json (produktions-snapshot)
REM   - danish_backdrops.py / danish_artwork_review.xlsx
REM
REM Brug: dobbeltklik filen, eller kør fra PowerShell:
REM   .\kor_statistik.bat
REM ============================================================

cd /d "%~dp0"

echo.
echo === EPGoal: STATISTIK-koersel (potentiel TMDb-daekning) ===
echo Roerer IKKE produktion - kun output_stats\ og channel_health_stats.json
echo.

echo [1/2] Koerer enrich_epg.py med config-stats.json ...
python scripts\enrich_epg.py config-stats.json
if errorlevel 1 (
    echo.
    echo *** FEJL: enrich_epg.py fejlede - stopper her. ***
    pause
    exit /b 1
)

echo.
echo [2/2] Koerer channel_health.py --stats ...
python scripts\channel_health.py --stats
if errorlevel 1 (
    echo.
    echo *** FEJL: channel_health.py fejlede. ***
    pause
    exit /b 1
)

echo.
echo === Faerdig! ===
echo Resultat gemt i: data\channel_health_stats.json
echo XML-filer gemt i: output_stats\
echo.
pause
