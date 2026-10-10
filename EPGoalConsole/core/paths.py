from pathlib import Path
import json
CONSOLE_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = CONSOLE_ROOT / 'config' / 'gui_settings.json'
def settings():
    return json.loads(SETTINGS_FILE.read_text(encoding='utf-8'))
def paths():
    root = Path(settings()['epgoal_root'])
    return {
        'root': root,
        'scripts': root / 'scripts',
        'data': root / 'data',
        'output': root / 'output_epgshare',
        'priority': root / 'data' / 'channel_priority_v2.xlsx',
        'merge_log': root / 'data' / 'epgshare_merge_log.json',
        'normalize_log': root / 'data' / 'normalize_uhf_channel_ids_log.json',
        'danish_log': root / 'data' / 'danish_backdrops_run_log.json',
        'epgoal': root / 'output_epgshare' / 'epgoal.xml',
        'pipeline': root / 'scripts' / 'opdater_epgoal.py',
        'save_choices': root / 'scripts' / 'gem_epgoal_valg.py',
    }
