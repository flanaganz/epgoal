#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ret live-log i EPGoal Console ved at køre child-processen ubufferet."""
from pathlib import Path
from datetime import datetime
import ast
import re
import shutil

ROOT = Path(r"C:\EPGoal")
RUNNER = ROOT / "EPGoalConsole" / "core" / "runner.py"
REPORTS = ROOT / "reports"


def main():
    if not RUNNER.exists():
        raise SystemExit(f"FEJL: Mangler {RUNNER}")

    original = RUNNER.read_text(encoding="utf-8")
    source = original

    # Gør child-Python ubufferet. -u er den vigtigste rettelse.
    source, count_cmd = re.subn(
        r'\[sys\.executable\s*,\s*str\(target\)\]',
        '[sys.executable, "-u", str(target)]',
        source,
        count=1,
    )

    # Miljøvariablen sikrer også ubufferet stdout/stderr i child-processen.
    if 'env["PYTHONUNBUFFERED"]' not in source:
        marker = 'env["EPGOAL_SKIP_GIT"] = "1"'
        if marker not in source:
            raise SystemExit(
                "FEJL: Console-miljøblokken blev ikke fundet i runner.py. "
                "Ingen fil blev ændret."
            )
        source = source.replace(
            marker,
            marker + '\n    env["PYTHONUNBUFFERED"] = "1"',
            1,
        )

    # Fortæl straks GUI'en, at processen faktisk er startet.
    if 'PROCES STARTET: PID=' not in source:
        pattern = r'(\n\s*p\s*=\s*subprocess\.Popen\([^\n]+\)\n)'
        match = re.search(pattern, source)
        if not match:
            raise SystemExit(
                "FEJL: subprocess.Popen-linjen blev ikke fundet. "
                "Ingen fil blev ændret."
            )
        indent = re.match(r'\n(\s*)', match.group(1)).group(1)
        insertion = match.group(1) + f'{indent}on_line(f"PROCES STARTET: PID={{p.pid}} | {{target}}")\n'
        source = source[:match.start()] + insertion + source[match.end():]

    if count_cmd == 0 and '"-u", str(target)' not in source:
        raise SystemExit(
            "FEJL: Kunne ikke tilføje Python -u. Ingen fil blev ændret."
        )

    ast.parse(source, filename=str(RUNNER))

    REPORTS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = REPORTS / f"console_runner_before_live_log_fix_{stamp}.py"
    shutil.copy2(RUNNER, backup)
    RUNNER.write_text(source, encoding="utf-8")
    ast.parse(RUNNER.read_text(encoding="utf-8"), filename=str(RUNNER))

    print("=== LIVE-LOG RETTET ===")
    print(f"Opdateret : {RUNNER}")
    print(f"Backup    : {backup}")
    print("Child-processen kører nu med Python -u og PYTHONUNBUFFERED=1.")


if __name__ == "__main__":
    main()
