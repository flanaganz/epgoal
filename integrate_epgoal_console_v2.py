#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Robust integration af EPGoal Console i den aktuelle produktionspipeline."""
from pathlib import Path
from datetime import datetime
import ast, re, shutil

ROOT = Path(r"C:\EPGoal")
TARGET = ROOT / "scripts" / "opdater_epgoal.py"
RUNNER = ROOT / "EPGoalConsole" / "core" / "runner.py"
REPORTS = ROOT / "reports"

OVERRIDE_CODE = r'''
        # Kildevalg fra EPGoal Console. Uden aktiv override bruges normal prioritet.
        source_overrides = load_json(SOURCE_OVERRIDES_FILE, {})
        override = source_overrides.get(kanonisk_navn, {})
        override_enabled = bool(override.get("enabled", True))
        override_source = str(override.get("source", "auto")).strip().lower()
        override_id = str(override.get("source_channel_id", "")).strip()

        if override_enabled and override_source == "exclude":
            for cid in matched_epgshare_ids:
                for programme in list(epgshare_programmes_by_channel.get(cid, [])):
                    if programme in root:
                        root.remove(programme)
                for channel in list(root.findall("channel")):
                    if channel.get("id", "") == cid:
                        root.remove(channel)
            report.setdefault("excluded_by_console", []).append({
                "kanal": kanonisk_navn,
                "kildeoverride": "exclude",
            })
            print(f"Console-override: udelader {kanonisk_navn}")
            continue

        forced = None
        if override_enabled and override_source in {"epgshare", "openepg", "bss"}:
            if override_source == "epgshare":
                source_map = epgshare_programmes_by_channel
                source_stats = epgshare_stats
                candidate_ids = matched_epgshare_ids
            elif override_source == "openepg":
                source_map = openepg_programmes_by_channel
                source_stats = openepg_stats
                candidate_ids = list(source_map)
            else:
                source_map = bss_programmes_by_channel
                source_stats = bss_stats
                candidate_ids = list(source_map)

            selected_id = None
            if override_id in source_map:
                selected_id = override_id
            elif override_id:
                selected_id = next(
                    (cid for cid in candidate_ids if normalize_id(cid) == normalize_id(override_id)),
                    None,
                )
            elif override_source == "epgshare" and matched_epgshare_ids:
                selected_id = max(
                    matched_epgshare_ids,
                    key=lambda cid: source_stats.get(cid, {}).get("count", 0),
                )
            else:
                selected_id, _ = find_best_candidate(tokens, source_map, source_stats)

            selected_stat = source_stats.get(
                selected_id,
                {"count": 0, "span_hours": 0},
            ) if selected_id else {"count": 0, "span_hours": 0}

            if selected_id and selected_stat["count"] > 0:
                forced = (override_source, selected_id, selected_stat, source_map)
            else:
                print(
                    f"ADVARSEL: Console-override for {kanonisk_navn} kunne ikke finde "
                    f"{override_source}:{override_id or '(automatisk ID)'}. "
                    "Bruger normal prioritet.",
                    file=sys.stderr,
                )

        if forced:
            best_source, best_id, best_stat, best_programmes_map = forced

            if best_source == "epgshare":
                report["kept_as_is"].append({
                    "kanal": kanonisk_navn,
                    "kilde": "epgshare",
                    "programmer": best_stat["count"],
                    "vindue_timer": round(best_stat["span_hours"], 1),
                    "tvunget_kildeoverride": True,
                    "kilde_channel_id": best_id,
                })
                print(f"Console-override: {kanonisk_navn} -> epgshare:{best_id}")
                continue

            if matched_epgshare_ids:
                primary_id = max(
                    matched_epgshare_ids,
                    key=lambda cid: epgshare_stats.get(cid, {}).get("count", 0),
                )
            else:
                primary_id = f"{normalize_id(kanonisk_navn)}.dk"
                new_channel = ET.Element("channel", {"id": primary_id})
                display_name = ET.SubElement(new_channel, "display-name")
                display_name.text = kanonisk_navn
                root.append(new_channel)

            for cid in matched_epgshare_ids:
                for programme in list(epgshare_programmes_by_channel.get(cid, [])):
                    if programme in root:
                        root.remove(programme)

            inserted = 0
            for programme in best_programmes_map.get(best_id, []):
                clone = copy.deepcopy(programme)
                clone.set("channel", primary_id)
                strip_artwork(clone)
                root.append(clone)
                inserted += 1

            report["supplemented"].append({
                "kanal": kanonisk_navn,
                "epgshare_programmer_foer": epgshare_count,
                "epgshare_vindue_timer_foer": round(epgshare_span, 1),
                "kilde_brugt": best_source,
                "kilde_channel_id": best_id,
                "programmer_indsat": inserted,
                "vindue_timer_efter": round(best_stat["span_hours"], 1),
                "primaer_channel_id": primary_id,
                "tvunget_kildeoverride": True,
            })
            print(
                f"Console-override: {kanonisk_navn} -> "
                f"{best_source}:{best_id} ({inserted:,} programmer)"
            )
            continue
'''


def add_constants(source: str) -> str:
    if "SOURCE_OVERRIDES_FILE" in source:
        return source
    pattern = r'^(GIT_COMMIT_PREFIX\s*=.*)$'
    replacement = (
        r'\1\n'
        'CONSOLE_CONFIG_DIR = ROOT / "EPGoalConsole" / "config"\n'
        'SOURCE_OVERRIDES_FILE = CONSOLE_CONFIG_DIR / "source_overrides.json"'
    )
    updated, count = re.subn(pattern, replacement, source, count=1, flags=re.M)
    if count != 1:
        raise SystemExit("FEJL: Kunne ikke finde GIT_COMMIT_PREFIX. Ingen filer blev ændret.")
    return updated


def add_merge_override(source: str) -> str:
    if "# Kildevalg fra EPGoal Console." in source:
        return source
    function_start = source.find("def step3_merge_sources()")
    function_end = source.find("\n# ---------------------------------------------------------------------------", function_start)
    if function_start < 0 or function_end < 0:
        raise SystemExit("FEJL: Kunne ikke afgrænse step3_merge_sources(). Ingen filer blev ændret.")
    block = source[function_start:function_end]
    match = re.search(
        r'(\n\s{8}epgshare_span\s*=\s*max\(.*?\n\s{8}\)\n)',
        block,
        flags=re.S,
    )
    if not match:
        raise SystemExit("FEJL: Kunne ikke finde epgshare_span i step3. Ingen filer blev ændret.")
    insert_at = function_start + match.end()
    return source[:insert_at] + OVERRIDE_CODE + source[insert_at:]


def add_git_skip(source: str) -> str:
    marker = 'os.environ.get("EPGOAL_SKIP_GIT"'
    if marker in source:
        return source
    pattern = r'(def step7_git_push\(\)\s*->\s*None:\s*\n\s+section\([^\n]+\)\s*\n)'
    addition = (
        '    if os.environ.get("EPGOAL_SKIP_GIT", "").strip().lower() in '
        '{"1", "true", "yes"}:\n'
        '        print("Git springes over: Byg og valider blev startet fra EPGoal Console.")\n'
        '        return\n'
    )
    updated, count = re.subn(pattern, r'\1' + addition, source, count=1)
    if count != 1:
        raise SystemExit("FEJL: Kunne ikke integrere Git-skip. Ingen filer blev ændret.")
    return updated


def add_console_review_mode(source: str) -> str:
    marker = 'os.environ.get("EPGOAL_CONSOLE_MODE"'
    if marker in source:
        return source
    pattern = r'(def open_file_for_review\(path:\s*Path\)\s*->\s*None:\s*\n)'
    addition = (
        '    if os.environ.get("EPGOAL_CONSOLE_MODE", "").strip().lower() in '
        '{"1", "true", "yes"}:\n'
        '        print(f"Console-tilstand: åbner ikke {path.name} automatisk.")\n'
        '        return\n'
    )
    updated, count = re.subn(pattern, r'\1' + addition, source, count=1)
    if count != 1:
        raise SystemExit("FEJL: Kunne ikke integrere Console-reviewtilstand. Ingen filer blev ændret.")
    return updated


def patch_runner(runner: str) -> str:
    if 'env["EPGOAL_SKIP_GIT"]' in runner:
        return runner
    pattern = r'(?m)^(\s*)p\s*=\s*subprocess\.Popen\('
    match = re.search(pattern, runner)
    if not match:
        raise SystemExit("FEJL: Kunne ikke finde subprocess.Popen i runner.py. Ingen filer blev ændret.")
    indent = match.group(1)
    prefix = (
        f'{indent}import os\n'
        f'{indent}env = os.environ.copy()\n'
        f'{indent}env["EPGOAL_CONSOLE_MODE"] = "1"\n'
        f'{indent}env["EPGOAL_SKIP_GIT"] = "1"\n'
    )
    runner = runner[:match.start()] + prefix + runner[match.start():]
    runner = re.sub(
        r'(p\s*=\s*subprocess\.Popen\(\[sys\.executable,str\(target\)\],cwd=paths\(\)\[\'root\'\],)',
        r'\1env=env,',
        runner,
        count=1,
    )
    return runner


def main():
    if not TARGET.exists() or not RUNNER.exists():
        raise SystemExit("FEJL: opdater_epgoal.py eller Console runner.py mangler.")

    original_source = TARGET.read_text(encoding="utf-8")
    original_runner = RUNNER.read_text(encoding="utf-8")

    source = add_constants(original_source)
    source = add_merge_override(source)
    source = add_git_skip(source)
    source = add_console_review_mode(source)
    runner = patch_runner(original_runner)

    ast.parse(source, filename=str(TARGET))
    ast.parse(runner, filename=str(RUNNER))

    REPORTS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_source = REPORTS / f"opdater_epgoal_before_console_integration_{stamp}.py"
    backup_runner = REPORTS / f"console_runner_before_integration_{stamp}.py"
    shutil.copy2(TARGET, backup_source)
    shutil.copy2(RUNNER, backup_runner)

    TARGET.write_text(source, encoding="utf-8")
    RUNNER.write_text(runner, encoding="utf-8")

    ast.parse(TARGET.read_text(encoding="utf-8"), filename=str(TARGET))
    ast.parse(RUNNER.read_text(encoding="utf-8"), filename=str(RUNNER))

    print("=== EPGOAL CONSOLE INTEGRATION GENNEMFØRT ===")
    print(f"Produktionsscript: {TARGET}")
    print(f"Console runner   : {RUNNER}")
    print(f"Backup 1         : {backup_source}")
    print(f"Backup 2         : {backup_runner}")
    print("GUI-kørsler bruger nu source_overrides.json og springer Git-push over.")
    print("Manuel kørsel af opdater_epgoal.py beholder normal Git-adfærd.")


if __name__ == "__main__":
    main()
