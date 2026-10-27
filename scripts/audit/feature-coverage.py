"""Says what each row of docs/FEATURES.md is checked by, from docs/internal/feature-tests.json.

Usage: python3 scripts/audit/feature-coverage.py [--suite-report run.json] [--scoreboard corpus/reports/scoreboard.json]

A row is "checked" only by test files that the suite actually runs (not ones
listed in tests/not-in-suite.json) and, where given, by how its measured
everyday actions did on the standing set of real workbooks.
"""
import json, os, sys

here = os.path.join(os.path.dirname(__file__), "..", "..")
args = sys.argv[1:]
take = lambda flag, default: args[args.index(flag) + 1] if flag in args else default
mapping = json.load(open(os.path.join(here, "docs/internal/feature-tests.json")))
outside = json.load(open(os.path.join(here, "tests/not-in-suite.json")))
excluded, why = set(), {}
for reason, files in outside.items():
    if reason == "_about":
        continue
    for file in files:
        excluded.add(file); why[file] = reason.split(";")[0].split("(")[0].strip()
tests = sorted(f for f in os.listdir(os.path.join(here, "tests")) if f.endswith(".mjs"))
board_path = take("--scoreboard", os.path.join(here, "corpus/reports/scoreboard.json"))
board = json.load(open(board_path))["workbooks"] if os.path.exists(board_path) else {}
by_action = {}
for row in board.values():
    for item in row.get("actions", []):
        tally = by_action.setdefault(item["action"], [0, 0])
        if item["outcome"] == "skipped":
            continue
        tally[1] += 1
        tally[0] += item["outcome"] == "ok"

result = {}
for tab, rows in mapping.items():
    if tab == "_about":
        continue
    for feature, spec in rows.items():
        run, not_run = [], []
        for fragment in spec["tests"]:
            for file in tests:
                name = file[:-4].replace(".test", "")
                if name == fragment or name.startswith(fragment + "-") and fragment in ("compat-xlsx-calculation",):
                    (not_run if file in excluded else run).append(file)
        worked = sum(by_action.get(a, [0, 0])[0] for a in spec["actions"])
        tried = sum(by_action.get(a, [0, 0])[1] for a in spec["actions"])
        parts = []
        if run:
            parts.append(f"{len(run)} test file{'s' if len(run) > 1 else ''}")
        if tried:
            parts.append(f"{worked} of {tried} on real workbooks ({worked * 100 // tried}%)")
        if not parts:
            parts.append("not checked without the app window" if not not_run else f"tests exist but are not run ({why[not_run[0]]})")
        result.setdefault(tab, {})[feature] = {"text": "; ".join(parts), "run": run, "notRun": not_run, "worked": worked, "tried": tried}
json.dump(result, sys.stdout, indent=1, ensure_ascii=False)
