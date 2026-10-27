"""Writes the 'Checked by' column of docs/FEATURES.md from what is actually run.

Usage: python3 scripts/audit/feature-table.py     (rewrites docs/FEATURES.md in place)

The column comes from scripts/audit/feature-coverage.py: the test files the
suite runs for the row, and how the row's everyday actions did on the standing
set of real workbooks. Everything else in the file is left as written.
"""
import json, os, re, subprocess, sys

here = os.path.join(os.path.dirname(__file__), "..", "..")
coverage = json.loads(subprocess.check_output([sys.executable, os.path.join(here, "scripts/audit/feature-coverage.py")]))
path = os.path.join(here, "docs/FEATURES.md")
lines = open(path, encoding="utf-8").read().split("\n")
out, tab, missing = [], None, []
for line in lines:
    heading = re.match(r"^## (.+)$", line)
    if heading:
        tab = heading.group(1)
    cells = [c.strip() for c in line.strip().strip("|").split("|")] if line.startswith("|") else None
    if tab in coverage and cells:
        if cells[0] == "Feature":
            cells = cells[:4] + ["Checked by"]
        elif set(cells[0]) <= set("-: "):
            cells = ["---"] * 5
        elif cells[0] in coverage[tab]:
            cells = cells[:4] + [coverage[tab][cells[0]]["text"]]
        else:
            if cells[-1] != "Not planned":
                missing.append(f"{tab} / {cells[0]}")
            cells = cells[:4] + ["" if len(cells) < 5 else cells[4]]
        line = "| " + " | ".join(cells) + " |"
    out.append(line)
open(path, "w", encoding="utf-8").write("\n".join(out))
rows = [r for t in coverage.values() for r in t.values()]
print(f"{len(rows)} rows: {sum(bool(r['run']) for r in rows)} have tests the suite runs, "
      f"{sum(bool(r['tried']) for r in rows)} are measured on real workbooks, "
      f"{sum(not r['run'] and not r['tried'] for r in rows)} are checked by neither")
for name in missing:
    print("no entry in feature-tests.json for:", name)
