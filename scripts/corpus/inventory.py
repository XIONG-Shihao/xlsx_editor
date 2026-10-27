"""Lists what each workbook in the standing set contains, and how often each feature occurs.

Usage: python3 scripts/corpus/inventory.py [--json OUT] [FOLDER]

Reads the package directly (no workbook library), so it reports what is in the
file, not what a particular reader understands.
"""
import json, os, re, sys, zipfile

args = sys.argv[1:]
out = None
if "--json" in args:
    at = args.index("--json"); out = args[at + 1]; del args[at:at + 2]
folder = args[0] if args else os.path.join(os.path.dirname(__file__), "..", "..", "corpus", "files")

FEATURES = [
    "formulas", "shared formulas", "array formulas", "names", "tables", "pivot tables", "charts", "pictures",
    "notes", "threaded comments", "dropdowns", "highlight rules", "filter", "merged cells", "frozen panes",
    "links", "links to other workbooks", "hidden sheets", "several sheets", "extension records",
    "sheet names beyond ASCII", "1904 dates", "page setup", "protected sheet", "grouped rows or columns",
    "rich text", "embedded objects", "macros",
]


def describe(path):
    found = dict.fromkeys(FEATURES, 0)
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        read = lambda name: package.read(name).decode("utf-8", "replace")
        workbook = read("xl/workbook.xml")
        sheets = re.findall(r"<(?:\w+:)?sheet\b[^>]*>", workbook)
        found["several sheets"] = int(len(sheets) > 1)
        found["hidden sheets"] = sum('state="hidden"' in s or 'state="veryHidden"' in s for s in sheets)
        found["sheet names beyond ASCII"] = sum(bool(re.search(r'name="[^"]*[^\x00-\x7f]', s)) for s in sheets)
        found["names"] = len(re.findall(r"<(?:\w+:)?definedName\b(?![^>]*_xlnm)", workbook))
        found["1904 dates"] = int(bool(re.search(r'date1904="(1|true)"', workbook)))
        found["tables"] = sum(bool(re.match(r"xl/tables/[^/]+\.xml$", n)) for n in names)
        found["pivot tables"] = sum(bool(re.match(r"xl/pivotTables/[^/]+\.xml$", n)) for n in names)
        found["charts"] = sum(bool(re.match(r"xl/charts/chart[^/]*\.xml$", n)) for n in names)
        found["pictures"] = sum(n.startswith("xl/media/") for n in names)
        found["notes"] = sum(bool(re.match(r"xl/comments[^/]*\.xml$|xl/comments/", n)) for n in names)
        found["threaded comments"] = sum(n.startswith("xl/threadedComments/") for n in names)
        found["links to other workbooks"] = sum(bool(re.match(r"xl/externalLinks/[^/]+\.xml$", n)) for n in names)
        found["embedded objects"] = sum(n.startswith("xl/embeddings/") for n in names)
        found["macros"] = int(any(n.endswith("vbaProject.bin") for n in names))
        if "xl/sharedStrings.xml" in names and package.getinfo("xl/sharedStrings.xml").file_size < 30_000_000:
            found["rich text"] = len(re.findall(r"<(?:\w+:)?r>", read("xl/sharedStrings.xml")))
        cells = 0
        for name in names:
            if not re.match(r"xl/worksheets/[^/]+\.xml$", name) or package.getinfo(name).file_size > 60_000_000:
                continue
            text = read(name)
            cells += len(re.findall(r"<(?:\w+:)?c\b", text))
            found["formulas"] += len(re.findall(r"<(?:\w+:)?f\b", text))
            found["shared formulas"] += len(re.findall(r'<(?:\w+:)?f\b[^>]*t="shared"', text))
            found["array formulas"] += len(re.findall(r'<(?:\w+:)?f\b[^>]*t="array"', text))
            found["dropdowns"] += len(re.findall(r"<(?:\w+:)?dataValidation\b", text))
            found["highlight rules"] += len(re.findall(r"<(?:\w+:)?cfRule\b", text))
            found["filter"] += len(re.findall(r"<(?:\w+:)?autoFilter\b", text))
            found["merged cells"] += len(re.findall(r"<(?:\w+:)?mergeCell\b", text))
            found["frozen panes"] += len(re.findall(r'<(?:\w+:)?pane\b[^>]*state="frozen', text))
            found["links"] += len(re.findall(r"<(?:\w+:)?hyperlink\b", text))
            found["extension records"] += len(re.findall(r"<(?:\w+:)?ext\b", text))
            found["page setup"] += len(re.findall(r"<(?:\w+:)?(?:pageSetup|printOptions|headerFooter)\b", text))
            found["protected sheet"] += len(re.findall(r"<(?:\w+:)?sheetProtection\b", text))
            found["grouped rows or columns"] += len(re.findall(r'outlineLevel="[1-9]', text))
        app = read("docProps/app.xml") if "docProps/app.xml" in names else ""
        producer = re.search(r"<Application>([^<]*)<", app)
    return {
        "file": os.path.basename(path),
        "bytes": os.path.getsize(path),
        "sheets": len(sheets),
        "cells": cells,
        "producer": producer.group(1) if producer else "unknown",
        "features": {k: v for k, v in found.items() if v},
    }


rows = []
for name in sorted(os.listdir(folder)):
    if name.lower().endswith(".xlsx") and not name.startswith("~$"):
        try:
            rows.append(describe(os.path.join(folder, name)))
        except Exception as error:  # a file this script cannot read is itself worth knowing about
            rows.append({"file": name, "error": str(error)[:120], "features": {}})
print(f"{len(rows)} workbooks, {sum(r.get('cells', 0) for r in rows):,} cells, "
      f"{sum(r['features'].get('formulas', 0) for r in rows):,} formulas")
print("\nWorkbooks that have each feature:")
for feature in FEATURES:
    count = sum(feature in r["features"] for r in rows)
    print(f"  {count:4d}  {feature}")
producers = {}
for r in rows:
    key = re.sub(r"[/ ]?\d[\d.$_A-Za-z ]*.*$", "", r.get("producer", "unknown")) or r.get("producer", "unknown")
    producers[key] = producers.get(key, 0) + 1
print("\nSaved by:", ", ".join(f"{k} {v}" for k, v in sorted(producers.items(), key=lambda kv: -kv[1])))
if out:
    json.dump(rows, open(out, "w"), indent=1)
