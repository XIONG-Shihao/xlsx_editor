"""Independent stdlib ZIP/XML inspection of exported reference-corpus values."""
import hashlib
import json
import posixpath
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def sheet_path(archive, name):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {item.attrib["Id"]: item.attrib["Target"] for item in relations}
    sheet = next(item for item in workbook.findall("s:sheets/s:sheet", NS) if item.attrib["name"] == name)
    target = targets[sheet.attrib[RID]]
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)


def same_value(actual, expected):
    if isinstance(expected, bool):
        return isinstance(actual, bool) and actual == expected
    if isinstance(expected, (int, float)):
        return not isinstance(actual, bool) and isinstance(actual, (int, float)) and struct.pack(">d", float(actual)) == struct.pack(">d", float(expected))
    return type(actual) is type(expected) and actual == expected


request = json.load(sys.stdin)
failures = []
with zipfile.ZipFile(request["source"]) as source, zipfile.ZipFile(request["saved"]) as saved:
    sheet = sheet_path(saved, request["sheet"])
    cells = {item.attrib["r"]: item for item in ET.fromstring(saved.read(sheet)).findall("s:sheetData/s:row/s:c", NS)}
    marker = cells.get(request["marker"]["ref"])
    marker_text = None
    if marker is not None:
        if marker.attrib.get("t") == "inlineStr":
            marker_text = "".join(item.text or "" for item in marker.findall("s:is//s:t", NS))
        elif marker.attrib.get("t") == "s":
            strings = ET.fromstring(saved.read("xl/sharedStrings.xml")).findall("s:si", NS)
            marker_text = "".join(item.text or "" for item in strings[int(marker.find("s:v", NS).text)].findall(".//s:t", NS))
        elif marker.attrib.get("t") == "str":
            marker_text = marker.find("s:v", NS).text or ""
    if marker_text != request["marker"]["value"]:
        failures.append(request["marker"]["ref"] + ": accepted edit was not exported")
    for case in request["cases"]:
        cell = cells.get(case["ref"])
        if cell is None:
            failures.append(case["ref"] + ": missing cell")
            continue
        formula = cell.find("s:f", NS)
        if formula is None or (formula.text or "") != case["formula"]:
            failures.append(case["ref"] + ": authored formula changed")
        value = cell.find("s:v", NS)
        raw = None if value is None else value.text or ""
        kind = cell.attrib.get("t", "n")
        if case["status"] == "error":
            if kind != "e" or raw != case["error"]:
                failures.append(case["ref"] + ": wrong typed error cache")
        else:
            try:
                actual = raw if kind == "str" else raw == "1" if kind == "b" else float(raw) if kind == "n" else None
                if not same_value(actual, case["value"]):
                    failures.append(case["ref"] + ": wrong exact typed cache")
            except (TypeError, ValueError):
                failures.append(case["ref"] + ": missing or malformed cache")
    # Only the edited worksheet and calculation bookkeeping may change.
    allowed = {sheet, "xl/workbook.xml", "xl/calcChain.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
    preserved = []
    for member in source.namelist():
        if member in allowed:
            continue
        if member not in saved.namelist() or source.read(member) != saved.read(member):
            failures.append(member + ": unrelated package part changed")
        else:
            preserved.append(member)
    extras = set(saved.namelist()) - set(source.namelist())
    if extras:
        failures.append("unexpected package members: " + str(sorted(extras)))

report = {"passed": not failures, "method": "independent-python-stdlib-zip-xml-exact-binary64-and-types",
          "sourceSha256": hashlib.sha256(open(request["source"], "rb").read()).hexdigest(),
          "savedSha256": hashlib.sha256(open(request["saved"], "rb").read()).hexdigest(),
          "cases": len(request["cases"]), "preservedMembers": preserved, "failures": failures}
print(json.dumps(report))
sys.exit(0 if report["passed"] else 1)
