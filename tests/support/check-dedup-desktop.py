"""Independent selected-record compaction oracle, with fixed retained row IDs."""
import copy, json, re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q = lambda name: '{' + NS + '}' + name

def package(file):
    with zipfile.ZipFile(file) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)), 'Duplicate ZIP member'
        return {name: archive.read(name) for name in names}

def exact(node):
    return (node.tag, tuple(sorted(node.attrib.items())), node.text, node.tail,
            tuple(exact(child) for child in node))

def coordinate(reference):
    match = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', reference)
    assert match, 'Invalid cell address'
    col = 0
    for char in match[1]:
        col = col * 26 + ord(char) - 64
    return int(match[2]), col

def cells(rows):
    result = {}
    for row, node in rows.items():
        for cell in node:
            assert cell.tag == q('c'), 'Extended row child'
            reference = cell.attrib['r']
            assert coordinate(reference)[0] == row, 'Cell/container mismatch'
            assert reference not in result, 'Duplicate cell'
            result[reference] = cell
    return result

old, new = map(package, sys.argv[1:3])
want = json.loads(Path(sys.argv[3]).read_text())
assert want['range'] == 'A1:E14' and want['columns'] == ['A', 'B'] and want['hasHeaders'] is True
assert want['keptRows'] == [2, 4, 5, 6, 7, 8, 9, 10, 12, 13]
assert want['removedRows'] == [{'row': 3, 'keptRow': 2}, {'row': 11, 'keptRow': 2}, {'row': 14, 'keptRow': 12}]
sheet = 'xl/worksheets/sheet1.xml'
assert old.keys() == new.keys(), 'ZIP inventory changed'
for name in old:
    if name != sheet:
        assert old[name] == new[name], 'Unrelated part changed: ' + name
a, b = map(ET.fromstring, (old[sheet], new[sheet]))
ad, bd = a.find(q('sheetData')), b.find(q('sheetData'))
assert ad is not None and bd is not None
ar = {int(row.attrib['r']): row for row in ad}
br = {int(row.attrib['r']): row for row in bd}
assert len(ar) == len(ad) and len(br) == len(bd), 'Duplicate row'
assert ar.keys() == br.keys(), 'Outside F-column requires original row containers'
for row in ar:
    assert ar[row].attrib == br[row].attrib, 'Worksheet row metadata moved/changed: ' + str(row)
original, actual = cells(ar), cells(br)
expected = {reference: copy.deepcopy(cell) for reference, cell in original.items()
            if not (2 <= coordinate(reference)[0] <= 14 and coordinate(reference)[1] <= 5)}
receipts = []
for index, source_row in enumerate(want['keptRows']):
    destination = index + 2
    for col in 'ABCDE':
        old_ref, new_ref = col + str(source_row), col + str(destination)
        if old_ref not in original:
            continue
        cell = copy.deepcopy(original[old_ref])
        cell.attrib['r'] = new_ref
        if col == 'E':
            formula = cell.find(q('f'))
            value = cell.find(q('v'))
            assert formula is not None and value is not None
            assert formula.text == f'C{source_row}*10+$F$20' and formula.attrib == {}
            formula.text = f'C{destination}*10+$F$20'
            value.text = want['formulaRawValues'][index]
        expected[new_ref] = cell
    receipts.append({'sourceRow': source_row, 'destinationRow': destination})
assert actual.keys() == expected.keys(), 'Selected trailing cells or outside cells changed'
for reference in expected:
    assert exact(actual[reference]) == exact(expected[reference]), 'Raw/type/style/formula/cache changed: ' + reference
assert actual['A10'].find(q('v')).text == '9007199254740993'
assert actual['A11'].find(q('v')).text == '9007199254740994'
assert actual['A4'].attrib['t'] == 'b' and actual['A4'].find(q('v')).text == '1'
assert 'A5' not in actual and actual['A6'].attrib['t'] == 'inlineStr'
assert actual['A7'].attrib['t'] == 'e' and actual['A8'].attrib['t'] == 'inlineStr'
a.remove(ad)
b.remove(bd)
assert exact(a) == exact(b), 'Worksheet views/defaults/dimension/metadata changed'
print(json.dumps({'passed': True, 'retained': receipts, 'removed': want['removedRows'],
                  'exactNumericSource': True, 'typedBoundaries': True,
                  'outsideCellsAndRowsExact': True, 'unrelatedPartsExact': len(old) - 1}))
