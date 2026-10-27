"""Independent saved Freeze/cache/package checks; standard-library OOXML reader."""
import hashlib, json, sys, zipfile
import xml.etree.ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def canonical(node):
    return (node.tag, sorted(node.attrib.items()), node.text or '', node.tail or '',
            [canonical(child) for child in node])

def without_pane(xml):
    root = ET.fromstring(xml)
    views = root.findall('s:sheetViews/s:sheetView', NS)
    assert len(views) == 1 and views[0].get('workbookViewId') == '0'
    assert len(views[0].findall('s:pane', NS)) == 1
    for child in list(views[0]):
        if child.tag == '{' + NS['s'] + '}pane': views[0].remove(child)
    return canonical(root)

def inspect(filename, rows, columns):
    with zipfile.ZipFile(filename) as package:
        names = package.namelist()
        if len(set(names)) != len(names) or package.testzip():
            raise AssertionError('Duplicate or corrupt saved package')
        members = {name: package.read(name) for name in names}
    root = ET.fromstring(members['xl/worksheets/sheet1.xml'])
    views = root.findall('s:sheetViews/s:sheetView', NS)
    assert len(views) == 1 and len(views[0].findall('s:pane', NS)) == 1
    pane = views[0].find('s:pane', NS)
    assert pane.get('state') == 'frozen'
    assert int(pane.get('ySplit', '0')) == rows
    assert int(pane.get('xSplit', '0')) == columns
    assert pane.attrib == {'xSplit': str(columns), 'ySplit': str(rows),
                           'topLeftCell': 'B2' if rows == 1 else 'C3',
                           'activePane': 'bottomRight', 'state': 'frozen'}
    cells = {cell.get('r'): cell for cell in root.findall('s:sheetData/s:row/s:c', NS)}
    assert cells['C1'].find('s:f', NS).text == 'A1+B1'
    assert cells['C1'].get('t', 'n') == 'n' and float(cells['C1'].find('s:v', NS).text) == 30
    assert cells['C8'].get('t', 'n') == 'n' and float(cells['C8'].find('s:v', NS).text) == 17
    assert root.find("s:sheetData/s:row[@r='4']", NS).get('hidden') == '1'
    assert root.find("s:cols/s:col[@min='4']", NS).get('hidden') == '1'
    assert root.find('s:mergeCells/s:mergeCell', NS).get('ref') == 'E2:F3'
    return members, {'file': filename, 'sha256': hashlib.sha256(open(filename, 'rb').read()).hexdigest(),
                     'freeze': {'rows': rows, 'columns': columns}, 'typedNumericCache': 30, 'C8': 17}

before, baseline = inspect(sys.argv[1], 1, 1)
records = []
for filename, rows, columns in [(sys.argv[2], 2, 2), (sys.argv[3], 1, 1)]:
    after, record = inspect(filename, rows, columns)
    assert set(after) == set(before), 'Freeze must not add/remove package parts'
    changed = [name for name in before if before[name] != after[name]]
    assert all(name == 'xl/worksheets/sheet1.xml' for name in changed), changed
    assert without_pane(before['xl/worksheets/sheet1.xml']) == without_pane(after['xl/worksheets/sheet1.xml']), 'Unrelated worksheet values/metadata changed'
    record['changedParts'] = changed
    record['unrelatedPartsByteExact'] = True
    records.append(record)
print(json.dumps({'passed': True, 'scope': 'Independent package/XML checks, not actual Excel runtime',
                  'original': baseline, 'outputs': records}))
