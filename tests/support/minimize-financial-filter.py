"""Deterministic real-source reduction; no invented business values/formulas."""
import hashlib
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as E
import zipfile

source, fixture, manifest = map(pathlib.Path, sys.argv[1:])
SOURCE_HASH = '39ce613ef2b6ebab41c4879786aafdefbd401469bd1f5a4c4e2d4f7c329bb11f'
digest = lambda data: hashlib.sha256(data).hexdigest()
source_bytes = source.read_bytes()
assert digest(source_bytes) == SOURCE_HASH, 'Authorized financial source identity changed'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
with zipfile.ZipFile(source) as package:
    assert len(package.namelist()) == len(set(package.namelist()))
    before = {name: package.read(name) for name in package.namelist()}
worksheet = before['xl/worksheets/sheet10.xml']
def fragment(pattern, data):
    found = list(re.finditer(pattern, data, re.S))
    assert len(found) == 1
    item = found[0]
    return item.group(0), {'start': item.start(), 'end': item.end(), 'sha256': digest(item.group(0))}
opening, opening_span = fragment(rb'<worksheet\b[^>]*>', worksheet)
cell, cell_span = fragment(rb'<c\b[^>]*r="D4"[^>]*>.*?</c>', worksheet)
row, row_span = fragment(rb'<row\b[^>]*r="4"[^>]*>', worksheet)
filter_xml, filter_span = fragment(rb'<autoFilter\b.*?</autoFilter>', worksheet)
assert E.fromstring(cell).attrib == {'r': 'D4', 's': '127'}
assert E.fromstring(cell).find('v').text == '254063.62'
workbook = before['xl/workbook.xml']
sheet, sheet_span = fragment(rb'<sheet\b[^>]*r:id="rId10"[^>]*/>', workbook)
assert E.fromstring(b'<root xmlns:r="' + R.encode() + b'">' + sheet + b'</root>')[0].attrib['name'] == '项目成本'
relations = E.fromstring(before['xl/_rels/workbook.xml.rels'])
selected = []
for node in relations:
    if node.attrib.get('Id') == 'rId10' or node.attrib.get('Type') in [R + '/styles', R + '/theme']:
        assert node.attrib.get('TargetMode', 'Internal') == 'Internal'
        selected.append(node)
assert len(selected) == 3
targets = {node.attrib['Type']: node.attrib['Target'] for node in selected}
assert targets[R + '/worksheet'] == 'worksheets/sheet10.xml'
assert targets[R + '/styles'] == 'styles.xml' and targets[R + '/theme'] == 'theme/theme1.xml'
declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
after = {
    'xl/workbook.xml': declaration + b'<workbook xmlns="' + S.encode() + b'" xmlns:r="' + R.encode() + b'"><sheets>' + sheet + b'</sheets></workbook>',
    'xl/worksheets/sheet10.xml': declaration + opening + b'<sheetData>' + row + cell + b'</row></sheetData>' + filter_xml + b'</worksheet>',
    'xl/styles.xml': before['xl/styles.xml'],
    'xl/theme/theme1.xml': before['xl/theme/theme1.xml'],
    '_rels/.rels': declaration + b'<Relationships xmlns="' + P.encode() + b'"><Relationship Id="rId1" Type="' + R.encode() + b'/officeDocument" Target="xl/workbook.xml"/></Relationships>',
}
relationship_xml = ''.join('<Relationship ' + ' '.join(f'{key}="{value}"' for key, value in node.attrib.items()) + '/>' for node in selected)
after['xl/_rels/workbook.xml.rels'] = declaration + f'<Relationships xmlns="{P}">{relationship_xml}</Relationships>'.encode()
types = {'xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
         'xl/worksheets/sheet10.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
         'xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
         'xl/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml'}
after['[Content_Types].xml'] = declaration + (f'<Types xmlns="{C}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(f'<Override PartName="/{part}" ContentType="{kind}"/>' for part, kind in types.items()) + '</Types>').encode()
fixture.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(fixture, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as package:
    for name, data in sorted(after.items()):
        entry = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        entry.compress_type = zipfile.ZIP_DEFLATED
        entry.create_system = 3
        entry.external_attr = 0o100644 << 16
        package.writestr(entry, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
record = {
    'schema': 'workbook.xlsx.real-source-protected-case.v1',
    'id': 'financial-project-cost-opaque-filter',
    'classification': 'intentional-protection-regression; not a formula/value defect',
    'source': {'path': 'supplied financial workbook, not in Git (XLSX_EDITOR_FINANCIAL_WORKBOOK)', 'sha256': SOURCE_HASH, 'bytes': len(source_bytes),
               'provenance': 'Existing authorized financial workbook; no customer pricing workflow or external refresh is claimed.'},
    'fixture': {'path': 'fixtures/' + fixture.name, 'sha256': digest(fixture.read_bytes()), 'bytes': fixture.stat().st_size},
    'operation': {'operation': 'insertRow', 'sheet': '项目成本', 'index': 3, 'count': 1},
    'expected': {'refused': True, 'sourceUnchanged': True, 'publicStateUnchanged': True,
                 'cell': 'D4', 'rawValue': '254063.62', 'styleId': 127},
    'retainedFragments': [
        {'part': 'xl/worksheets/sheet10.xml', 'kind': 'worksheet-opening', **opening_span},
        {'part': 'xl/worksheets/sheet10.xml', 'kind': 'original-row4-opening', **row_span},
        {'part': 'xl/worksheets/sheet10.xml', 'kind': 'original-cost-cell', **cell_span},
        {'part': 'xl/worksheets/sheet10.xml', 'kind': 'original-opaque-autofilter', **filter_span},
        {'part': 'xl/workbook.xml', 'kind': 'original-sheet-reference', **sheet_span}],
    'retainedWholeParts': [{'part': name, 'sha256': digest(before[name])} for name in ['xl/styles.xml', 'xl/theme/theme1.xml']],
    'recipe': {'generator': 'tests/support/minimize-financial-filter.py',
               'retained': 'One original literal cost cell, its row/opening, original WPS autoFilter extension and exact styles/theme; original sheet name/ID/relationship identity.',
               'removed': 'All other cells, formulas, worksheets, links, media, embeddings, properties and unrelated package members.',
               'generated': 'Minimal package root/content-type/workbook relationship scaffolding and sheetData wrapper; no new business values or formulas.',
               'zip': 'Sorted parts, fixed 1980 timestamps, deflate9, fixed regular-file permissions.',
               'counterfactual': 'Remove only the retained autoFilter fragment; ordinary row insertion must then qualify.'},
    'partSha256': {name: digest(data) for name, data in sorted(after.items())},
}
manifest.parent.mkdir(parents=True, exist_ok=True)
manifest.write_text(json.dumps(record, ensure_ascii=False, indent=2) + '\n')
assert digest(source.read_bytes()) == SOURCE_HASH
print(json.dumps({'fixture': str(fixture), 'sha256': record['fixture']['sha256'], 'bytes': fixture.stat().st_size}))
