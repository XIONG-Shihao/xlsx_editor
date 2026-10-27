"""Independent OOXML/source attribution for the minimized protected case."""
import hashlib
import json
import pathlib
import sys
import xml.etree.ElementTree as E
import zipfile

source, fixture, manifest = map(pathlib.Path, sys.argv[1:])
record = json.loads(manifest.read_text())
hash_bytes = lambda data: hashlib.sha256(data).hexdigest()
def package(path):
    with zipfile.ZipFile(path) as archive:
        assert len(archive.namelist()) == len(set(archive.namelist()))
        return {name: archive.read(name) for name in archive.namelist()}
before, after = package(source), package(fixture)
assert hash_bytes(source.read_bytes()) == record['source']['sha256'] == '39ce613ef2b6ebab41c4879786aafdefbd401469bd1f5a4c4e2d4f7c329bb11f'
assert hash_bytes(fixture.read_bytes()) == record['fixture']['sha256']
assert set(after) == set(record['partSha256']) == {'[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet10.xml', 'xl/styles.xml', 'xl/theme/theme1.xml'}
for part, digest in record['partSha256'].items():
    assert hash_bytes(after[part]) == digest
for fragment in record['retainedFragments']:
    value = before[fragment['part']][fragment['start']:fragment['end']]
    assert hash_bytes(value) == fragment['sha256']
    assert after[fragment['part']].count(value) == 1, 'Retained source fragment is not exact and unique'
for part in record['retainedWholeParts']:
    assert after[part['part']] == before[part['part']] and hash_bytes(after[part['part']]) == part['sha256']
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P = '{http://schemas.openxmlformats.org/package/2006/relationships}'
C = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sheet = E.fromstring(after['xl/worksheets/sheet10.xml'])
assert sheet.tag == S + 'worksheet' and [node.tag for node in sheet] == [S + 'sheetData', S + 'autoFilter']
rows = list(sheet.find(S + 'sheetData'))
assert len(rows) == 1 and rows[0].attrib['r'] == '4' and len(rows[0]) == 1
cell = rows[0][0]
assert cell.tag == S + 'c' and cell.attrib == {'r': 'D4', 's': '127'}
assert [node.tag for node in cell] == [S + 'v'] and cell[0].text == '254063.62'
filter_node = sheet.find(S + 'autoFilter')
assert filter_node.attrib == {'ref': 'A1:H63', '{http://www.wps.cn/officeDocument/2017/etCustomData}filterBottomFollowUsedRange': '0'}
assert len(filter_node) == 1 and filter_node[0].tag == S + 'extLst' and len(filter_node[0]) == 0
workbook = E.fromstring(after['xl/workbook.xml'])
assert workbook.tag == S + 'workbook' and [node.tag for node in workbook] == [S + 'sheets']
sheets = list(workbook[0])
assert len(sheets) == 1 and sheets[0].attrib == {'name': '项目成本', 'sheetId': '38', R + 'id': 'rId10'}
relations = E.fromstring(after['xl/_rels/workbook.xml.rels'])
assert relations.tag == P + 'Relationships' and len(relations) == 3
assert len({node.attrib['Id'] for node in relations}) == 3
targets = {'worksheets/sheet10.xml': 'worksheet', 'styles.xml': 'styles', 'theme/theme1.xml': 'theme'}
for node in relations:
    assert node.tag == P + 'Relationship' and node.attrib['Target'] in targets
    assert node.attrib['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + targets[node.attrib['Target']]
    assert node.attrib.get('TargetMode', 'Internal') == 'Internal'
    assert 'xl/' + node.attrib['Target'] in after
root_rels = E.fromstring(after['_rels/.rels'])
assert root_rels.tag == P + 'Relationships' and len(root_rels) == 1 and root_rels[0].attrib == {'Id': 'rId1', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target': 'xl/workbook.xml'}
types = E.fromstring(after['[Content_Types].xml'])
assert types.tag == C + 'Types'
overrides = {node.attrib['PartName']: node.attrib['ContentType'] for node in types if node.tag == C + 'Override'}
assert len(types) == 6 and len(overrides) == 4 and overrides == {
    '/xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    '/xl/worksheets/sheet10.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    '/xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    '/xl/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml'}
assert {node.attrib['Extension']: node.attrib['ContentType'] for node in types if node.tag == C + 'Default'} == {'rels': 'application/vnd.openxmlformats-package.relationships+xml', 'xml': 'application/xml'}
assert all(node.tag in [C + 'Override', C + 'Default'] for node in types)
styles = E.fromstring(after['xl/styles.xml'])
formats = styles.find(S + 'cellXfs')
assert int(formats.attrib['count']) == len(formats) and len(formats) > 127
assert not sheet.findall('.//' + S + 'f')
print(json.dumps({'passed': True, 'sourceSha256': record['source']['sha256'], 'fixtureSha256': record['fixture']['sha256'], 'retainedCells': 1, 'retainedFormulas': 0, 'attributedFragments': len(record['retainedFragments']), 'guardFeature': 'Original WPS autoFilter attribute and extLst', 'externalRuntimeClaim': False}))
