"""Audit exact retained spill cell deltas, then independent metadata closure.

This is scripted OOXML preservation evidence, not an Excel runtime capture.
"""
import hashlib
import math
import json
import pathlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as E
import zipfile

source, saved, specification, report = map(pathlib.Path, sys.argv[1:])
plan = json.loads(specification.read_text())
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
S = '{' + NS + '}'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
GRAPH = {'xl/metadata.xml', 'xl/richData/rdrichvalue.xml',
         'xl/richData/rdrichvaluestructure.xml', 'xl/richData/rdRichValueTypes.xml',
         'xl/_rels/workbook.xml.rels', '[Content_Types].xml'}

def parts(path):
    with zipfile.ZipFile(path) as package:
        assert len(package.namelist()) == len(set(package.namelist()))
        return {name: package.read(name) for name in package.namelist()}

def parse(data):
    assert not re.search(br'<!DOCTYPE|<!ENTITY', data, re.I)
    return E.fromstring(data, parser=E.XMLParser(target=E.TreeBuilder(insert_comments=True, insert_pis=True)))

def tree(node):
    return (node.tag, sorted(node.attrib.items()), node.text, node.tail, [tree(item) for item in node])

def remove_raw(data, local, attributes):
    # A closed empty relationship/override entry; preserve every other byte,
    # including source namespace declarations and opaque whitespace.
    matches = []
    for match in re.finditer(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' + local.encode() + rb'\b[^<>]*/>', data):
        raw = match.group(0)
        opening = re.match(rb'<[^\s/>]+', raw).group(0)
        fragment = b'<item' + raw[len(opening):]
        node = E.fromstring(fragment)
        if node.attrib == attributes:
            matches.append(match)
    assert len(matches) == 1, 'Exact calculation-chain node cannot be located'
    match = matches[0]
    return data[:match.start()] + data[match.end():]

a, b = parts(source), parts(saved)
normalized = dict(a)
removed = set(a) - set(b)
assert removed <= {'xl/calcChain.xml'}
assert set(b) - set(a) <= GRAPH, 'Unexpected added package member'
normalizations = []
if removed:
    for part, namespace, name, matches in [
        ('xl/_rels/workbook.xml.rels', REL, 'Relationship', lambda node: node.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain' and node.attrib.get('Target') in ['calcChain.xml', '/xl/calcChain.xml'] and node.attrib.get('TargetMode', 'Internal') == 'Internal'),
        ('[Content_Types].xml', CT, 'Override', lambda node: node.attrib == {'PartName': '/xl/calcChain.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'})]:
        nodes = [node for node in parse(a[part]) if node.tag == f'{{{namespace}}}{name}' and matches(node)]
        assert len(nodes) == 1, 'Ambiguous calculation-chain removal'
        normalized[part] = remove_raw(a[part], name, nodes[0].attrib)
        normalizations.append({'part': part, 'removedElement': nodes[0].tag, 'attributes': nodes[0].attrib})
    del normalized['xl/calcChain.xml']
    normalizations.append({'removedPart': 'xl/calcChain.xml'})

old, new = parse(a['xl/workbook.xml']), parse(b['xl/workbook.xml'])
before_calc, after_calc = old.find(S + 'calcPr'), new.find(S + 'calcPr')
assert after_calc is not None
assert len(after_calc) == 0 and not (after_calc.text or '').strip()
if before_calc is not None:
    assert len(before_calc) == 0 and before_calc.text == after_calc.text and before_calc.tail == after_calc.tail
declared = {'fullCalcOnLoad': '1', 'forceFullCalc': '1', 'calcMode': 'auto', 'calcOnSave': '1'}
for key in set(before_calc.attrib if before_calc is not None else {}) | set(after_calc.attrib):
    prior = before_calc.attrib.get(key) if before_calc is not None else None
    if prior != after_calc.attrib.get(key):
        assert key in declared and after_calc.attrib[key] == declared[key]
        normalizations.append({'part': 'xl/workbook.xml', 'calcPr': key, 'before': prior, 'after': after_calc.attrib[key]})
if before_calc is not None:
    old.remove(before_calc)
new.remove(after_calc)
assert tree(old) == tree(new), 'Other workbook metadata changed'
normalized['xl/workbook.xml'] = b['xl/workbook.xml']

targets = plan['cells']
for part, expected in targets.items():
    old, new = parse(a[part]), parse(b[part])
    assert old.tag == new.tag == S + 'worksheet'
    old_data, new_data = old.find(S + 'sheetData'), new.find(S + 'sheetData')
    assert old_data is not None and new_data is not None
    old_cells = {cell.attrib['r']: cell for row in old_data for cell in row if cell.tag == S + 'c'}
    new_cells = {cell.attrib['r']: cell for row in new_data for cell in row if cell.tag == S + 'c'}
    assert len(new_cells) == sum(cell.tag == S + 'c' for row in new_data for cell in row), 'Duplicate worksheet cell'
    ordinary = plan.get('literals', {}).get(part, {})
    uncached = plan.get('uncalculated', {}).get(part, [])
    consumers = plan.get('consumers', {}).get(part, {})
    formula_results = plan.get('formulaResults', {}).get(part, {})
    categories = [set(expected), set(ordinary), set(uncached), set(consumers), set(formula_results)]
    assert all(not a & b for index, a in enumerate(categories) for b in categories[index + 1:])
    for ref, value in expected.items():
        cell = new_cells[ref]
        original = old_cells.get(ref)
        original_attrs = {k: v for k, v in (original.attrib if original is not None else {'r': ref}).items() if k not in ['t', 'cm', 'vm']}
        assert {k: v for k, v in cell.attrib.items() if k not in ['t', 'cm', 'vm']} == original_attrs, 'Source cell attributes/style changed'
        assert cell.find(S + 'v') is not None and len(cell.find(S + 'v')) == 0
        assert cell.find(S + 'v').text == '#VALUE!' if value['value'] == '#SPILL!' else float(cell.find(S + 'v').text) == value['value']
        mapping = plan['masters'].get(part + '\0' + ref)
        if mapping is not None:
            assert cell.attrib.get('cm') == str(mapping['cm'])
            assert cell.attrib.get('vm') == (str(mapping['vm']) if 'vm' in mapping else None)
            formula = cell.find(S + 'f')
            assert formula is not None and formula.text == value['formula'] and len(formula) == 0
            attributes = {'t': 'array', 'ref': mapping['range']}
            for flag in ['aca', 'ca']:
                if flag in mapping:
                    attributes[flag] = mapping[flag]
            assert formula.attrib == attributes
            assert [child.tag for child in cell] == [S + 'f', S + 'v']
            assert cell.attrib.get('t', 'n') == ('e' if 'vm' in mapping else 'n')
        else:
            assert cell.find(S + 'f') is None and cell.attrib.get('t', 'n') == 'n'
            assert 'cm' not in cell.attrib and 'vm' not in cell.attrib
            assert [child.tag for child in cell] == [S + 'v']
    for ref, value in ordinary.items():
        cell, original = new_cells[ref], old_cells.get(ref)
        assert {k: v for k, v in cell.attrib.items() if k != 't'} == {k: v for k, v in (original.attrib if original is not None else {'r': ref}).items() if k != 't'}
        assert cell.attrib.get('t', 'n') == 'n' and [child.tag for child in cell] == [S + 'v']
        assert len(cell.find(S + 'v')) == 0 and float(cell.find(S + 'v').text) == value
    # Scalar consumers have no follower ownership. Imported singleton arrays
    # preserve the exact original f/metadata; ordinary authored aliases store
    # the explicitly expected canonical formula with a fresh numeric cache.
    for ref, value in consumers.items():
        assert set(value) <= {'formula', 'value', 'singleton'} and isinstance(value['formula'], str)
        assert type(value['value']) in [int, float] and math.isfinite(value['value'])
        assert 'singleton' not in value or type(value['singleton']) is bool
        cell, original = new_cells[ref], old_cells.get(ref)
        assert {k: v for k, v in cell.attrib.items() if k != 't'} == {k: v for k, v in (original.attrib if original is not None else {'r': ref}).items() if k != 't'}
        assert cell.attrib.get('t', 'n') == 'n' and [child.tag for child in cell] == [S + 'f', S + 'v']
        formula, cached = cell.find(S + 'f'), cell.find(S + 'v')
        assert len(formula) == len(cached) == 0 and formula.text == value['formula']
        assert float(cached.text) == value['value']
        if value.get('singleton'):
            assert original is not None
            assert tree(formula) == tree(original.find(S + 'f')) and formula.attrib == {'t': 'array', 'ref': ref}
            assert cell.attrib.get('cm') == original.attrib['cm'] and 'vm' not in cell.attrib
        else:
            assert formula.attrib == {} and 'cm' not in cell.attrib and 'vm' not in cell.attrib
    for ref, value in formula_results.items():
        assert set(value) == {'formula', 'value', 'dataType'} and isinstance(value['formula'], str)
        cell, original = new_cells[ref], old_cells.get(ref)
        assert {k: v for k, v in cell.attrib.items() if k != 't'} == {k: v for k, v in (original.attrib if original is not None else {'r': ref}).items() if k != 't'}
        assert 'cm' not in cell.attrib and 'vm' not in cell.attrib and [child.tag for child in cell] == [S + 'f', S + 'v']
        formula, cached = cell.find(S + 'f'), cell.find(S + 'v')
        assert formula.attrib == {} and len(formula) == len(cached) == 0 and formula.text == value['formula']
        if original is not None and original.find(S + 'f') is not None:
            assert original.find(S + 'f').attrib == {}
        if value['dataType'] == 'number':
            assert type(value['value']) in [int, float] and math.isfinite(value['value'])
            assert cell.attrib.get('t', 'n') == 'n' and float(cached.text) == value['value']
        elif value['dataType'] == 'text':
            assert isinstance(value['value'], str) and cell.attrib.get('t') == 'str' and cached.text == value['value']
        elif value['dataType'] == 'boolean':
            assert type(value['value']) is bool and cell.attrib.get('t') == 'b' and cached.text == ('1' if value['value'] else '0')
        else:
            assert False, 'Unqualified formula result type'
    for ref in uncached:
        original, cell = old_cells[ref], new_cells[ref]
        assert original.find(S + 'f') is not None and original.find(S + 'f').attrib == {}, 'Only ordinary unresolved formula caches can be removed'
        assert cell.find(S + 'v') is None and cell.find(S + 'is') is None
        original.attrib.pop('t', None)
        cell.attrib.pop('t', None)
        for value in original.findall(S + 'v'):
            original.remove(value)
        assert tree(original) == tree(cell), 'Authored unresolved formula changed'
    addressed = set(expected) | set(ordinary) | set(uncached) | set(consumers) | set(formula_results)
    old_rows = {row.attrib['r'] for row in old_data}
    for data in [old_data, new_data]:
        for row in list(data):
            for cell in list(row):
                if cell.tag == S + 'c' and cell.attrib['r'] in addressed:
                    row.remove(cell)
            if data is new_data and row.attrib['r'] not in old_rows:
                assert row.attrib == {'r': row.attrib['r']} and len(row) == 0 and not (row.text or '').strip()
                data.remove(row)
    assert tree(old) == tree(new), 'Unaddressed worksheet content changed'
    normalized[part] = b[part]

for part in a:
    if part not in GRAPH | set(targets) | {'xl/workbook.xml', 'xl/calcChain.xml'}:
        assert a[part] == b[part], f'Unrelated package bytes changed: {part}'
with tempfile.TemporaryDirectory(prefix='spill-export-audit-') as directory:
    directory = pathlib.Path(directory)
    clone, mapping = directory / 'normalized.xlsx', directory / 'masters.json'
    with zipfile.ZipFile(clone, 'w', zipfile.ZIP_DEFLATED) as package:
        for part, data in normalized.items():
            package.writestr(part, data)
    mapping.write_text(json.dumps(plan['masters']))
    result = subprocess.run([sys.executable, str(pathlib.Path(__file__).with_name('check-spill-metadata.py')), str(clone), str(saved), str(mapping)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    metadata = json.loads(result.stdout)
record = {'passed': True, 'method': 'exact-retained-spill-cell-and-metadata-delta',
          'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
          'savedSha256': hashlib.sha256(saved.read_bytes()).hexdigest(),
          'normalizations': normalizations, 'qualifiedCells': sum(map(len, targets.values())),
          'scalarConsumers': sum(map(len, plan.get('consumers', {}).values())),
          'formulaResults': sum(map(len, plan.get('formulaResults', {}).values())),
          'metadata': metadata}
report.parent.mkdir(parents=True, exist_ok=True)
report.write_text(json.dumps(record, indent=2) + '\n')
print(json.dumps({'passed': True, 'report': str(report)}))
