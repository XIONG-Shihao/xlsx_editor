"""Independent declared typed caches/policy/edits and exact unrelated OPC/XML preservation."""
import sys, json, zipfile, hashlib, struct, copy, math, re
import xml.etree.ElementTree as E

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = '{' + NS + '}'
CELL = re.compile(r'^[A-Z]{1,3}[1-9][0-9]{0,6}$')
ERRORS = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'}
def sha(b): return hashlib.sha256(b).hexdigest()
def bits(x):
    assert isinstance(x, (int, float)) and not isinstance(x, bool), 'Expected finite number'
    assert math.isfinite(float(x)), 'Nonfinite numeric cache'
    return struct.pack('>d', float(x)).hex()
def parts(file):
    with zipfile.ZipFile(file) as z:
        names = z.namelist()
        assert len(names) == len(set(names)), 'Duplicate ZIP member'
        assert all(not n.startswith('/') and '\\' not in n and '..' not in n.split('/') for n in names), 'Unsafe ZIP path'
        return {n: z.read(n) for n in names}
def parse(b):
    assert b'<!DOCTYPE' not in b and b'<!ENTITY' not in b, 'Unsupported XML entity/doctype'
    return E.fromstring(b)
def canonical(node):
    return node.tag, tuple(sorted(node.attrib.items())), node.text or '', node.tail or '', tuple(canonical(c) for c in node)
def leaf(node):
    assert not node.attrib and not list(node), 'Unexpected cache leaf attributes/children'
def cells(root):
    assert root.tag == N+'worksheet', 'Wrong worksheet namespace'
    data = root.findall(N+'sheetData')
    assert len(data) == 1, 'Exactly one sheetData required'
    found = {}; rows = set()
    for row in data[0]:
        assert row.tag == N+'row', 'Unexpected sheetData child'
        number = row.get('r')
        assert number and re.fullmatch(r'[1-9][0-9]*', number) and number not in rows, 'Duplicate/noncanonical row'
        rows.add(number)
        for cell in row:
            assert cell.tag == N+'c', 'Unexpected row child'
            ref = cell.get('r')
            assert ref and CELL.fullmatch(ref) and ref not in found, 'Duplicate/noncanonical cell'
            assert re.search(r'[0-9]+$', ref)[0] == number, 'Cell/row disagreement'
            found[ref] = cell
    return found
def dimension_reference(value):
    assert isinstance(value, str), 'Dimension expectation must be a string'
    refs = value.split(':'); assert len(refs) in (1, 2), 'One dimension rectangle required'
    coordinates = []
    for ref in refs:
        assert CELL.fullmatch(ref), 'Noncanonical dimension expectation'
        match = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', ref)
        column = 0
        for letter in match[1]: column = column * 26 + ord(letter) - 64
        row = int(match[2]); assert 1 <= column <= 16384 and 1 <= row <= 1048576
        coordinates.append((row, column))
    first, last = coordinates[0], coordinates[-1]
    assert first[0] <= last[0] and first[1] <= last[1], 'Reversed dimension expectation'
    return value
def one_formula(cell, wanted):
    formulas = cell.findall(N+'f')
    assert len(formulas) == 1 and formulas[0].text == wanted, 'Formula identity mismatch'
    assert not list(formulas[0]), 'Formula is not a leaf'
    return formulas[0]
def typed_cache(cell, wanted):
    assert set(wanted) <= {'type', 'value'} and 'type' in wanted, 'Unknown cache expectation field'
    kind = wanted['type']; values = cell.findall(N+'v')
    assert not cell.findall(N+'is'), 'Formula cache has inline-string content'
    if kind == 'absent':
        assert 'value' not in wanted and not values, 'Changed/unresolved formula retained a cache'
        assert cell.get('t') in (None, 'n'), 'Absent cache retained a borrowed typed-cache label'
        return {'type': 'absent'}
    assert len(values) == 1 and 'value' in wanted, 'Exactly one declared cache required'
    leaf(values[0]); raw = values[0].text or ''
    if kind == 'number':
        assert cell.get('t') in (None, 'n'), 'Numeric cache relabeled'
        assert raw.strip() == raw and raw, 'Invalid numeric raw value'
        actual = float(raw); assert math.isfinite(actual)
        assert bits(actual) == bits(wanted['value']), 'Numeric cache bits differ'
        return {'type':kind, 'raw':raw, 'bits':bits(actual)}
    if kind == 'boolean':
        assert isinstance(wanted['value'], bool) and cell.get('t') == 'b'
        assert raw == ('1' if wanted['value'] else '0'), 'Boolean cache differs'
    elif kind == 'text':
        assert isinstance(wanted['value'], str) and cell.get('t') == 'str'
        assert raw == wanted['value'], 'Text cache differs'
    elif kind == 'error':
        assert wanted['value'] in ERRORS and cell.get('t') == 'e'
        assert raw == wanted['value'], 'Error cache differs'
    else: raise AssertionError('Unknown cache expectation type')
    return {'type':kind, 'raw':raw}
def typed_literal(cell, wanted):
    assert set(wanted) == {'type','value'} and not cell.findall(N+'f'), 'Wrong literal declaration/formula'
    kind, value = wanted['type'], wanted['value']
    if kind == 'text':
        assert isinstance(value, str) and cell.get('t') == 'inlineStr' and not cell.findall(N+'v')
        strings = cell.findall(N+'is'); assert len(strings) == 1 and not strings[0].attrib
        texts = strings[0].findall(N+'t'); assert len(texts) == len(strings[0]) == 1 and not list(texts[0])
        assert set(texts[0].attrib) <= {'{http://www.w3.org/XML/1998/namespace}space'}
        assert (texts[0].text or '') == value, 'Literal text differs'
        return {'type':'text','raw':value}
    return typed_cache(cell, wanted)
def mask(cell, formula=False, literal=False):
    cell.attrib.pop('t', None)
    for child in list(cell):
        if child.tag == N+'v' or (literal and child.tag in (N+'f',N+'is')): cell.remove(child)
    if formula:
        found = cell.findall(N+'f')
        assert len(found) == 1
        found[0].text = '__DECLARED_FORMULA__'

def inspect(source, saved, expectation):
    assert expectation.get('schema') == 'workbook.xlsx.calculation-lifecycle-expected.v1'
    before, after = parts(source), parts(saved)
    assert set(before) == set(after), 'Undeclared package graph/member change'
    sheets = expectation['sheetParts']; assert len(set(sheets.values())) == len(sheets)
    dimensions = expectation.get('dimensions')
    if 'dimensions' in expectation:
        assert isinstance(dimensions, dict) and set(dimensions) == set(sheets), 'Closed per-sheet dimension expectations required'
        for target in dimensions.values():
            assert isinstance(target, dict) and set(target) == {'source', 'saved'}, 'Unknown dimension expectation field'
            dimension_reference(target['source']); dimension_reference(target['saved'])
    workbook = 'xl/workbook.xml'; oldwb, newwb = parse(before[workbook]), parse(after[workbook])
    assert oldwb.tag == newwb.tag == N+'workbook'
    oldcalc, newcalc = oldwb.findall(N+'calcPr'), newwb.findall(N+'calcPr')
    assert len(oldcalc) == len(newcalc) == 1, 'Exactly one calcPr required'
    assert not list(oldcalc[0]) and not list(newcalc[0]), 'calcPr is a leaf'
    policy = expectation['calcPr']; assert set(policy) <= {'calcMode','calcOnSave','fullCalcOnLoad','forceFullCalc'}
    for key, value in policy.items():
        assert newcalc[0].get(key) == value, 'Wrong saved '+key
    for key in policy:
        oldcalc[0].attrib.pop(key,None); newcalc[0].attrib.pop(key,None)
    assert canonical(oldwb) == canonical(newwb), 'Undeclared workbook/calculation metadata change'
    declared = expectation.get('cells',{}); literals = expectation.get('literals',{})
    assert not (set(declared) & set(literals)), 'Overlapping formula/literal declarations'
    seen = set(); readback = []; dimension_readback = []
    for sheet, part in sheets.items():
        old, new = parse(before[part]), parse(after[part]); oc, nc = cells(old), cells(new)
        assert set(oc) == set(nc), 'Undeclared cell membership change'
        for ref in oc:
            key = sheet+'!'+ref; a,b = oc[ref],nc[ref]
            if key in declared:
                target=declared[key]; assert set(target) == {'formula','cache'}
                assert isinstance(target['formula'],str) and target['formula']
                one_formula(a, a.find(N+'f').text if a.find(N+'f') is not None else None)
                f = one_formula(b,target['formula'])
                assert a.find(N+'f').attrib == f.attrib, 'Formula metadata changed'
                receipt = typed_cache(b,target['cache'])
                assert set(child.tag for child in b) <= {N+'f',N+'v'}, 'Unexpected formula-cache graph'
                readback.append({'cell':key,'formula':target['formula'],**receipt}); seen.add(key)
                mask(a,formula=True);mask(b,formula=True)
            elif key in literals:
                readback.append({'cell':key,**typed_literal(b,literals[key])});seen.add(key)
                mask(a,literal=True);mask(b,literal=True)
        if dimensions is not None:
            before_dimension, after_dimension = old.findall(N+'dimension'), new.findall(N+'dimension')
            assert len(before_dimension) == len(after_dimension) == 1, 'Exactly one declared dimension required'
            a, b = before_dimension[0], after_dimension[0]; target = dimensions[sheet]
            assert a.attrib == {'ref': target['source']} and b.attrib == {'ref': target['saved']}, 'Wrong source/saved dimension: '+sheet
            assert not list(a) and not list(b) and not (a.text or '') and not (b.text or ''), 'Dimension must remain a leaf'
            dimension_readback.append({'sheet':sheet,**target})
            # Only a proven exact ref change is masked. Membership, every other
            # attribute/node, row/style/formula fields and all OPC parts stay exact.
            a.set('ref','__DECLARED_DIMENSION__'); b.set('ref','__DECLARED_DIMENSION__')
        assert canonical(old) == canonical(new), 'Unrelated worksheet/formula/style fields changed: '+sheet
    assert seen == set(declared)|set(literals), 'Declared cache/literal missing'
    allowed = {workbook}|set(sheets.values())
    for part in before:
        if part not in allowed: assert before[part] == after[part], 'Unrelated ZIP part changed: '+part
    return {'schema':'workbook.xlsx.calculation-lifecycle-readback.v1','passed':True,'sourceSha256':sha(open(source,'rb').read()),'savedSha256':sha(open(saved,'rb').read()),'calcPr':policy,'targets':readback,**({'dimensions':dimension_readback} if dimensions is not None else {}),'changedParts':[p for p in before if before[p]!=after[p]],'unrelatedPackageAndCellMetadataExact':True,'scope':'Declared typed saved caches, formula edits and calcPr, plus exact source/saved dimension refs only when explicitly declared; pending/current authority is checked through accepted owner receipts, never inferred from saved XML'}

if __name__ == '__main__':
    source,saved,expected=sys.argv[1:]
    print(json.dumps(inspect(source,saved,json.load(open(expected))),ensure_ascii=False))
