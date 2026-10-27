"""Independent typed cache and exact unrelated OOXML preservation checker."""
import copy
import datetime
import hashlib
import json
import math
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def digest(filename):
    h = hashlib.sha256()
    with open(filename, 'rb') as f:
        for block in iter(lambda: f.read(65536), b''):
            h.update(block)
    return h.hexdigest()


def tree(node):
    return (node.tag, sorted(node.attrib.items()), node.text or '',
            node.tail or '', tuple(tree(child) for child in node))


def serial(case, calendar):
    value = datetime.date(*case['date'])
    if calendar == 1904:
        return (value - datetime.date(1904, 1, 1)).days
    days = (value - datetime.date(1899, 12, 31)).days
    return days + (value >= datetime.date(1900, 3, 1))


def cells(root):
    data = root.findall(NS + 'sheetData')
    if len(data) != 1:
        raise ValueError('One direct sheetData required')
    output = {}
    for row in data[0]:
        if row.tag != NS + 'row':
            raise ValueError('Unexpected sheetData child')
        for cell in row:
            ref = cell.get('r')
            if cell.tag != NS + 'c' or not ref or ref in output:
                raise ValueError('Invalid/duplicate worksheet cell')
            output[ref] = cell
    return output


def without_targets(root, targets):
    root = copy.deepcopy(root)
    data = root.find(NS + 'sheetData')
    for row in list(data):
        for cell in list(row):
            if cell.get('r') in targets:
                row.remove(cell)
            elif (len(cell.findall(NS + 'f')) == 1 and
                  len(cell.findall(NS + 'v')) == 1 and cell.get('t') in (None, 'n')):
                value = cell.find(NS + 'v')
                number = float(value.text)
                if not math.isfinite(number):
                    raise ValueError('Nonfinite existing formula cache')
                # Formula recalculation may spell 9 as 9.0. Preserve exact type,
                # attributes/shape and binary64 bits rather than decimal spelling.
                value.text = struct.pack('>d', number).hex()
        # A new target-only row may be removed; authored extra metadata may not.
        if not list(row) and set(row.attrib) == {'r'} and not (row.text or '').strip():
            data.remove(row)
    return tree(root)


def check(source, saved, expected):
    if expected.get('schema') != 'workbook.xlsx.function-library-expected.v1':
        raise ValueError('Unknown expected profile')
    targets = {case['cell'] for case in expected['cases']}
    if len(targets) != len(expected['cases']) or not targets:
        raise ValueError('Duplicate/empty declared targets')
    worksheet = expected['sheet']
    readback = []
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(saved) as exported:
        for package in (original, exported):
            if len(package.namelist()) != len(set(package.namelist())):
                raise ValueError('Duplicate package member')
        if set(original.namelist()) != set(exported.namelist()):
            raise ValueError('Unqualified package member delta')
        for name in original.namelist():
            if name != worksheet and original.read(name) != exported.read(name):
                if name != 'xl/workbook.xml':
                    raise ValueError('Unrelated package bytes changed: ' + name)
                old_book, new_book = ET.fromstring(original.read(name)), ET.fromstring(exported.read(name))
                old_calc, new_calc = old_book.findall(NS+'calcPr'), new_book.findall(NS+'calcPr')
                if len(old_calc) != 1 or len(new_calc) != 1:
                    raise ValueError('One calculation settings node required')
                # Save/Copy can update scheduling hints. Calendar, mode, names
                # and all other workbook settings retain their exact meaning.
                for node in [old_calc[0],new_calc[0]]:
                    for key in ['calcOnSave','fullCalcOnLoad','forceFullCalc']:
                        if key in node.attrib:
                            if node.attrib[key] not in ['0','1','false','true']:
                                raise ValueError('Invalid calculation scheduling flag')
                            del node.attrib[key]
                if tree(old_book) != tree(new_book):
                    raise ValueError('Unrelated workbook setting changed')
        prior = ET.fromstring(original.read(worksheet))
        current = ET.fromstring(exported.read(worksheet))
        indexed = cells(current)
        if without_targets(prior, targets) != without_targets(current, targets):
            raise ValueError('Unrelated worksheet metadata, cells or caches changed')
        for case in expected['cases']:
            ref = case['cell']
            cell = indexed.get(ref)
            if cell is None:
                raise ValueError('Missing target: ' + ref)
            formulas, values = cell.findall(NS + 'f'), cell.findall(NS + 'v')
            if len(formulas) != 1 or len(values) != 1 or formulas[0].attrib or formulas[0].text != case['formula']:
                raise ValueError('Formula/cache shape changed: ' + ref)
            allowed = {NS + 'f', NS + 'v'}
            if any(child.tag not in allowed for child in cell) or set(cell.attrib) - {'r', 's', 't'}:
                raise ValueError('Unexpected target metadata: ' + ref)
            kind, text = case['type'], values[0].text or ''
            if kind in ('number', 'date'):
                wanted = serial(case, expected['calendar']) if kind == 'date' else case['value']
                if isinstance(wanted, bool) or cell.get('t') not in (None, 'n'):
                    raise ValueError('Numeric cache type changed: ' + ref)
                actual = float(text)
                if not math.isfinite(actual) or struct.pack('>d', actual) != struct.pack('>d', float(wanted)):
                    raise ValueError('Numeric binary64 cache changed: ' + ref)
                value = actual
            elif kind == 'boolean':
                wanted = '1' if case['value'] is True else '0'
                if type(case['value']) is not bool or cell.get('t') != 'b' or text != wanted:
                    raise ValueError('Boolean cache type/value changed: ' + ref)
                value = case['value']
            elif kind in ('string', 'error'):
                if cell.get('t') != ('str' if kind == 'string' else 'e') or text != case['value']:
                    raise ValueError('Text/error cache type/value changed: ' + ref)
                value = text
            else:
                raise ValueError('Unknown expected scalar type')
            readback.append({'cell': ref, 'formula': case['formula'], 'type': kind, 'value': value})
    return {'passed': True, 'sourceSha256': digest(source), 'savedSha256': digest(saved),
            'calendar': expected['calendar'], 'targets': readback,
            'typedCaches': True, 'numericBinary64': True,
            'allUnrelatedPartsByteExact': True,
            'allUnrelatedCellsTypedAndBinary64CachesPreserved': True,
            'cacheLexicalBoundary': 'Existing ordinary numeric formula caches may spell9 as9.0; their types, structure and exact binary64 bits remain unchanged'}


if __name__ == '__main__':
    with open(sys.argv[3], encoding='utf-8') as f:
        expected = json.load(f)
    print(json.dumps(check(sys.argv[1], sys.argv[2], expected), ensure_ascii=False))
