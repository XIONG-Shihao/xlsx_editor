#!/usr/bin/env python3
"""Bounded-memory point-edit oracle. Unchanged package parts and cell envelopes stay authoritative.
Usage: SOURCE SAVED WORKSHEET_PART TARGET_REF... . This is not an external-consumer equivalence oracle.
"""
import hashlib
import importlib.util
import json
import pathlib
import sys
import xml.etree.ElementTree as E
import zipfile

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def literal_data(file):
    """External-consumer profile for formula-free large samples; no full cell map."""
    spec = importlib.util.spec_from_file_location('consumer_helpers', pathlib.Path(__file__).with_name('xlsx-consumer-inspect.py'))
    helpers = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helpers)
    exact, binary = hashlib.sha256(), hashlib.sha256()
    physical, meaningful = 0, 0
    with zipfile.ZipFile(file) as archive:
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            with archive.open('xl/sharedStrings.xml') as stream:
                root = None
                for event, node in E.iterparse(stream, events=('start', 'end')):
                    if root is None:
                        root = node
                    if event == 'end' and node.tag == S+'si':
                        strings.append(helpers.decode(helpers.visible_text(node)))
                        root.clear()
        book = E.fromstring(archive.read('xl/workbook.xml'))
        links = {n.attrib['Id']: n.attrib['Target'] for n in E.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        sheets = [(n.attrib['name'], n.attrib.get('state', 'visible')) for n in book.find(S+'sheets')]
        for sheet in book.find(S+'sheets'):
            target = links[sheet.attrib[R+'id']]
            part = target.lstrip('/') if target.startswith('/') else 'xl/'+target
            with archive.open(part) as stream:
                data = None
                for event, node in E.iterparse(stream, events=('start', 'end')):
                    if event == 'start' and node.tag == S+'sheetData':
                        data = node
                    if event != 'end':
                        continue
                    if node.tag == S+'c':
                        physical += 1
                        if node.find(S+'f') is not None:
                            raise ValueError('Large consumer data profile requires formula-free worksheets')
                        kind = node.attrib.get('t', 'n')
                        raw = node.findtext(S+'v')
                        value = raw
                        if kind == 's':
                            kind, value = 'text', strings[int(raw)]
                        elif kind == 'inlineStr':
                            kind, value = 'text', helpers.decode(helpers.visible_text(node.find(S+'is')))
                        elif kind == 'str':
                            kind, value = 'text', helpers.decode(raw)
                        scalar = helpers.exact_number(raw) if kind == 'n' and raw is not None else value
                        machine = float(raw).hex() if kind == 'n' and raw is not None else value
                        if raw is not None or kind == 'text':
                            meaningful += 1
                            for digest_value, payload in [(exact, scalar), (binary, machine)]:
                                digest_value.update(json.dumps([sheet.attrib['name'], node.attrib['r'], kind, payload], ensure_ascii=False, separators=(',', ':')).encode())
                                digest_value.update(b'\n')
                        node.clear()
                    elif node.tag == S+'row':
                        node.clear()
                        data.clear()
        properties = book.find(S+'workbookPr')
        date1904 = properties.attrib.get('date1904', '0') in ['1', 'true'] if properties is not None else False
    return {'physicalCells': physical, 'nonblankTypedCells': meaningful, 'exactDecimalDigest': exact.hexdigest(), 'binary64Digest': binary.hexdigest(), 'sheets': sheets, 'date1904': date1904}


def compare_literal_data(source, saved):
    before, after = literal_data(source), literal_data(saved)
    return {'passed': all(before[k] == after[k] for k in ['nonblankTypedCells', 'exactDecimalDigest', 'sheets', 'date1904']), 'binary64DataRetained': before['binary64Digest'] == after['binary64Digest'], 'scope': 'Formula-free large-sample typed literal values, coordinates, sheet names/visibility and date system. Physical blanks, fonts/styles/layout and unsupported objects are not qualified by this data-only profile.', 'source': source, 'saved': saved, 'before': before, 'after': after}


def tree(node):
    return [node.tag, sorted(node.attrib.items()), node.text or '', [tree(child) for child in node]]


def digest(stream):
    result = hashlib.sha256()
    while chunk := stream.read(1024 * 1024):
        result.update(chunk)
    return result.hexdigest()


def worksheet(archive, part, targets, allow_caches=False):
    result = hashlib.sha256()
    semantic = hashlib.sha256()
    cells = 0
    formulas = 0
    envelopes = {}
    def feed(value, normalized=None):
        for output, record in [(result, value), (semantic, value if normalized is None else normalized)]:
            output.update(json.dumps(record, ensure_ascii=False, separators=(',', ':')).encode())
            output.update(b'\n')
    with archive.open(part) as stream:
        root = None
        data = None
        for event, node in E.iterparse(stream, events=('start', 'end')):
            if root is None:
                root = node
            if event == 'start' and node.tag == S + 'sheetData':
                data = node
            if event == 'start' and node.tag == S + 'row':
                feed(['row', sorted(node.attrib.items())])
            if event != 'end':
                continue
            if node.tag == S + 'c':
                ref = node.attrib['r']
                cells += 1
                if ref in targets:
                    envelopes[ref] = [sorted((k, v) for k, v in node.attrib.items() if k != 't'), [tree(child) for child in node if child.tag not in [S+'v', S+'is', S+'f']]]
                else:
                    normalized = None
                    if allow_caches and node.find(S+'f') is not None:
                        formulas += 1
                        clone = E.fromstring(E.tostring(node))
                        clone.attrib.pop('t', None)
                        for cache in list(clone.findall(S+'v')):
                            clone.remove(cache)
                        formula = clone.find(S+'f')
                        # The inactive calculation default has no semantic effect.
                        # Active or unknown attributes are never discarded.
                        if formula.attrib.get('aca') in ['0', 'false']:
                            formula.attrib.pop('aca')
                        normalized = tree(clone)
                    feed(tree(node), normalized)
                node.clear()
            elif node.tag == S + 'row':
                node.clear()
                data.clear()
        feed(tree(root))
    return {'cells': cells, 'untouchedDigest': result.hexdigest(), 'semanticDigest': semantic.hexdigest(), 'formulaCacheCells': formulas, 'targetEnvelopes': envelopes}


def compare(source, saved, part, targets, allow_caches=False):
    errors = []
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(saved) as b:
        left, right = set(a.namelist()), set(b.namelist())
        added, removed = sorted(right-left), sorted(left-right)
        if added or any(name != 'xl/calcChain.xml' for name in removed):
            errors.append('Unexpected package membership change')
        changed = []
        for name in sorted(left & right):
            with a.open(name) as original, b.open(name) as returned:
                if digest(original) == digest(returned):
                    continue
            changed.append(name)
            if name not in [part, 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml']:
                errors.append('Unexpected changed part: '+name)
            elif name == 'xl/workbook.xml':
                roots = [E.fromstring(z.read(name)) for z in [a, b]]
                for root in roots:
                    for child in root.findall(S+'calcPr'):
                        for attribute in ['calcId', 'fullCalcOnLoad', 'forceFullCalc', 'calcMode', 'calcOnSave', 'calcCompleted']:
                            child.attrib.pop(attribute, None)
                        if not child.attrib and not list(child) and not (child.text or '').strip():
                            root.remove(child)
                if tree(roots[0]) != tree(roots[1]):
                    errors.append('Workbook content changed outside calculation scheduling')
            elif name != part:
                # The only permitted graph cleanup is an optional calculation-chain removal.
                roots = [E.fromstring(z.read(name)) for z in [a, b]]
                for root in roots:
                    for child in list(root):
                        if child.attrib.get('PartName') == '/xl/calcChain.xml' or child.attrib.get('Type', '').endswith('/calcChain'):
                            root.remove(child)
                if tree(roots[0]) != tree(roots[1]):
                    errors.append('Unqualified package graph change: '+name)
        before, after = worksheet(a, part, set(targets), allow_caches), worksheet(b, part, set(targets), allow_caches)
        key = 'semanticDigest' if allow_caches else 'untouchedDigest'
        if any(before[name] != after[name] for name in [key, 'cells', 'targetEnvelopes', 'formulaCacheCells']):
            errors.append('Untouched worksheet content or target envelope changed')
        if set(before['targetEnvelopes']) != set(targets):
            errors.append('Point targets must already exist in source')
    return {'passed': not errors, 'scope': 'Point-edit preservation: target content, calculation scheduling and optional chain cleanup are permitted fields. When explicitly requested, recalculated caches and omission of inactive aca=false/0 are permitted; authored formulas, other attributes/extensions, row metadata/settings and unrelated member bytes remain protected. Exact and semantic digests are reported separately.', 'allowRecalculatedCaches': allow_caches, 'source': source, 'saved': saved, 'part': part, 'targets': targets, 'changedParts': changed, 'added': added, 'removed': removed, 'before': before, 'after': after, 'errors': errors}


def compare_sheet_append(source,saved,new_name):
    errors=[]
    def parse(data):return E.fromstring(data,E.XMLParser(target=E.TreeBuilder(insert_comments=True,insert_pis=True)))
    with zipfile.ZipFile(source) as a,zipfile.ZipFile(saved) as b:
        books=[parse(z.read('xl/workbook.xml')) for z in [a,b]]
        old_sheets=list(books[0].find(S+'sheets'));new_sheets=list(books[1].find(S+'sheets'))
        added=[n for n in new_sheets if n.attrib.get('name')==new_name and all(o.attrib.get('name')!=new_name for o in old_sheets)]
        if len(added)!=1 or len(new_sheets)!=len(old_sheets)+1 or new_sheets[-1] is not added[0]:return {'passed':False,'errors':['Exactly one worksheet must append at the end']}
        entry=added[0]
        if set(entry.attrib)!={'name','sheetId',R+'id'}:errors.append('New worksheet carries foreign attributes')
        books[1].find(S+'sheets').remove(entry)
        for book in books:
            for calc in book.findall(S+'calcPr'):
                for flag in ['fullCalcOnLoad','forceFullCalc']:calc.attrib.pop(flag,None)
                if not calc.attrib and not list(calc) and not (calc.text or '').strip():book.remove(calc)
        if tree(books[0])!=tree(books[1]):errors.append('Existing workbook metadata changed')
        relationships=[parse(z.read('xl/_rels/workbook.xml.rels')) for z in [a,b]]
        links=[n for n in relationships[1] if n.attrib.get('Id')==entry.attrib[R+'id']]
        if len(links)!=1:return {'passed':False,'errors':['New worksheet binding is ambiguous']}
        relation=links[0];target=relation.attrib.get('Target','');part=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        if relation.attrib.get('Type')!='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' or relation.attrib.get('TargetMode','Internal')!='Internal':errors.append('Invalid appended worksheet relationship')
        relationships[1].remove(relation)
        if tree(relationships[0])!=tree(relationships[1]):errors.append('Existing workbook relationships changed')
        types=[parse(z.read('[Content_Types].xml')) for z in [a,b]]
        overrides=[n for n in types[1] if n.attrib.get('PartName')=='/'+part]
        if len(overrides)!=1 or overrides[0].attrib.get('ContentType')!='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml':errors.append('Invalid appended worksheet content type')
        else:types[1].remove(overrides[0])
        if tree(types[0])!=tree(types[1]):errors.append('Existing content types changed')
        if set(b.namelist())-set(a.namelist())!={part} or set(a.namelist())-set(b.namelist()):errors.append('Unexpected package membership change')
        new=E.fromstring(b.read(part))
        dimensions=new.findall(S+'dimension')
        if any(n.attrib!={'ref':'A1'} or list(n) for n in dimensions) or len(dimensions)>1:errors.append('Appended blank worksheet dimension is invalid')
        for n in dimensions:new.remove(n)
        if new.tag!=S+'worksheet' or new.attrib or len(new)!=1 or new[0].tag!=S+'sheetData' or new[0].attrib or list(new[0]):errors.append('Appended worksheet is not a blank standard worksheet')
        checked=0
        for name in a.namelist():
            if name in ['xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml']:continue
            with a.open(name) as x,b.open(name) as y:
                if digest(x)!=digest(y):errors.append('Existing part changed: '+name)
            checked+=1
    return {'passed':not errors,'source':source,'saved':saved,'newWorksheet':new_name,'newPart':part,'unchangedPartsChecked':checked,'errors':errors,'scope':'One blank worksheet appended at the end; existing worksheet/formula/object bytes and original workbook/OPC metadata remain, with only two calculation refresh flags allowed.'}


if __name__ == '__main__':
    if sys.argv[1] == '--consumer-data':
        report = compare_literal_data(sys.argv[2], sys.argv[3])
    elif sys.argv[1] == '--append-sheet':
        report=compare_sheet_append(sys.argv[2],sys.argv[3],sys.argv[4])
    else:
        arguments = sys.argv[4:]
        allow_caches = '--allow-recalculated-caches' in arguments
        targets = [value for value in arguments if value != '--allow-recalculated-caches']
        report = compare(sys.argv[1], sys.argv[2], sys.argv[3], targets, allow_caches)
    print(json.dumps(report, ensure_ascii=False))
    sys.exit(0 if report['passed'] else 1)
