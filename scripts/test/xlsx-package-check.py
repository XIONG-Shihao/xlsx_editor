#!/usr/bin/env python3
"""Inspect delivered XLSX independently of JavaScript readers/writers.

Read a JSON request on stdin. A changed worksheet path alone never authorizes
changes to all its cells: permitted targets and formula caches are checked apart.
"""
import importlib.util
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

spec = importlib.util.spec_from_file_location('xlsx_reference', Path(__file__).with_name('xlsx-reference.py'))
reference = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reference)
X = reference.X
NS = reference.NS


def parts(file):
    with zipfile.ZipFile(file) as package:
        return {name: package.read(name) for name in package.namelist()}


def tree(node):
    return [node.tag, sorted(node.attrib.items()), node.text or '', [tree(child) for child in node]]


def worksheet_without_targets_and_caches(data, targets, metadata=None, omitted_blanks=()):
    root = ET.fromstring(data)
    for tag, key in [('autoFilter', 'filter'), ('cols', 'columns'), ('sheetViews', 'pane')]:
        if key in (metadata or {}):
            for node in root.findall('x:' + tag, NS):
                root.remove(node)
    for dimension in list(root.findall('x:dimension', NS)):
        if set(dimension.attrib) == {'ref'} and not list(dimension) and not (dimension.text or '').strip():
            root.remove(dimension)
    for row in root.findall('x:sheetData/x:row', NS):
        if 'hiddenRows' in (metadata or {}):
            row.attrib.pop('hidden', None)
        if 'rowHeights' in (metadata or {}):
            row.attrib.pop('ht', None)
            row.attrib.pop('customHeight', None)
        for cell in list(row):
            if cell.tag != f'{{{X}}}c':
                continue
            if cell.attrib['r'] in targets or cell.attrib['r'] in omitted_blanks:
                row.remove(cell)
            elif cell.find('x:f', NS) is not None:
                for cache in cell.findall('x:v', NS):
                    cell.remove(cache)
                cell.attrib.pop('t', None)
        if not list(row) and set(row.attrib) == {'r'}:
            root.find('x:sheetData', NS).remove(row)
    return tree(root)


def strict_cell_envelopes(before_data, after_data, targets, before_parts, after_parts, style_targets=(), added_targets=(), cleared_targets=()):
    """Point edits own values, not target attributes, extensions, annotations or other cells."""
    parser = lambda: ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
    before, after = ET.fromstring(before_data, parser()), ET.fromstring(after_data, parser())
    failures = []
    def field(cell, package):
        if cell.attrib.get('t') == 's':
            try:
                strings = ET.fromstring(package['xl/sharedStrings.xml'], parser())
                return strings.findall('x:si', NS)[int(cell.find('x:v', NS).text)]
            except (KeyError, ValueError, IndexError, AttributeError):
                raise ValueError('Invalid source shared-string identity')
        return cell.find('x:is', NS)
    def protected(value):
        if value is None:
            return False
        allowed = {'{http://www.w3.org/XML/1998/namespace}space'}
        return any(k not in allowed for k in value.attrib) or any(child.tag != f'{{{X}}}t' or list(child) or any(k not in allowed for k in child.attrib) for child in value)
    def stable(node):
        tag = node.tag if isinstance(node.tag, str) else '#comment' if node.tag is ET.Comment else '#pi'
        return [tag, sorted(node.attrib.items()), node.text or '', node.tail or '', [stable(child) for child in node]]
    old_cells = {c.attrib['r']: c for c in before.findall('x:sheetData/x:row/x:c', NS)}
    new_cells = {c.attrib['r']: c for c in after.findall('x:sheetData/x:row/x:c', NS)}
    def remove(root,cell):
        for row in root.findall('x:sheetData/x:row',NS):
            if cell in list(row):row.remove(cell);return
    for ref in targets:
        old, new = old_cells.get(ref), new_cells.get(ref)
        if ref in added_targets:
            if old is not None or new is None or new.attrib.keys()-({'r','t','s'} if ref in style_targets else {'r','t'}) or new.find('x:f',NS) is not None or any(child.tag not in [f'{{{X}}}v',f'{{{X}}}is'] for child in new) or protected(field(new,after_parts)):
                failures.append({'kind':'invalid-declared-added-literal','cell':ref})
            if new is not None:remove(after,new)
            continue
        if old is None or new is None:
            plain_clear=ref in cleared_targets and old is not None and new is None and old.attrib.keys()<={'r','t','s'} and old.attrib.get('s','0')=='0' and old.find('x:f',NS) is None and not protected(field(old,before_parts)) and all(child.tag in [f'{{{X}}}v',f'{{{X}}}is'] and not child.attrib for child in old)
            if plain_clear:remove(before,old)
            else:failures.append({'kind': 'point-edit-target-presence', 'cell': ref})
            continue
        value = field(old, before_parts)
        if protected(value):
            next_value = field(new, after_parts)
            if next_value is None or stable(value) != stable(next_value):
                failures.append({'kind': 'target-text-metadata-loss', 'cell': ref})
        for tag in ['v', 'f']:
            original_field = old.find('x:' + tag, NS)
            if original_field is not None and (list(original_field) or any(k not in ({'t','si','ref'} if tag == 'f' else set()) for k in original_field.attrib)):
                next_field = new.find('x:' + tag, NS)
                if next_field is None or stable(original_field) != stable(next_field):
                    failures.append({'kind': 'target-value-metadata-loss', 'cell': ref})
        for cell in (old, new):
            cell.attrib.pop('t', None)
            if ref in style_targets:
                cell.attrib.pop('s', None)
            for child in list(cell):
                if child.tag in [f'{{{X}}}{tag}' for tag in ('v', 'f', 'is')]:
                    cell.remove(child)
        if stable(old) != stable(new):
            failures.append({'kind': 'target-envelope-loss', 'cell': ref})
    old_rows={row.attrib['r'] for row in ET.fromstring(before_data).findall('x:sheetData/x:row',NS)}
    new_rows={row.attrib['r'] for row in ET.fromstring(after_data).findall('x:sheetData/x:row',NS)}
    for root,other_rows in [(before,new_rows),(after,old_rows)]:
        data=root.find('x:sheetData',NS)
        if data is not None:
            for row in list(data):
                if row.tag==f'{{{X}}}row' and set(row.attrib)=={'r'} and row.attrib['r'] not in other_rows and not list(row) and not (row.text or '').strip():data.remove(row)
        # Ordinary dimensions are derived bounds, not foreign metadata authority.
        for dimension in list(root.findall('x:dimension', NS)):
            if set(dimension.attrib) == {'ref'} and not list(dimension) and not (dimension.text or '').strip():
                root.remove(dimension)
        for cell in root.findall('x:sheetData/x:row/x:c', NS):
            if cell.attrib['r'] not in targets and cell.find('x:f', NS) is not None:
                cell.attrib.pop('t', None)
                for cache in list(cell.findall('x:v', NS)):
                    if not cache.attrib and not list(cache):
                        cell.remove(cache)
    if stable(before) != stable(after):
        failures.append({'kind': 'worksheet-envelope-loss'})
    return failures


def workbook_without_calculation_metadata(data, restore_names=None, remove_new_container=False):
    root = ET.fromstring(data)
    for calc in root.findall('x:calcPr', NS):
        root.remove(calc)
    for container in root.findall('x:definedNames', NS):
        for name in container.findall('x:definedName', NS):
            if name.attrib['name'] in (restore_names or {}):
                previous = restore_names[name.attrib['name']]
                if previous is None:
                    container.remove(name)
                else:
                    name.text = previous
        if remove_new_container and not list(container):
            root.remove(container)
    return tree(root)


def styles_append_only(before, after):
    original, result = ET.fromstring(before), ET.fromstring(after)
    appended = {f'{{{X}}}{name}' for name in ['fonts','fills','borders','cellXfs','numFmts']}
    seen=set()
    for old in original:
        if old.tag in seen:return False
        seen.add(old.tag)
        matches=[new for new in result if new.tag==old.tag]
        if len(matches)!=1:return False
        new=matches[0]
        if old.tag in appended:
            if [tree(n) for n in list(old)] != [tree(n) for n in list(new)[:len(old)]]:return False
            if {k:v for k,v in old.attrib.items() if k!='count'}!={k:v for k,v in new.attrib.items() if k!='count'}:return False
            for extra in list(new)[len(old):]:new.remove(extra)
            new.attrib=dict(old.attrib)
        elif tree(old)!=tree(new):return False
    for new in list(result):
        if new.tag not in seen:
            if new.tag!=f'{{{X}}}numFmts':return False
            result.remove(new)
    return tree(original) == tree(result)


def check(request):
    original = reference.inspect(request['source'], True)
    saved = reference.inspect(request['saved'], True)
    a, b = parts(request['source']), parts(request['saved'])
    removed = sorted(set(a) - set(b))
    added = sorted(set(b) - set(a))
    changed = sorted(name for name in a.keys() & b.keys() if a[name] != b[name])
    allowed = set(request.get('allowedChanged', []))
    failures = []
    chain_rebuild = request.get('allowCalculationChainRebuild') is True
    if added or any(name != 'xl/calcChain.xml' or not chain_rebuild for name in removed):
        failures.append({'kind': 'package-members', 'removed': removed, 'added': added})
    targets = request.get('targets', {})
    worksheet_paths = {item['part']: item['name'] for item in original['sheets']}
    for sheet,refs in request.get('addedLiteralCells',{}).items():
        if not set(refs)<=set(targets.get(sheet,[])) or set(refs)&set(request.get('clearedLiteralCells',{}).get(sheet,[])):failures.append({'kind':'ambiguous-cell-lifecycle-authority','sheet':sheet})
    for sheet,refs in request.get('clearedLiteralCells',{}).items():
        if not set(refs)<=set(targets.get(sheet,[])):failures.append({'kind':'ambiguous-cell-lifecycle-authority','sheet':sheet})

    omitted_blanks = {}
    if request.get('allowDefaultBlankNormalization'):
        for part, sheet in worksheet_paths.items():
            original_root = ET.fromstring(a[part], ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True)))
            after_refs = {c.attrib['r'] for c in ET.fromstring(b[part]).findall('x:sheetData/x:row/x:c', NS)}
            # This small profile has no axes, links, extension payloads or row
            # formatting which could make an explicit blank differ from absence.
            plain = not original_root.attrib and all(isinstance(n.tag, str) for n in original_root.iter()) and all(n.tag in [f'{{{X}}}dimension', f'{{{X}}}sheetData'] for n in original_root)
            refs = set()
            if plain:
                for row in original_root.findall('x:sheetData/x:row', NS):
                    if set(row.attrib) != {'r'}:
                        continue
                    for cell in row:
                        if cell.tag == f'{{{X}}}c' and cell.attrib.get('r') not in after_refs and cell.attrib.keys() <= {'r','s'} and cell.attrib.get('s','0') == '0' and not list(cell) and not (cell.text or '').strip():
                            refs.add(cell.attrib['r'])
            omitted_blanks[sheet] = refs
    formula_targets = {worksheet_paths[name] + '!' + cell.attrib['r']
                       for name in worksheet_paths
                       for cell in ET.fromstring(a[name]).findall('x:sheetData/x:row/x:c', NS)
                       if cell.find('x:f', NS) is not None}
    semantic_checks = []
    calculation_nodes = ET.fromstring(b['xl/workbook.xml']).findall('x:calcPr', NS)
    calculation_properties = calculation_nodes[0].attrib if len(calculation_nodes) == 1 else None
    for name, wanted in request.get('expectedCalculationProperties', {}).items():
        if calculation_properties is None or calculation_properties.get(name) != wanted:
            failures.append({'kind': 'calculation-property', 'property': name, 'expected': wanted,
                             'actual': None if calculation_properties is None else calculation_properties.get(name)})
    absent_cache_checks = []
    for sheet, refs in request.get('absentFormulaCaches', {}).items():
        part = next((part for part, name in worksheet_paths.items() if name == sheet), None)
        cells = {} if part is None else {cell.attrib['r']: cell for cell in ET.fromstring(b[part]).findall('x:sheetData/x:row/x:c', NS)}
        for ref in refs:
            cell = cells.get(ref)
            passed = cell is not None and cell.find('x:f', NS) is not None and not cell.findall('x:v', NS)
            absent_cache_checks.append({'target': sheet + '!' + ref, 'passed': passed})
            if not passed:
                failures.append({'kind': 'formula-cache-not-absent', 'target': sheet + '!' + ref})
    if request.get('strictPointEdits') and any(request.get(key) for key in ['allowedChanged', 'structural', 'rowMappings', 'expectedSheetMetadata', 'expectedNames', 'expectedTables']):
        failures.append({'kind': 'strict-profile-ambiguous-authority'})
    structural = request.get('structural', {})
    metadata = request.get('expectedSheetMetadata', {})
    for part, sheet in worksheet_paths.items():
        if sheet not in metadata:
            continue
        root, expected = ET.fromstring(b[part]), metadata[sheet]
        if 'hiddenRows' in expected:
            rows = [int(row.attrib['r']) for row in root.findall('x:sheetData/x:row', NS) if row.attrib.get('hidden') == '1']
            if rows != expected['hiddenRows']:
                failures.append({'kind': 'hidden-rows', 'sheet': sheet, 'expected': expected['hiddenRows'], 'actual': rows})
        if 'rowHeights' in expected:
            actual = [{'row': row.attrib['r'], 'height': row.attrib['ht'], 'customHeight': row.attrib.get('customHeight')} for row in root.findall('x:sheetData/x:row', NS) if 'ht' in row.attrib]
            if actual != expected['rowHeights']:
                failures.append({'kind': 'row-heights', 'sheet': sheet, 'expected': expected['rowHeights'], 'actual': actual})
        if 'filter' in expected:
            filters = root.findall('x:autoFilter', NS)
            column = filters[0].find('x:filterColumn', NS) if filters else None
            # A filter that is switched on but shows every row has no filterColumn.
            actual = None if not filters else {'ref': filters[0].attrib.get('ref'), 'columnId': None if column is None else column.attrib.get('colId'), 'values': [f.attrib['val'] for f in filters[0].findall('x:filterColumn/x:filters/x:filter', NS)]}
            if len(filters) > 1 or actual != expected['filter']:
                failures.append({'kind': 'filter', 'sheet': sheet, 'expected': expected['filter'], 'actual': actual})
        if 'columns' in expected:
            actual = [col.attrib for col in root.findall('x:cols/x:col', NS)]
            if actual != expected['columns']:
                failures.append({'kind': 'columns', 'sheet': sheet, 'expected': expected['columns'], 'actual': actual})
        if 'pane' in expected:
            panes = root.findall('x:sheetViews/x:sheetView/x:pane', NS)
            actual = panes[0].attrib if panes else None
            if len(panes) > 1 or actual != expected['pane']:
                failures.append({'kind': 'pane', 'sheet': sheet, 'expected': expected['pane'], 'actual': actual})
    expected_names = request.get('expectedNames', {})
    original_names = {item.attrib['name']: item.text for item in ET.fromstring(a['xl/workbook.xml']).findall('x:definedNames/x:definedName', NS)}
    saved_names = {item.attrib['name']: item.text for item in ET.fromstring(b['xl/workbook.xml']).findall('x:definedNames/x:definedName', NS)}
    for name, formula in expected_names.items():
        candidates = [item for item in ET.fromstring(b['xl/workbook.xml']).findall('x:definedNames/x:definedName', NS) if item.attrib.get('name') == name]
        if saved_names.get(name) != formula or len(candidates) != 1:
            failures.append({'kind': 'defined-name', 'name': name, 'expected': formula, 'actual': saved_names.get(name)})
        elif name not in original_names and candidates[0].attrib != {'name': name}:
            failures.append({'kind': 'defined-name-attributes', 'name': name, 'actual': candidates[0].attrib})
    for name, wanted in request.get('expectedTables', {}).items():
        table = ET.fromstring(b[name]) if name in b else None
        if table is None or table.attrib.get('ref') != wanted or any(item.attrib.get('ref') != wanted for item in table.findall('x:autoFilter', NS)):
            failures.append({'kind': 'table-range', 'part': name, 'expected': wanted})
    for name in changed:
        if name in worksheet_paths:
            sheet = worksheet_paths[name]
            if sheet in structural or sheet in request.get('rowMappings', {}):
                # Movement has its own independent cell and literal formula keys below.
                semantic_checks.append({'part': name, 'kind': 'structural-cell-oracle', 'passed': True})
                continue
            if request.get('strictPointEdits'):
                if metadata.get(sheet) or sheet in structural or sheet in request.get('rowMappings', {}):
                    failures.append({'kind': 'strict-profile-excludes-structural-metadata-edits', 'part': name})
                failures.extend(dict(item, part=name) for item in strict_cell_envelopes(a[name], b[name], targets.get(sheet, []), a, b, request.get('styleTargets', {}).get(sheet, []), request.get('addedLiteralCells',{}).get(sheet,[]), request.get('clearedLiteralCells',{}).get(sheet,[])))
            passed = worksheet_without_targets_and_caches(a[name], targets.get(sheet, []), metadata.get(sheet), omitted_blanks.get(sheet, ())) == worksheet_without_targets_and_caches(b[name], targets.get(sheet, []), metadata.get(sheet))
            semantic_checks.append({'part': name, 'kind': 'only-targets-and-formula-caches', 'passed': passed})
            if not passed:
                failures.append({'kind': 'worksheet-semantic-delta', 'part': name})
        elif name == 'xl/workbook.xml':
            restore_names = {key: original_names.get(key) for key in expected_names}
            before_has_container = bool(ET.fromstring(a[name]).findall('x:definedNames', NS))
            passed = workbook_without_calculation_metadata(a[name]) == workbook_without_calculation_metadata(b[name], restore_names, not before_has_container)
            semantic_checks.append({'part': name, 'kind': 'calculation-metadata-only', 'passed': passed})
            if not passed:
                failures.append({'kind': 'workbook-semantic-delta', 'part': name})
        elif name in request.get('expectedTables', {}):
            before_table, after_table = ET.fromstring(a[name]), ET.fromstring(b[name])
            wanted = request['expectedTables'][name]
            passed = after_table.attrib.get('ref') == wanted
            for auto_filter in after_table.findall('x:autoFilter', NS):
                passed = passed and auto_filter.attrib.get('ref') == wanted
                auto_filter.set('ref', before_table.find('x:autoFilter', NS).attrib['ref'])
            after_table.set('ref', before_table.attrib['ref'])
            passed = passed and tree(before_table) == tree(after_table)
            semantic_checks.append({'part': name, 'kind': 'expected-table-range-only', 'passed': passed})
            if not passed:
                failures.append({'kind': 'table-semantic-delta', 'part': name})
        elif chain_rebuild and name in ['[Content_Types].xml', 'xl/_rels/workbook.xml.rels']:
            before_graph, after_graph = ET.fromstring(a[name]), ET.fromstring(b[name])
            for graph in (before_graph, after_graph):
                for child in list(graph):
                    if child.tag.endswith('}Override') and child.attrib.get('PartName') == '/xl/calcChain.xml' or child.tag.endswith('}Relationship') and child.attrib.get('Type', '').endswith('/calcChain'):
                        graph.remove(child)
            passed = tree(before_graph) == tree(after_graph)
            if not passed:
                failures.append({'kind': 'calculation-chain-graph-delta', 'part': name})
            semantic_checks.append({'part': name, 'kind': 'calculation-chain-only', 'passed': passed})
        elif name == 'xl/styles.xml' and request.get('appendStyles'):
            passed = styles_append_only(a[name], b[name])
            semantic_checks.append({'part': name, 'kind': 'original-styles-prefix-retained', 'passed': passed})
            if not passed:
                failures.append({'kind': 'styles-semantic-delta'})
        elif name not in allowed:
            failures.append({'kind': 'unauthorized-member-change', 'part': name})
    if original['dateSystem'] != saved['dateSystem']:
        failures.append({'kind': 'date-system-changed'})

    expected = request.get('expected', {})
    checks = reference.check_literals(saved, {'cells': expected}, calculated=True)
    for result in checks:
        target = result['target']
        key, cell = expected[target], saved['cells'].get(target)
        for property in ('formula', 'rawValue', 'style'):
            if property in key and (cell is None or cell[property] != key[property]):
                result['passed'] = False
                result['mismatches'].append(property)
        result.pop('referenceExclusion', None)
    failures.extend({'kind': 'expected-cell', 'target': result['target'], 'mismatches': result['mismatches']} for result in checks if not result['passed'])

    # Every original cell not explicitly changed must retain its stored meaning.
    # Formula caches may be refreshed/cleared, but formula text and style may not
    # change unless a handwritten structural formula key says how they shift.
    for target, cell in original['cells'].items():
        sheet, ref = target.rsplit('!', 1)
        mapped = ref
        if sheet in structural:
            import re
            column, row = re.fullmatch(r'([A-Z]+)(\d+)', ref).groups()
            rule = structural[sheet]
            mapped = column + str(int(row) + (rule.get('count', 1) if int(row) > rule['index'] else 0))
        elif sheet in request.get('rowMappings', {}):
            import re
            column, row = re.fullmatch(r'([A-Z]+)(\d+)', ref).groups()
            mapped = column + str(request['rowMappings'][sheet].get(row, int(row)))
        after_target = sheet + '!' + mapped
        if mapped in targets.get(sheet, []):
            if mapped not in request.get('styleTargets', {}).get(sheet, []) and saved['cells'].get(after_target, {}).get('style') != cell['style']:
                failures.append({'kind': 'target-style-changed', 'target': target})
            continue
        after = saved['cells'].get(after_target)
        if after is None and ref in omitted_blanks.get(sheet, ()):
            continue
        if after is None:
            failures.append({'kind': 'original-cell-missing', 'target': target, 'mapped': after_target})
            continue
        if target in formula_targets:
            wanted = expected.get(after_target, {}).get('formula', cell['formula'])
            if after['formula'] != wanted or after['style'] != cell['style']:
                failures.append({'kind': 'formula-or-style-changed', 'target': target, 'mapped': after_target})
        elif any(after[prop] != cell[prop] for prop in ('value', 'rawValue', 'formula', 'style', 'type')):
            failures.append({'kind': 'literal-cell-changed', 'target': target, 'mapped': after_target})

    return {'passed': not failures, 'originalSha256': original['sha256'], 'savedSha256': saved['sha256'],
            'bytes': saved['bytes'], 'dateSystem': saved['dateSystem'],
            'memberDelta': {'removed': removed, 'added': added, 'changed': changed},
            'byteIdenticalMembers': len(a) - len(changed) - len(removed),
            'calculationProperties': calculation_properties, 'absentFormulaCacheChecks': absent_cache_checks,
            'semanticChecks': semantic_checks, 'checks': checks, 'failures': failures}


if __name__ == '__main__':
    result = check(json.load(sys.stdin))
    print(json.dumps(result, ensure_ascii=False))
    sys.exit(0 if result['passed'] else 1)
