"""Independent closed admission/preservation checker; not Excel runtime evidence."""
import io
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as E
import zipfile

source, saved, mapping = map(pathlib.Path, sys.argv[1:])
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RICH = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
RICH2 = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2'
XDA = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
GRAPH = {
    'xl/metadata.xml': ('metadata.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'),
    'xl/richData/rdrichvalue.xml': ('richData/rdrichvalue.xml', 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue', 'application/vnd.ms-excel.rdrichvalue+xml'),
    'xl/richData/rdrichvaluestructure.xml': ('richData/rdrichvaluestructure.xml', 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure', 'application/vnd.ms-excel.rdrichvaluestructure+xml'),
    'xl/richData/rdRichValueTypes.xml': ('richData/rdRichValueTypes.xml', 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes', 'application/vnd.ms-excel.rdrichvaluetypes+xml'),
}


def parts(path):
    with zipfile.ZipFile(path) as z:
        assert len(z.namelist()) == len(set(z.namelist())), 'Duplicate package part'
        return {name: z.read(name) for name in z.namelist()}


def parse(data):
    assert not re.search(br'<!DOCTYPE|<!ENTITY', data, re.I)
    pending, stack, scopes, root = {}, [], {}, None
    for event, item in E.iterparse(io.BytesIO(data), events=['start', 'end', 'start-ns']):
        if event == 'start-ns':
            pending[item[0]] = item[1]
        elif event == 'start':
            current = dict(stack[-1] if stack else {})
            current.update(pending)
            pending = {}
            scopes[item] = current
            stack.append(current)
            if root is None:
                root = item
        else:
            stack.pop()
    assert root is not None
    return root, scopes


def identical(old, new, old_scopes, new_scopes):
    assert old.tag == new.tag and old.attrib == new.attrib
    assert old.text == new.text and old.tail == new.tail, 'Opaque text/tail changed'
    assert all(new_scopes[new].get(k) == v for k, v in old_scopes[old].items()), 'Original namespace binding changed'
    assert len(old) == len(new)
    for a, b in zip(old, new, strict=True):
        identical(a, b, old_scopes, new_scopes)


def appended(old, new, old_scopes, new_scopes):
    assert old.tag == new.tag
    assert {k: v for k, v in old.attrib.items() if k != 'count'} == {k: v for k, v in new.attrib.items() if k != 'count'}
    assert old.text == new.text and old.tail == new.tail
    assert all(new_scopes[new].get(k) == v for k, v in old_scopes[old].items())
    assert len(new) >= len(old)
    for a, b in zip(old, list(new)[:len(old)], strict=True):
        identical(a, b, old_scopes, new_scopes)


def index(value, one=False):
    if isinstance(value, str):
        assert re.fullmatch(r'(?:0|[1-9][0-9]*)', value), 'Noncanonical metadata index'
        value = int(value)
    assert type(value) is int and value >= (1 if one else 0), 'Invalid metadata index'
    return value - 1 if one else value


def children(node, tag):
    items = list(node)
    assert all(child.tag == tag for child in items), 'Unqualified metadata child namespace'
    return items


def counted(node, tag):
    items = children(node, tag)
    assert index(node.attrib['count']) == len(items), 'Incorrect metadata count'
    return items


def unique(parent, tag, required=True):
    items = parent.findall(tag)
    assert len(items) <= 1 and (not required or len(items) == 1)
    return items[0] if items else None


def whitespace(node):
    assert not (node.text or '').strip(), 'Unexpected spill metadata text'
    assert all(not (item.tail or '').strip() for item in node), 'Unexpected spill metadata tail'


def leaf(node):
    whitespace(node)
    assert len(node) == 0, 'Unexpected spill metadata leaf child'


def extension_leaf(block, uri, tag):
    assert block.attrib == {}, 'Unexpected spill block attributes'
    whitespace(block)
    items = children(block, f'{{{NS}}}extLst')
    assert len(items) == 1 and items[0].attrib == {}
    whitespace(items[0])
    extensions = children(items[0], f'{{{NS}}}ext')
    assert len(extensions) == 1 and extensions[0].attrib == {'uri': uri}, 'Unqualified spill extension URI'
    whitespace(extensions[0])
    records = children(extensions[0], tag)
    assert len(records) == 1, 'Unexpected spill extension graph'
    leaf(records[0])
    return records[0]


def referenced_record(block):
    assert block.attrib == {}
    whitespace(block)
    records = children(block, f'{{{NS}}}rc')
    assert len(records) == 1 and set(records[0].attrib) == {'t', 'v'}
    leaf(records[0])
    return records[0]


def no_unused_append(old_count, items, required, label):
    assert set(range(old_count, len(items))) <= required, f'Unused added {label}'


def old_container(name, future_name=None):
    if 'xl/metadata.xml' not in before:
        return None
    root, _ = parse(before['xl/metadata.xml'])
    matches = [item for item in root if item.tag == f'{{{NS}}}{name}' and
               (future_name is None or item.attrib.get('name') == future_name)]
    assert len(matches) <= 1
    return matches[0] if matches else None


def cell(value):
    assert isinstance(value, str)
    found = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]*)', value)
    assert found, 'Noncanonical spill coordinate'
    column = 0
    for char in found[1]:
        column = column * 26 + ord(char) - ord('A') + 1
    row = int(found[2])
    assert column <= 16384 and row <= 1048576, 'Spill coordinate outside worksheet'
    return row, column


def rectangle(value):
    assert isinstance(value, str)
    refs = value.split(':')
    assert len(refs) in [1, 2]
    first, last = cell(refs[0]), cell(refs[-1])
    rows, columns = last[0] - first[0] + 1, last[1] - first[1] + 1
    assert rows > 0 and columns > 0 and rows * columns <= 1000, 'Unqualified spill geometry'
    return first, rows, columns


before, after = parts(source), parts(saved)
request = json.loads(mapping.read_text())
assert isinstance(request, dict)
assert all(isinstance(key, str) and isinstance(master, dict) and
           set(master) <= {'cm', 'vm', 'range', 'aca', 'ca', 'intendedRange'} and 'cm' in master
           for key, master in request.items()), 'Unexpected spill mapping'
allowed = set(GRAPH) | {'xl/_rels/workbook.xml.rels', '[Content_Types].xml'}
assert not set(before)-set(after), 'No package part removal is authorized'
assert set(after)-set(before) <= set(GRAPH), 'Unexpected new package member'
for name in before:
    if name not in allowed or name == 'xl/richData/rdRichValueTypes.xml':
        assert before[name] == after[name], f'Unrelated part changed: {name}'
        continue
    old, old_ns = parse(before[name])
    new, new_ns = parse(after[name])
    if name == 'xl/metadata.xml':
        assert old.tag == new.tag and old.attrib == new.attrib and old.text == new.text and old.tail == new.tail
        assert all(new_ns[new].get(k) == v for k, v in old_ns[old].items())
        matched_positions = []
        for child in old:
            matches = [item for item in new if item.tag == child.tag and item.attrib.get('name') == child.attrib.get('name')]
            assert len(matches) == 1
            matched_positions.append(list(new).index(matches[0]))
            appendable = child.tag in [f'{{{NS}}}metadataTypes', f'{{{NS}}}cellMetadata', f'{{{NS}}}valueMetadata']
            appendable |= child.tag == f'{{{NS}}}futureMetadata' and child.attrib.get('name') in ['XLDAPR', 'XLRICHVALUE']
            if appendable:
                appended(child, matches[0], old_ns, new_ns)
            else:
                identical(child, matches[0], old_ns, new_ns)
        assert matched_positions == sorted(matched_positions), 'Original metadata container order changed'
    else:
        appended(old, new, old_ns, new_ns)

metadata, scopes = parse(after['xl/metadata.xml'])
assert metadata.tag == f'{{{NS}}}metadata'
known = {f'{{{NS}}}{name}' for name in ['metadataTypes', 'futureMetadata', 'cellMetadata', 'valueMetadata', 'extLst']}
assert all(child.tag in known for child in metadata), 'Unknown added root metadata container'
types = counted(unique(metadata, f'{{{NS}}}metadataTypes'), f'{{{NS}}}metadataType')
assert all(t.attrib.get('name') for t in types) and len({t.attrib['name'] for t in types}) == len(types)
future = {}
for child in metadata.findall(f'{{{NS}}}futureMetadata'):
    name = child.attrib['name']
    assert name and name not in future
    future[name] = counted(child, f'{{{NS}}}bk')
cells = counted(unique(metadata, f'{{{NS}}}cellMetadata'), f'{{{NS}}}bk')
value_node = unique(metadata, f'{{{NS}}}valueMetadata', False)
values = counted(value_node, f'{{{NS}}}bk') if value_node is not None else []
for block in cells + values:
    for rc in children(block, f'{{{NS}}}rc'):
        kind = types[index(rc.attrib['t'], True)]
        assert index(rc.attrib['v']) < len(future[kind.attrib['name']])

rels, _ = parse(after['xl/_rels/workbook.xml.rels'])
assert rels.tag == f'{{{REL}}}Relationships'
relations = children(rels, f'{{{REL}}}Relationship')
assert all(r.attrib.get('Id') for r in relations) and len({r.attrib['Id'] for r in relations}) == len(relations)
content, _ = parse(after['[Content_Types].xml'])
assert content.tag == f'{{{CT}}}Types'
assert all(item.tag in [f'{{{CT}}}Override', f'{{{CT}}}Default'] for item in content)
overrides = content.findall(f'{{{CT}}}Override')
assert len({r.attrib['PartName'] for r in overrides}) == len(overrides)
old_rel_count = len(parse(before['xl/_rels/workbook.xml.rels'])[0])
old_type_count = len(parse(before['[Content_Types].xml'])[0])
required_parts = {'xl/metadata.xml'} if request else set()
if any('vm' in master for master in request.values()):
    required_parts |= set(GRAPH) - {'xl/metadata.xml'}
added_parts = required_parts - set(before)
assert set(after) - set(before) == added_parts, 'Unrequested new spill package part'
assert len(relations) - old_rel_count == len(added_parts), 'Unexpected added spill relationships'
assert len(content) - old_type_count == len(added_parts), 'Unexpected added spill content types'
for r in relations[old_rel_count:]:
    assert any(r.attrib == {'Id': r.attrib['Id'], 'Type': typ, 'Target': target}
               for part, (target, typ, _) in GRAPH.items() if part in added_parts), 'Invalid added spill relationship'
for r in list(content)[old_type_count:]:
    assert any(r.tag == f'{{{CT}}}Override' and r.attrib == {'PartName': '/'+part, 'ContentType': typ}
               for part, (_, _, typ) in GRAPH.items() if part in added_parts), 'Invalid added spill content type'
for part, (target, relation_type, content_type) in GRAPH.items():
    if part not in after:
        continue
    assert len([r for r in relations if r.attrib.get('Type') == relation_type and r.attrib.get('Target') == target and r.attrib.get('TargetMode', 'Internal') == 'Internal']) == 1
    assert len([r for r in overrides if r.attrib.get('PartName') == '/'+part and r.attrib.get('ContentType') == content_type]) == 1

rich_values, rich_structures = [], []
if 'xl/richData/rdrichvalue.xml' in after:
    root, _ = parse(after['xl/richData/rdrichvalue.xml'])
    assert root.tag == f'{{{RICH}}}rvData'
    rich_values = counted(root, f'{{{RICH}}}rv')
    root, _ = parse(after['xl/richData/rdrichvaluestructure.xml'])
    assert root.tag == f'{{{RICH}}}rvStructures'
    rich_structures = counted(root, f'{{{RICH}}}s')
    for value in rich_values:
        structure = rich_structures[index(value.attrib['s'])]
        assert len(children(value, f'{{{RICH}}}v')) == len(children(structure, f'{{{RICH}}}k'))
    rich_types, rich_type_scopes = parse(after['xl/richData/rdRichValueTypes.xml'])
    assert rich_types.tag == f'{{{RICH2}}}rvTypesInfo'
    if 'xl/richData/rdRichValueTypes.xml' not in before:
        mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        assert rich_types.attrib == {f'{{{mc}}}Ignorable': 'x'}
        whitespace(rich_types)
        assert rich_type_scopes[rich_types].get('x') == NS and rich_type_scopes[rich_types].get('mc') == mc
        assert len(children(rich_types, f'{{{RICH2}}}global')) == 1
        global_node = unique(rich_types, f'{{{RICH2}}}global')
        whitespace(global_node)
        assert global_node.attrib == {} and len(children(global_node, f'{{{RICH2}}}keyFlags')) == 1
        flags = unique(global_node, f'{{{RICH2}}}keyFlags')
        whitespace(flags)
        assert flags.attrib == {}
        names = ['_Self', '_DisplayString', '_Flags', '_Format', '_SubLabel', '_Attribution', '_Icon', '_Display', '_CanonicalPropertyNames', '_ClassificationId']
        keys = children(flags, f'{{{RICH2}}}key')
        assert [item.attrib for item in keys] == [{'name': name} for name in names]
        for item in keys:
            whitespace(item)
            expected = [{'name': 'ExcludeFromFile', 'value': '1'}] if item.attrib['name'] == '_Self' else []
            expected.append({'name': 'ExcludeFromCalcComparison', 'value': '1'})
            assert [flag.attrib for flag in children(item, f'{{{RICH2}}}flag')] == expected
            for flag in item:
                leaf(flag)

result = {}
needed = {'types': set(), 'cells': set(), 'values': set(), 'XLDAPR': set(),
          'XLRICHVALUE': set(), 'richValues': set(), 'richStructures': set()}
for key, master in request.items():
    sheet, origin = key.split('\0')
    assert sheet.startswith('xl/worksheets/') and sheet.endswith('.xml')
    origin = cell(origin)
    first, _, _ = rectangle(master['range'])
    assert first == origin, 'Mapped spill range origin changed'
    if 'aca' in master or 'ca' in master:
        assert 'vm' in master and master.get('aca') == master.get('ca') == '1', 'Invalid mapped calculation flags'
    if 'vm' in master:
        assert master['range'] == key.split('\0')[1], 'Blocked spill published range must be its master'
        intended, intended_rows, intended_columns = rectangle(master['intendedRange'])
        assert intended == origin, 'Intended spill origin changed'
    else:
        assert not set(master) & {'intendedRange', 'aca', 'ca'}, 'Calculated spill has blocked metadata'
    cell_index = index(master['cm'], True)
    rc = referenced_record(cells[cell_index])
    type_index = index(rc.attrib['t'], True)
    kind = types[type_index]
    assert kind.attrib['name'] == 'XLDAPR' and kind.attrib['cellMeta'] == '1'
    future_index = index(rc.attrib['v'])
    flags = extension_leaf(future['XLDAPR'][future_index],
                           '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}', f'{{{XDA}}}dynamicArrayProperties')
    assert set(flags.attrib) <= {'fDynamic', 'fCollapsed'}
    assert flags.attrib['fDynamic'] == '1' and flags.attrib.get('fCollapsed', '0') == '0'
    needed['types'].add(type_index)
    needed['cells'].add(cell_index)
    needed['XLDAPR'].add(future_index)
    output = {'dynamic': True}
    if 'vm' in master:
        value_index = index(master['vm'], True)
        rc = referenced_record(values[value_index])
        type_index = index(rc.attrib['t'], True)
        kind = types[type_index]
        assert kind.attrib['name'] == 'XLRICHVALUE' and kind.attrib.get('cellMeta', '0') == '0'
        future_index = index(rc.attrib['v'])
        record = extension_leaf(future['XLRICHVALUE'][future_index],
                                '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}', f'{{{RICH}}}rvb')
        assert set(record.attrib) == {'i'}
        rich_index = index(record.attrib['i'])
        value = rich_values[rich_index]
        assert set(value.attrib) == {'s'}
        structure_index = index(value.attrib['s'])
        structure = rich_structures[structure_index]
        keys = children(structure, f'{{{RICH}}}k')
        assert structure.attrib == {'t': '_error'}
        assert all(set(k.attrib) == {'n', 't'} and k.attrib['t'] == 'i' for k in keys)
        assert [k.attrib['n'] for k in keys] == ['colOffset', 'errorType', 'rwOffset', 'subType']
        whitespace(structure)
        for k in keys:
            leaf(k)
        fields = {}
        assert not (value.text or '').strip()
        for k, v in zip(keys, children(value, f'{{{RICH}}}v'), strict=True):
            assert v.attrib == {} and len(v) == 0 and not (v.tail or '').strip()
            fields[k.attrib['n']] = index(v.text)
        assert fields['errorType'] == 8 and fields['subType'] == 1
        assert fields['colOffset'] == intended_columns - 1 and fields['rwOffset'] == intended_rows - 1, 'Rich spill offsets do not match intended geometry'
        needed['types'].add(type_index)
        needed['values'].add(value_index)
        needed['XLRICHVALUE'].add(future_index)
        needed['richValues'].add(rich_index)
        needed['richStructures'].add(structure_index)
        output.update(fields)
    result[key] = output

old_types = old_container('metadataTypes')
old_cells = old_container('cellMetadata')
old_values = old_container('valueMetadata')
for original, items, required, label in [
        (old_types, types, needed['types'], 'metadata type'),
        (old_cells, cells, needed['cells'], 'cell metadata record'),
        (old_values, values, needed['values'], 'value metadata record')]:
    no_unused_append(len(original) if original is not None else 0, items, required, label)

new_type_attributes = {name: '1' for name in ['copy', 'pasteAll', 'pasteValues', 'merge',
                       'splitFirst', 'rowColShift', 'clearFormats', 'clearComments', 'assign', 'coerce']}
new_type_attributes['minSupportedVersion'] = '120000'
for item in types[len(old_types) if old_types is not None else 0:]:
    assert item.attrib['name'] in ['XLDAPR', 'XLRICHVALUE']
    expected = {'name': item.attrib['name'], **new_type_attributes}
    if item.attrib['name'] == 'XLDAPR':
        expected['cellMeta'] = '1'
    assert item.attrib == expected, 'Unqualified appended metadata type'
    leaf(item)

for name, blocks in future.items():
    original = old_container('futureMetadata', name)
    no_unused_append(len(original) if original is not None else 0, blocks,
                     needed.get(name, set()), f'{name} future record')
    if original is None:
        assert name in ['XLDAPR', 'XLRICHVALUE'] and blocks, 'Unused added future container'
        container = next(item for item in metadata if item.tag == f'{{{NS}}}futureMetadata' and item.attrib['name'] == name)
        assert container.attrib == {'name': name, 'count': str(len(blocks))}

old_root = parse(before['xl/metadata.xml'])[0] if 'xl/metadata.xml' in before else None
old_keys = {(item.tag, item.attrib.get('name')) for item in old_root} if old_root is not None else set()
for item in metadata:
    if (item.tag, item.attrib.get('name')) in old_keys:
        continue
    assert item.tag != f'{{{NS}}}extLst', 'Unrequested root extension container'
    if item.tag != f'{{{NS}}}futureMetadata':
        assert item.attrib == {'count': str(len(item))} and len(item), 'Unused added metadata container'
    whitespace(item)
if old_root is None:
    assert metadata.attrib == {}
    whitespace(metadata)

for part, items, required, label in [
        ('xl/richData/rdrichvalue.xml', rich_values, needed['richValues'], 'rich value'),
        ('xl/richData/rdrichvaluestructure.xml', rich_structures, needed['richStructures'], 'rich structure')]:
    old_count = len(parse(before[part])[0]) if part in before else 0
    no_unused_append(old_count, items, required, label)
    if part in after and part not in before:
        root, _ = parse(after[part])
        assert root.attrib == {'count': str(len(items))} and items
        whitespace(root)
print(json.dumps({'passed': True, 'masters': result}))
