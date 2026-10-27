"""Independent, stdlib-only OOXML object oracle. No product imports/evaluation.

CLI: inventory FILE; verify BEFORE AFTER EXPECT.json. Prints JSON, exits 1 on failure.
The expectation is authored by the test, never derived from the product's reply.
"""
import argparse
import fnmatch
import hashlib
import json
import math
import pathlib
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
import xml.parsers.expat as expat

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = dict(s=S, d=D, a=A, c=C, r=R, p=P, t=T)
CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
CT_DRAW = 'application/vnd.openxmlformats-officedocument.drawing+xml'
UUID = re.compile(r'^workbook:[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$')


def digest(data):
    return hashlib.sha256(data).hexdigest()


def xml(data):
    if re.search(br'<!DOCTYPE|<!ENTITY', data, re.I):
        raise ValueError('XML entities/DTD are refused by the oracle')
    return ET.fromstring(data)


def part_name(name):
    value = name.rstrip('/')
    if not value or value.startswith('/') or '\\' in value or '\0' in value or any(s in ('', '.', '..') for s in value.split('/')):
        raise ValueError('Invalid ZIP member: ' + name)
    return name


def resolve(owner, target):
    if not target or any(ch in target for ch in ('\\', '?', '#', '%', '\0')):
        raise ValueError('Unqualified relationship target: ' + target)
    value = target[1:] if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
    part_name(value)
    return value


def rels_name(owner):
    return posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner) + '.rels') if owner else '_rels/.rels'


def integer(value, minimum=0, maximum=10**12):
    if value is None or not re.fullmatch(r'-?\d+', str(value)):
        raise ValueError('Expected integer: ' + str(value))
    number = int(value)
    if not minimum <= number <= maximum:
        raise ValueError('Integer outside bound: ' + str(value))
    return number


def marker(node):
    return dict(col=integer(node.findtext('d:col', namespaces=NS), 0, 16383),
                row=integer(node.findtext('d:row', namespaces=NS), 0, 1048575),
                colOff=integer(node.findtext('d:colOff', namespaces=NS)),
                rowOff=integer(node.findtext('d:rowOff', namespaces=NS)))


def anchor_geometry(node):
    kind = node.tag.split('}')[-1]
    if kind not in ('oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'):
        return {'kind': kind, 'protected': True}
    result = {'kind': kind}
    if kind != 'absoluteAnchor':
        result['from'] = marker(node.find('d:from', NS))
    if kind == 'twoCellAnchor':
        result['to'] = marker(node.find('d:to', NS))
        if (result['to']['row'], result['to']['col']) < (result['from']['row'], result['from']['col']):
            raise ValueError('Reversed two-cell anchor')
        result['editAs'] = node.get('editAs', 'twoCell')
    else:
        extent = node.find('d:ext', NS)
        result['extent'] = {'cx': integer(extent.get('cx'), 1), 'cy': integer(extent.get('cy'), 1)}
        if kind == 'absoluteAnchor':
            position = node.find('d:pos', NS)
            result['position'] = {'x': integer(position.get('x')), 'y': integer(position.get('y'))}
    return result


def raw_anchors(data):
    """Keep lexical anchor bytes: canonical XML equality cannot prove preservation."""
    parser = expat.ParserCreate(namespace_separator='}')
    depth, start, result = 0, None, []
    def tag_end(offset):
        quote = None
        for index in range(offset, len(data)):
            value = data[index]
            if quote:
                if value == quote:
                    quote = None
            elif value in (34, 39):
                quote = value
            elif value == 62:
                return index + 1
        raise ValueError('Unterminated raw drawing tag')
    def begin(_tag, _attrs):
        nonlocal depth, start
        if depth == 1:
            start = parser.CurrentByteIndex
        depth += 1
    def end(_tag):
        nonlocal depth, start
        depth -= 1
        if depth == 1:
            first_end = tag_end(start)
            final_end = first_end if data[start:first_end].rstrip().endswith(b'/>') else tag_end(parser.CurrentByteIndex)
            result.append(data[start:final_end]); start = None
    parser.StartElementHandler = begin; parser.EndElementHandler = end; parser.Parse(data, True)
    return result


def cache(container):
    if container is None:
        return None
    count = integer(container.find('c:ptCount', NS).get('val'), 0, 10000)
    values, seen = [None] * count, set()
    numeric = container.tag.endswith('numCache') or container.tag.endswith('numLit')
    for point in container.findall('c:pt', NS):
        index = integer(point.get('idx'), 0, max(0, count - 1))
        if index in seen or index >= count:
            raise ValueError('Duplicate/out-of-range chart point')
        seen.add(index)
        value = point.findtext('c:v', namespaces=NS)
        if numeric:
            value = float(value)
            if not math.isfinite(value):
                raise ValueError('Nonfinite/error chart cache value')
        values[index] = value
    return values


def reference(container):
    if container is None:
        return {'formula': None, 'cache': None}
    cached = container.find('c:numCache', NS)
    if cached is None:
        cached = container.find('c:strCache', NS)
    return {'formula': container.findtext('c:f', namespaces=NS), 'cache': cache(cached)}


def chart_semantics(data):
    root = xml(data)
    if root.tag != '{' + C + '}chartSpace':
        raise ValueError('Invalid chart root')
    reasons = []
    for name in ('externalData', 'pivotSource'):
        if root.find('.//c:' + name, NS) is not None:
            reasons.append(name)
    if any(node.tag.split('}')[-1] == 'extLst' for node in root.iter()):
        reasons.append('extended-chart')
    plot = root.find('c:chart/c:plotArea', NS)
    types = [node for node in plot if node.tag.split('}')[-1].endswith('Chart')]
    kind = types[0].tag.split('}')[-1] if len(types) == 1 else 'combo'
    if kind == 'barChart':
        direction = types[0].find('c:barDir', NS).get('val')
        if direction not in ('col', 'bar') or types[0].find('c:grouping', NS).get('val') != 'clustered':
            reasons.append('unsupported-bar-layout')
        else:
            kind = 'column' if direction == 'col' else 'bar'
            if kind == 'bar':
                category = plot.find('c:catAx/c:axPos', NS)
                value = plot.find('c:valAx/c:axPos', NS)
                if category is None or category.get('val') != 'l' or value is None or value.get('val') != 'b':
                    reasons.append('unsupported-bar-axes')
    elif kind == 'lineChart':
        grouping = types[0].find('c:grouping', NS)
        if grouping is not None and grouping.get('val') != 'standard':
            reasons.append('unsupported-line-layout')
        kind = 'line'
    else:
        reasons.append('unsupported-chart-type')
    series = []
    for chart_type in types:
        for ser in chart_type.findall('c:ser', NS):
            title = ser.find('c:tx', NS)
            tx = reference(title.find('c:strRef', NS) if title is not None else None)
            cat = ser.find('c:cat', NS)
            values = ser.find('c:val', NS)
            cat_ref = cat.find('c:strRef', NS) if cat is not None else None
            if cat_ref is None and cat is not None:
                cat_ref = cat.find('c:numRef', NS)
            val_ref = values.find('c:numRef', NS) if values is not None else None
            series.append({'index': integer(ser.find('c:idx', NS).get('val')),
                           'order': integer(ser.find('c:order', NS).get('val')),
                           'title': tx, 'categories': reference(cat_ref), 'values': reference(val_ref)})
    indices = [item['index'] for item in series]
    orders = [item['order'] for item in series]
    if len(indices) != len(set(indices)) or len(orders) != len(set(orders)):
        raise ValueError('Duplicate series index/order')
    axes = {node.find('c:axId', NS).get('val'): node for node in plot if node.tag.split('}')[-1] in ('catAx', 'valAx')}
    if kind in ('column', 'line'):
        if len(axes) != 2 or any(node.find('c:crossAx', NS).get('val') not in axes for node in axes.values()):
            raise ValueError('Missing/unbound chart axes')
        for chart_type in types:
            if {node.get('val') for node in chart_type.findall('c:axId', NS)} != set(axes):
                raise ValueError('Series/plot axes mismatch')
    title = root.find('c:chart/c:title', NS)
    title_text = ''.join(node.text or '' for node in title.findall('.//a:t', NS)) if title is not None else ''
    legend = root.find('c:chart/c:legend/c:legendPos', NS)
    blanks = root.find('c:chart/c:dispBlanksAs', NS)
    visible = root.find('c:chart/c:plotVisOnly', NS)
    return dict(kind=kind, title=title_text, legend=legend.get('val') if legend is not None else 'none',
                blanks=blanks.get('val') if blanks is not None else None,
                includeHidden=visible is not None and visible.get('val') == '0', series=series, protected=reasons)


class Package:
    def __init__(self, filename):
        with zipfile.ZipFile(filename) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)) or len(names) > 10000:
                raise ValueError('Duplicate/excess ZIP members')
            if sum(item.file_size for item in archive.infolist()) > 512 * 1024 * 1024:
                raise ValueError('Package exceeds independent oracle budget')
            self.parts = {part_name(name): archive.read(name) for name in names}
        self.sha256 = digest(pathlib.Path(filename).read_bytes())
        self.relationships, self.issues, self.objects = {}, [], []
        types = xml(self.parts['[Content_Types].xml'])
        defaults, overrides = {}, {}
        for node in types:
            key = node.get('Extension') if node.tag == '{' + T + '}Default' else node.get('PartName')
            target = defaults if node.tag == '{' + T + '}Default' else overrides
            if key in target:
                raise ValueError('Duplicate content type')
            target[key] = node.get('ContentType')
        self.types = {name: overrides.get('/' + name, defaults.get(name.rsplit('.', 1)[-1])) for name in self.parts}
        for name, data in self.parts.items():
            if not name.endswith('.rels'):
                continue
            owner = '' if name == '_rels/.rels' else name.replace('/_rels/', '/').removesuffix('.rels')
            relations = {}
            for node in xml(data):
                identity = node.get('Id')
                if not identity or identity in relations:
                    raise ValueError('Missing/duplicate relationship ID')
                external = node.get('TargetMode') == 'External'
                target = node.get('Target')
                resolved = None if external else resolve(owner, target)
                if not external and resolved not in self.parts:
                    raise ValueError('Missing relationship target: ' + resolved)
                relations[identity] = dict(type=node.get('Type'), target=target, resolved=resolved, external=external)
            self.relationships[owner] = relations
        workbook = xml(self.parts['xl/workbook.xml'])
        self.sheets = {}
        for sheet in workbook.findall('s:sheets/s:sheet', NS):
            rel = self.relationships['xl/workbook.xml'][sheet.get('{' + R + '}id')]
            if rel['external'] or rel['type'] != R + '/worksheet':
                raise ValueError('Invalid sheet ownership')
            self.sheets[sheet.get('name')] = rel['resolved']
        self._inventory()

    def typed_cells(self):
        result, count = {}, 0
        strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            strings = [''.join(item.itertext()) for item in xml(self.parts['xl/sharedStrings.xml']).findall('s:si', NS)]
        for sheet, part in self.sheets.items():
            values = {}
            for cell in xml(self.parts[part]).findall('s:sheetData/s:row/s:c', NS):
                count += 1
                if count > 10000:
                    raise ValueError('Typed-cell oracle is limited to 10,000 reference-fixture cells')
                raw = cell.findtext('s:v', namespaces=NS)
                kind = cell.get('t', 'n')
                formula = cell.find('s:f', NS)
                value = None
                if kind == 'e':
                    value, kind = raw, 'error'
                elif kind == 'inlineStr':
                    value, kind = ''.join(node.text or '' for node in cell.findall('s:is//s:t', NS)), 'text'
                elif kind == 's':
                    value, kind = strings[integer(raw, 0, len(strings) - 1)], 'text'
                elif kind == 'b':
                    value, kind = bool(integer(raw, 0, 1)), 'boolean'
                elif kind == 'str':
                    value, kind = raw or '', 'text'
                elif raw is not None:
                    value, kind = float(raw), 'number'
                else:
                    kind = 'blank'
                record = {'type': kind, 'value': value}
                if kind == 'number':
                    record['unsafeNumber'] = not math.isfinite(value) or bool(re.fullmatch(r'-?\d+', raw) and abs(int(raw)) > 2**53 - 1)
                if formula is not None:
                    record.update(formula=formula.text or '', status='special' if formula.get('t') else 'unresolved' if raw is None else 'error' if kind == 'error' else 'calculated')
                values[cell.get('r')] = record
            result[sheet] = values
        return result

    def graph_parts(self, start):
        seen, stack, total = set(), set(), 0
        def visit(part):
            nonlocal total
            if part in stack:
                raise ValueError('Cyclic object relationship')
            if part in seen:
                return
            seen.add(part); stack.add(part); total += len(self.parts[part])
            if len(seen) > 128 or total > 16 * 1024 * 1024:
                raise ValueError('Object graph exceeds independent oracle budget')
            for relation in self.relationships.get(part, {}).values():
                if not relation['external']:
                    visit(relation['resolved'])
            stack.remove(part)
        visit(start)
        return seen

    def _inventory(self):
        managed = set()
        for sheet_name, sheet_part in self.sheets.items():
            worksheet = xml(self.parts[sheet_part])
            for drawing in worksheet.findall('s:drawing', NS):
                rel = self.relationships.get(sheet_part, {}).get(drawing.get('{' + R + '}id'))
                if not rel or rel['external'] or rel['type'] != R + '/drawing':
                    raise ValueError('Invalid sheet/drawing relationship')
                part = rel['resolved']
                if self.types[part] != CT_DRAW:
                    raise ValueError('Drawing content type mismatch')
                document = xml(self.parts[part])
                if document.tag != '{' + D + '}wsDr':
                    raise ValueError('Invalid drawing root')
                ids = set()
                raw = raw_anchors(self.parts[part])
                if len(raw) != len(document):
                    raise ValueError('Drawing lexical/parsed anchor count mismatch')
                for index, anchor in enumerate(document):
                    properties = anchor.find('.//d:cNvPr', NS)
                    identity = properties.get('name') if properties is not None else 'opaque-anchor-' + str(len(ids))
                    nonvisual = properties.get('id') if properties is not None else identity
                    if nonvisual in ids:
                        raise ValueError('Duplicate nonvisual drawing ID')
                    ids.add(nonvisual)
                    if UUID.fullmatch(identity):
                        if identity in managed:
                            raise ValueError('Managed object ID reused across document scopes')
                        managed.add(identity)
                    object_data = dict(sheet=sheet_name, id=identity, kind='opaque', drawingPart=part,
                                       anchor=anchor_geometry(anchor), altText=properties.get('descr', '') if properties is not None else '', protected=[])
                    object_data['rawAnchorSha256'] = digest(raw[index])
                    allowed = {'from', 'to', 'pos', 'ext', 'pic', 'graphicFrame', 'clientData'}
                    if any(node.tag.split('}')[-1] not in allowed for node in anchor):
                        object_data['protected'].append('unqualified-anchor-children')
                    if any(node.tag.split('}')[-1] in ('srcRect', 'effectLst', 'effectDag') for node in anchor.iter()):
                        object_data['protected'].append('unsupported-image-effects')
                    chart_link = anchor.find('.//c:chart', NS)
                    image_link = anchor.find('.//a:blip', NS)
                    relationship = None
                    if chart_link is not None:
                        relationship = self.relationships.get(part, {}).get(chart_link.get('{' + R + '}id'))
                        object_data['kind'] = 'chart'
                    elif image_link is not None:
                        relationship = self.relationships.get(part, {}).get(image_link.get('{' + R + '}embed'))
                        object_data['kind'] = 'image'
                    else:
                        object_data['protected'].append('opaque-object')
                    owned = []
                    if object_data['kind'] != 'opaque':
                        if relationship is None:
                            raise ValueError('Missing object relationship')
                        if relationship['external']:
                            object_data['protected'].append('external-object-relationship')
                        else:
                            target = relationship['resolved']
                            expected = R + ('/chart' if object_data['kind'] == 'chart' else '/image')
                            if relationship['type'] != expected:
                                raise ValueError('Object relationship type mismatch')
                            object_data['part'] = target
                            owned = sorted(self.graph_parts(target))
                            if object_data['kind'] == 'chart':
                                if self.types[target] != CT_CHART:
                                    raise ValueError('Chart content type mismatch')
                                object_data['chart'] = chart_semantics(self.parts[target])
                                object_data['protected'] += object_data['chart']['protected']
                            else:
                                mime = self.types[target]
                                data = self.parts[target]
                                if mime not in ('image/png', 'image/jpeg'):
                                    object_data['protected'].append('unsupported-image-type')
                                elif not (data.startswith(b'\x89PNG\r\n\x1a\n') if mime == 'image/png' else data.startswith(b'\xff\xd8') and data.endswith(b'\xff\xd9')):
                                    raise ValueError('Invalid media signature/content type')
                                object_data['media'] = dict(mime=mime, bytes=len(data), sha256=digest(data))
                            for edge in self.relationships.get(target, {}).values():
                                if edge['external']:
                                    object_data['protected'].append('external-dependent-relationship')
                                elif edge['resolved'] in (target, part, sheet_part):
                                    raise ValueError('Cyclic object relationship')
                                elif edge['type'] not in (R + '/image', 'http://schemas.microsoft.com/office/2011/relationships/chartStyle', 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'):
                                    object_data['protected'].append('unknown-dependent-relationship')
                                else:
                                    owned.append(edge['resolved'])
                    if any(node.tag.split('}')[-1] == 'extLst' for node in anchor.iter()):
                        object_data['protected'].append('extended-anchor')
                    owned = sorted(set(owned))
                    graph_data = raw[index] + json.dumps(relationship, sort_keys=True).encode()
                    for name in owned:
                        graph_data += name.encode() + self.parts[name]
                        related = rels_name(name)
                        if related in self.parts:
                            graph_data += related.encode() + self.parts[related]
                    object_data['graphHash'] = digest(graph_data)
                    self.objects.append(object_data)
        self.objects.sort(key=lambda item: (item['sheet'], item['id']))
        self.incoming = {}
        for owner, relations in self.relationships.items():
            for identity, relation in relations.items():
                if not relation['external']:
                    self.incoming.setdefault(relation['resolved'], []).append({'owner': owner, 'id': identity})
        for item in self.objects:
            if 'part' in item:
                item['sharedResource'] = len(self.incoming.get(item['part'], [])) > 1
                if item['kind'] == 'chart' and item['sharedResource']:
                    item['protected'].append('shared-mutable-chart')
            if len(self.incoming.get(item['drawingPart'], [])) > 1:
                item['protected'].append('shared-mutable-drawing')

    def report(self):
        return dict(schema=1, packageSha256=self.sha256, objects=self.objects, typedCells=self.typed_cells(),
                    partHashes={name: digest(data) for name, data in sorted(self.parts.items())}, incoming=self.incoming)


def require_subset(actual, expected, location, failures):
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            failures.append(location + ': expected mapping')
            return
        for key, value in expected.items():
            require_subset(actual.get(key), value, location + '.' + key, failures)
    elif isinstance(expected, list):
        if not isinstance(actual, list) or len(actual) != len(expected):
            failures.append(location + ': list size differs')
        else:
            for index, value in enumerate(expected):
                require_subset(actual[index], value, location + '[' + str(index) + ']', failures)
    elif actual != expected:
        failures.append(location + ': expected ' + repr(expected) + ', got ' + repr(actual))


def verify(before_file, after_file, expectation):
    before, after = Package(before_file), Package(after_file)
    failures = []
    prior = {(obj['sheet'], obj['id']): obj for obj in before.objects}
    current = {(obj['sheet'], obj['id']): obj for obj in after.objects}
    actions = expectation.get('actions', [])
    action_keys = {(action['sheet'], action['id']) for action in actions}
    for action in actions:
        key = (action['sheet'], action['id'])
        if action['kind'] == 'create' and (key in prior or key not in current):
            failures.append('create ownership mismatch: ' + repr(key))
        elif action['kind'] == 'update' and (key not in prior or key not in current):
            failures.append('update ownership mismatch: ' + repr(key))
        elif action['kind'] == 'delete' and (key not in prior or key in current):
            failures.append('delete ownership mismatch: ' + repr(key))
        if action['kind'] not in ('create', 'update', 'delete'):
            failures.append('unknown action kind: ' + action['kind'])
        if action['kind'] in ('create', 'update') and key in current and current[key]['protected']:
            failures.append('authoring left a protected object: ' + repr(key))
        if action['kind'] in ('update', 'delete') and key in prior and prior[key]['protected'] and not expectation.get('allowProtectedActions', False):
            failures.append('protected object authoring was not qualified: ' + repr(key))
        if action['kind'] in ('create', 'update') and not UUID.fullmatch(action['id']) and not expectation.get('allowImportedAuthoring', False):
            failures.append('managed authored ID required: ' + repr(key))
    for key in set(prior) | set(current):
        if key not in action_keys and (key not in prior or key not in current or prior[key]['graphHash'] != current[key]['graphHash']):
            failures.append('unscoped object graph changed: ' + repr(key))
    for expected in expectation.get('objects', []):
        key = (expected['sheet'], expected['id'])
        if key not in current:
            failures.append('expected object missing: ' + repr(key))
        else:
            require_subset(current[key], expected, repr(key), failures)
    allowed = expectation.get('allowChangedParts', [])
    for name in set(before.parts) | set(after.parts):
        if not any(fnmatch.fnmatchcase(name, pattern) for pattern in allowed) and before.parts.get(name) != after.parts.get(name):
            failures.append('unapproved package part change: ' + name)
    for name in expectation.get('removedParts', []):
        if name in after.parts:
            failures.append('expected removed part remains: ' + name)
    for name in expectation.get('preservedParts', []):
        if name not in before.parts or before.parts[name] != after.parts.get(name):
            failures.append('preserved bytes differ: ' + name)
    for expected in expectation.get('sharedResources', []):
        if len(after.incoming.get(expected['part'], [])) != expected['references']:
            failures.append('shared-resource reference count differs: ' + expected['part'])
    if 'typedCells' in expectation:
        require_subset(after.typed_cells(), expectation['typedCells'], 'typedCells', failures)
    return dict(schema=1, passed=not failures, failures=failures, sourceSha256=before.sha256,
                outputSha256=after.sha256, objects=after.objects, method='independent-stdlib-ooxml-object-graph')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    inventory = sub.add_parser('inventory'); inventory.add_argument('file')
    media = sub.add_parser('decode-media'); media.add_argument('file')
    check = sub.add_parser('verify'); check.add_argument('before'); check.add_argument('after'); check.add_argument('expectation')
    args = parser.parse_args()
    try:
        if args.command == 'decode-media':
            from PIL import Image
            import io
            package = Package(args.file); decoded = []
            for name, data in package.parts.items():
                if package.types[name] in ('image/png', 'image/jpeg'):
                    with Image.open(io.BytesIO(data)) as image:
                        width, height, form = image.width, image.height, image.format
                        if width <= 0 or height <= 0 or width * height > 16 * 1024 * 1024:
                            raise ValueError('Decoded image exceeds fixture policy')
                        image.verify()
                    decoded.append({'part': name, 'width': width, 'height': height, 'format': form, 'sha256': digest(data)})
            result = {'schema': 1, 'passed': True, 'decodedMedia': decoded, 'method': 'independent-pillow-image-decode'}
        else:
            result = Package(args.file).report() if args.command == 'inventory' else verify(args.before, args.after, json.loads(pathlib.Path(args.expectation).read_text()))
    except Exception as error:
        result = {'schema': 1, 'passed': False, 'failures': [type(error).__name__ + ': ' + str(error)]}
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if result.get('passed', True) else 1


if __name__ == '__main__':
    sys.exit(main())
