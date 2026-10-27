"""Independent streaming CSV/XLSX qualification; cell values stay in private targets.

audit --xlsx FILE --csv FILE --targets PRIVATE.json --report REPORT.json
verify --source FILE --saved FILE --edits PRIVATE.json --report REPORT.json
       [--allow-calculation-metadata]

Uses Python stdlib only. Shared strings are indexed on temporary disk; worksheet
rows are removed from the XML tree as they are consumed. No application reader,
writer, calculation engine or generated answer key participates in the oracle.
"""
import argparse
import collections
import csv
import decimal
import functools
import hashlib
import itertools
import json
import os
import pathlib
import posixpath
import re
import sqlite3
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile

X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NUMERIC = re.compile(r'^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$')
ESCAPE = re.compile(r'_x([0-9A-Fa-f]{4})_')


def require(condition, message):
    if not condition:
        raise ValueError(message)


def stream_hash(stream):
    result = hashlib.sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b''):
        result.update(block)
    return result.hexdigest()


def file_hash(file):
    with open(file, 'rb') as stream:
        return stream_hash(stream)


def identity(file):
    return {'path': str(pathlib.Path(file).resolve()), 'bytes': os.path.getsize(file), 'sha256': file_hash(file)}


def write_json(file, value, private=False):
    destination = pathlib.Path(file)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(destination, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600 if private else 0o644)
    if private:
        os.fchmod(fd, 0o600)
    with os.fdopen(fd, 'w', encoding='utf-8') as stream:
        json.dump(value, stream, indent=2, ensure_ascii=False)
        stream.write('\n')


def decode_text(value):
    value = ESCAPE.sub(lambda match: chr(int(match[1], 16)), value)
    return value.encode('utf-16-le', 'surrogatepass').decode('utf-16-le')


def text_content(node):
    # Phonetic annotations are not part of a cell's displayed literal text.
    return ''.join(decode_text(child.text or '') for child in node.iter(X + 't')
                   if not any(child is nested for phonetic in node.iter(X + 'rPh') for nested in phonetic.iter()))


def semantics(node):
    text = node.text or ''
    if len(node) and not text.strip():
        text = ''
    return [node.tag, sorted(node.attrib.items()), text, [semantics(child) for child in node]]


def column_index(address):
    match = re.fullmatch(r'([A-Z]+)([1-9]\d*)', address)
    require(match is not None, 'Malformed cell address')
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    return column - 1, int(match[2])


def column_name(index):
    result = ''
    index += 1
    while index:
        index, part = divmod(index - 1, 26)
        result = chr(65 + part) + result
    return result


class Workbook:
    def __init__(self, file):
        self.zip = zipfile.ZipFile(file)
        names = self.zip.namelist()
        require(len(names) == len(set(names)), 'Duplicate package members are not qualified')
        self.workbook = ET.fromstring(self.zip.read('xl/workbook.xml'))
        relationships = ET.fromstring(self.zip.read('xl/_rels/workbook.xml.rels'))
        targets = {item.get('Id'): item.get('Target') for item in relationships}
        self.sheets = {}
        for sheet in self.workbook.find(X + 'sheets'):
            target = targets[sheet.get(R + 'id')]
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
            self.sheets[sheet.get('name')] = part
        self.temporary = tempfile.TemporaryDirectory(prefix='xlsx-independent-strings-')
        self.db = sqlite3.connect(pathlib.Path(self.temporary.name) / 'strings.sqlite')
        self.db.execute('PRAGMA journal_mode=OFF')
        self.db.execute('CREATE TABLE strings (id INTEGER PRIMARY KEY, value TEXT NOT NULL)')
        self.shared_string_count = 0
        if 'xl/sharedStrings.xml' in names:
            with self.zip.open('xl/sharedStrings.xml') as stream:
                iterator = ET.iterparse(stream, events=('start', 'end'))
                _, root = next(iterator)
                batch = []
                for event, element in iterator:
                    if event == 'end' and element.tag == X + 'si':
                        batch.append((self.shared_string_count, text_content(element)))
                        self.shared_string_count += 1
                        root.remove(element)
                        if len(batch) == 1000:
                            self.db.executemany('INSERT INTO strings VALUES (?,?)', batch)
                            batch.clear()
                self.db.executemany('INSERT INTO strings VALUES (?,?)', batch)
                self.db.commit()
        self.lookup = functools.lru_cache(maxsize=512)(self._lookup)
        self.metadata = {}

    def _lookup(self, index):
        row = self.db.execute('SELECT value FROM strings WHERE id=?', (index,)).fetchone()
        require(row is not None, 'Missing shared string')
        return row[0]

    def rows(self, sheet):
        part = self.sheets[sheet]
        with self.zip.open(part) as stream:
            iterator = ET.iterparse(stream, events=('start', 'end'))
            _, root = next(iterator)
            data = None
            previous = 0
            for event, element in iterator:
                if event == 'start' and element.tag == X + 'sheetData':
                    data = element
                elif event == 'end' and element.tag == X + 'row':
                    number = int(element.get('r'))
                    require(number > previous, 'Worksheet rows are not ordered uniquely')
                    previous = number
                    cells = {}
                    for cell in element.findall(X + 'c'):
                        address = cell.get('r')
                        column, row = column_index(address)
                        require(row == number and column not in cells, 'Duplicate or mismatched cell coordinate')
                        kind, raw = cell.get('t', 'n'), cell.findtext(X + 'v')
                        formula = cell.find(X + 'f')
                        if kind == 's':
                            value, value_type = self.lookup(int(raw)), 'text'
                        elif kind == 'inlineStr':
                            inline = cell.find(X + 'is')
                            value, value_type = text_content(inline) if inline is not None else '', 'text'
                        elif kind == 'str':
                            value, value_type = decode_text(raw or ''), 'text'
                        elif raw is None:
                            value, value_type = None, 'blank'
                        elif kind == 'b':
                            require(raw in ('0', '1'), 'Invalid boolean cell')
                            value, value_type = raw == '1', 'boolean'
                        elif kind == 'n':
                            require(NUMERIC.fullmatch(raw) is not None, 'Invalid numeric cell')
                            value, value_type = raw, 'number'
                        else:
                            value, value_type = raw, kind
                        cells[column] = {'type': value_type, 'value': value,
                            'formula': semantics(formula) if formula is not None else None,
                            'attributes': sorted((key, val) for key, val in cell.attrib.items() if key not in ('r', 't')),
                            'extra': [semantics(child) for child in cell if child.tag not in (X + 'f', X + 'v', X + 'is')]}
                    yield number, sorted(element.attrib.items()), cells
                    data.remove(element)
            self.metadata[sheet] = semantics(root)

    def close(self):
        self.lookup.cache_clear()
        self.db.close()
        self.zip.close()
        self.temporary.cleanup()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


def csv_equal(cell, value):
    if cell is None or cell['type'] == 'blank':
        return value == ''
    if cell['formula'] is not None:
        return False
    if cell['type'] == 'number':
        return bool(NUMERIC.fullmatch(value)) and decimal.Decimal(cell['value']) == decimal.Decimal(value)
    if cell['type'] == 'boolean':
        return value.lower() == ('true' if cell['value'] else 'false')
    return cell['type'] == 'text' and cell['value'] == value


def exact_json_number(value):
    number = decimal.Decimal(value)
    if number == number.to_integral_value():
        integer = int(number)
        require(abs(integer) <= 9007199254740991, 'Selected value is not an exact JavaScript integer')
        return integer
    result = float(number)
    require(decimal.Decimal(str(result)) == number, 'Selected value loses numeric precision')
    return result


def audit(args):
    before = {'xlsx': identity(args.xlsx), 'csv': identity(args.csv)}
    csv_types, lengths = collections.Counter(), collections.Counter()
    with tempfile.TemporaryDirectory(prefix='xlsx-independent-keys-') as temporary:
        keys = sqlite3.connect(pathlib.Path(temporary) / 'keys.sqlite')
        keys.execute('CREATE TABLE ids (value TEXT PRIMARY KEY, count INTEGER NOT NULL)')
        with open(args.csv, encoding='utf-8-sig', newline='') as stream:
            reader = csv.reader(stream)
            header = next(reader)
            require('id' in header and 'price' in header, 'Sample needs id and price columns')
            key_column, price_column = header.index('id'), header.index('price')
            csv_rows = 1
            lengths[len(header)] += 1
            for row in reader:
                csv_rows += 1
                lengths[len(row)] += 1
                require(len(row) == len(header), f'CSV column count differs at row {csv_rows}')
                keys.execute('INSERT INTO ids VALUES (?,1) ON CONFLICT(value) DO UPDATE SET count=count+1', (row[key_column],))
                for value in row:
                    csv_types['blank' if value == '' else 'numeric-lexeme' if NUMERIC.fullmatch(value) else 'text-lexeme'] += 1
        keys.commit()
        desired = {'first': 2, 'middle': 1 + (csv_rows // 2), 'last': csv_rows}
        selected, mismatches, private_examples = {}, [], []
        mismatch_count, lexical_difference_count = 0, 0
        mismatch_digest = hashlib.sha256()
        xlsx_types, xlsx_cells, formula_cells, xlsx_rows = collections.Counter(), 0, 0, 0
        with Workbook(args.xlsx) as book, open(args.csv, encoding='utf-8-sig', newline='') as stream:
            require(len(book.sheets) == 1, 'CSV equivalence requires one explicitly selected sheet')
            sheet = next(iter(book.sheets))
            for index, (csv_row, xlsx_row) in enumerate(itertools.zip_longest(csv.reader(stream), book.rows(sheet)), 1):
                require(csv_row is not None and xlsx_row is not None, 'CSV/XLSX row counts differ')
                number, _, cells = xlsx_row
                require(number == index, f'Worksheet has a missing row at {index}')
                xlsx_rows += 1
                xlsx_cells += len(cells)
                for cell in cells.values():
                    xlsx_types[cell['type']] += 1
                    formula_cells += cell['formula'] is not None
                if set(cells) - set(range(len(header))):
                    raise ValueError(f'XLSX contains cells beyond CSV columns at row {index}')
                for column, value in enumerate(csv_row):
                    cell = cells.get(column)
                    if cell and cell['value'] is not None and str(cell['value']) != value:
                        lexical_difference_count += 1
                    if not csv_equal(cell, value):
                        mismatch_count += 1
                        actual = None if cell is None else cell['value']
                        pair = json.dumps([index, column, value, actual], ensure_ascii=False)
                        mismatch_digest.update(pair.encode('utf-8'))
                        mismatch_digest.update(b'\n')
                        if len(mismatches) < 20:
                            mismatches.append({'row': index, 'column': column_name(column), 'reason': 'typed value differs',
                                'csvValueSha256': hashlib.sha256(value.encode('utf-8')).hexdigest(),
                                'xlsxValueSha256': hashlib.sha256(str(actual).encode('utf-8')).hexdigest()})
                            private_examples.append({'cell': f'{column_name(column)}{index}', 'csv': value, 'xlsx': actual})
                if index == 1 or not csv_row[price_column] or not NUMERIC.fullmatch(csv_row[price_column]):
                    continue
                old = decimal.Decimal(csv_row[price_column])
                if (not old.is_finite() or cells[price_column]['type'] != 'number'
                        or not csv_equal(cells[price_column], csv_row[price_column])
                        or not csv_equal(cells.get(key_column), csv_row[key_column])):
                    continue
                for position, target_row in desired.items():
                    if position not in selected or abs(index - target_row) < abs(selected[position]['row'] - target_row):
                        selected[position] = {'position': position, 'sheet': sheet, 'row': index,
                            'cell': f'{column_name(price_column)}{index}', 'keyCell': f'{column_name(key_column)}{index}',
                            'key': exact_json_number(cells[key_column]['value']) if cells[key_column]['type'] == 'number' else cells[key_column]['value'],
                            'csvKeyText': csv_row[key_column], 'expectedValue': exact_json_number(str(old)), 'value': exact_json_number(str(old + 1)),
                            'expectedRawValue': cells[price_column]['value'], 'csvExpectedText': csv_row[price_column],
                            'csvReplacementText': str(old + 1)}
            string_count = book.shared_string_count
            expanded = sum(info.file_size for info in book.zip.infolist())
            parts = len(book.zip.infolist())
        edits = list(selected.values())
        require(len(edits) == 3 and len({edit['cell'] for edit in edits}) == 3, 'Three distinct numeric price targets are required')
        for edit in edits:
            require(keys.execute('SELECT count FROM ids WHERE value=?', (edit['csvKeyText'],)).fetchone()[0] == 1, 'Selected business key is not unique')
        duplicates = keys.execute('SELECT count(*) FROM ids WHERE count>1').fetchone()[0]
        keys.close()
    require(all(file_hash(before[k]['path']) == before[k]['sha256'] for k in before), 'Source changed during audit')
    private = {'schema': 'workbook.xlsx.real-sample-targets.v1', 'source': before, 'sheet': sheet,
        'keyColumn': column_name(key_column), 'valueColumn': column_name(price_column), 'edits': edits,
        'expectedRule': 'Increase each of three independently selected original numeric prices by exactly 1; all other cell values, formulas and metadata are unchanged.'}
    write_json(args.targets, private, private=True)
    examples = pathlib.Path(args.targets).with_name('source-difference-examples.json')
    write_json(examples, private_examples, private=True)
    return {'passed': mismatch_count == 0, 'auditCompleted': True, 'source': before, 'rowsIncludingHeader': xlsx_rows, 'dataRows': xlsx_rows - 1,
        'columns': len(header), 'headers': header, 'csvRowWidths': dict(lengths), 'csvLexicalTypesExcludingHeader': dict(csv_types),
        'xlsxStoredCellTypes': dict(xlsx_types), 'xlsxStoredCells': xlsx_cells,
        'xlsxImplicitBlankCells': xlsx_rows * len(header) - xlsx_cells, 'formulaCells': formula_cells,
        'sharedStrings': string_count, 'packageMembers': parts, 'expandedBytes': expanded,
        'csvXlsxTypedEquivalence': mismatch_count == 0, 'typedMismatchCount': mismatch_count,
        'lexicalDifferenceCountIncludingNumericSpellings': lexical_difference_count,
        'typedMismatchDigest': mismatch_digest.hexdigest(), 'mismatches': mismatches, 'privateDifferenceExamples': str(examples.resolve()), 'duplicateIdCount': duplicates,
        'selectedTargetCount': len(edits), 'targetsQualified': True, 'privateTargets': str(pathlib.Path(args.targets).resolve()),
        'sourceUnchanged': True, 'qualification': 'Real price-by-ID literal updates; no quantity or existing total/formula column is present. Target values are kept only in the private targets file.'}


def verify_csv(source, saved, edits):
    changed = {edit['row']: edit for edit in edits}
    checked, blanks = 0, 0
    with open(source, encoding='utf-8-sig', newline='') as before, open(saved, encoding='utf-8-sig', newline='') as after:
        for row_number, (old, new) in enumerate(itertools.zip_longest(csv.reader(before), csv.reader(after)), 1):
            require(old is not None and new is not None, 'CSV row count changed')
            require(len(old) == len(new), f'CSV column count changed at row {row_number}')
            edit = changed.get(row_number)
            expected = list(old)
            if edit:
                column, _ = column_index(edit['cell'])
                require(old[column] == edit['csvExpectedText'], f'CSV expected old value differs at row {row_number}')
                require(NUMERIC.fullmatch(new[column]) is not None and decimal.Decimal(new[column]) == decimal.Decimal(str(edit['value'])), f'CSV new value differs at row {row_number}')
                expected[column] = new[column]
            require(expected == new, f'Unexpected CSV data change at row {row_number}')
            checked += len(old)
            blanks += sum(value == '' for value in old)
    require(all(edit['row'] <= row_number for edit in edits), 'An expected target was not visited')
    return {'passed': True, 'rowsIncludingHeader': row_number, 'cellsChecked': checked, 'originalBlankFields': blanks, 'intendedEdits': len(edits), 'allOtherDataUnchanged': True}


def verify_xlsx(source, saved, edits, allow_calculation_metadata):
    changes = {(edit['sheet'], edit['cell']): edit for edit in edits}
    visited, row_count, cell_count = set(), 0, 0
    with Workbook(source) as before, Workbook(saved) as after:
        require(before.sheets == after.sheets, 'Sheet identities or relationship targets changed')
        old_parts = {name for name in before.zip.namelist() if not name.endswith('/')}
        new_parts = {name for name in after.zip.namelist() if not name.endswith('/')}
        require(old_parts == new_parts, 'Package members were added or removed')
        allowed_sheets = {before.sheets[edit['sheet']] for edit in edits}
        changed_parts, preserved_parts = [], []
        for name in sorted(old_parts):
            with before.zip.open(name) as a, after.zip.open(name) as b:
                identical = stream_hash(a) == stream_hash(b)
            if identical:
                preserved_parts.append(name)
            else:
                changed_parts.append(name)
                require(name in allowed_sheets or (name == 'xl/workbook.xml' and allow_calculation_metadata), f'Unqualified changed package member: {name}')
        calc_before = [dict(node.attrib) for node in before.workbook.findall(X + 'calcPr')]
        calc_after = [dict(node.attrib) for node in after.workbook.findall(X + 'calcPr')]
        if allow_calculation_metadata:
            for root in [before.workbook, after.workbook]:
                for node in root.findall(X + 'calcPr'):
                    root.remove(node)
        require(semantics(before.workbook) == semantics(after.workbook), 'Workbook metadata changed beyond the declared calculation metadata')
        for sheet in before.sheets:
            for old, new in itertools.zip_longest(before.rows(sheet), after.rows(sheet)):
                require(old is not None and new is not None, 'Worksheet row count changed')
                require(old[:2] == new[:2], f'Worksheet row attributes changed at {sheet} row {old[0]}')
                number, _, cells = old
                row_count += 1
                require(cells.keys() == new[2].keys(), f'Cell coordinates changed at {sheet} row {number}')
                for column, original in cells.items():
                    cell_count += 1
                    address = f'{column_name(column)}{number}'
                    edit = changes.get((sheet, address))
                    actual = new[2][column]
                    if edit:
                        require(original['type'] == actual['type'] == 'number', f'Price target type changed at {address}')
                        require(original['value'] == edit['expectedRawValue'], f'Expected original numeric storage differs at {address}')
                        require(decimal.Decimal(actual['value']) == decimal.Decimal(str(edit['value'])), f'Saved price differs at {address}')
                        require({k: v for k, v in original.items() if k != 'value'} == {k: v for k, v in actual.items() if k != 'value'}, f'Price target metadata changed at {address}')
                        visited.add((sheet, address))
                    else:
                        require(original == actual, f'Unexpected cell semantic change at {sheet}!{address}')
            require(before.metadata[sheet] == after.metadata[sheet], f'Non-cell worksheet metadata changed in {sheet}')
        require(visited == changes.keys(), 'Expected price target was not visited')
    return {'passed': True, 'rowsChecked': row_count, 'cellsChecked': cell_count, 'intendedEdits': len(visited),
        'allOtherCellSemanticsUnchanged': True, 'worksheetMetadataUnchanged': True,
        'changedMembers': changed_parts, 'byteIdenticalMembers': preserved_parts,
        'declaredCalculationMetadata': {'allowed': allow_calculation_metadata, 'before': calc_before, 'after': calc_after}}


def verify(args):
    specification = json.loads(pathlib.Path(args.edits).read_text())
    source, saved = identity(args.source), identity(args.saved)
    kind = pathlib.Path(args.source).suffix.lower().lstrip('.')
    require(kind in ('csv', 'xlsx'), 'Only CSV and XLSX are supported')
    require(source['sha256'] == specification['source'][kind]['sha256'], 'Independent target fixture hash does not match source')
    edits = specification['edits']
    result = verify_csv(args.source, args.saved, edits) if kind == 'csv' else verify_xlsx(args.source, args.saved, edits, args.allow_calculation_metadata)
    require(file_hash(args.source) == source['sha256'] and file_hash(args.saved) == saved['sha256'], 'File changed during verification')
    return {**result, 'source': source, 'saved': saved, 'sourceUnchangedDuringVerification': True, 'targetSpecificationSha256': file_hash(args.edits)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    inventory = sub.add_parser('audit')
    for name in ('xlsx', 'csv', 'targets', 'report'):
        inventory.add_argument('--' + name, required=True)
    check = sub.add_parser('verify')
    for name in ('source', 'saved', 'edits', 'report'):
        check.add_argument('--' + name, required=True)
    check.add_argument('--allow-calculation-metadata', action='store_true')
    args = parser.parse_args()
    start = time.monotonic()
    report = {'schema': 'workbook.xlsx.independent-real-sample.v1', 'command': args.command,
        'verifierSha256': file_hash(__file__), 'procedure': vars(args),
        'method': 'Streaming Python stdlib CSV/XML and exact Decimal values; temporary disk-backed shared strings. No application implementation or full worksheet XML tree.'}
    try:
        report.update(audit(args) if args.command == 'audit' else verify(args))
    except Exception as error:
        report.update(passed=False, error=str(error))
    report['elapsedSeconds'] = time.monotonic() - start
    write_json(args.report, report)
    print(json.dumps({'passed': report.get('passed', False), 'report': str(pathlib.Path(args.report).resolve()), 'elapsedSeconds': report['elapsedSeconds'], **({'error': report['error']} if 'error' in report else {})}))
    return 0 if report.get('passed') else 1


if __name__ == '__main__':
    raise SystemExit(main())
