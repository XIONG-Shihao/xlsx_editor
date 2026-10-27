"""Generate uncached DEC2HEX inputs; independently extract a saved Excel oracle.

Only source formulas, literal inputs, styles and collision geometry are generated.
Expected values are read exclusively from the separately recalculated saved file.
"""
import argparse
import hashlib
import importlib.util
import json
import pathlib
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

support_path = pathlib.Path(__file__).with_name('dynamic-array-reference.py')
spec = importlib.util.spec_from_file_location('dec2hex_dynamic_support', support_path)
support = importlib.util.module_from_spec(spec)
spec.loader.exec_module(support)
NS = support.NS
FORMULAS = {}
LABELS = {}
ARRAY_MASTERS = set()
STYLES = {'E143': 1}
MERGES = ['F154:G154']
LITERALS = {
    'B1': 255, 'B2': 254, 'B3': '255', 'B4': True, 'B5': None,
    'B7': 15, 'B8': 16, 'E137': 'occupied', 'E143': None, 'E149': 0,
    'N1': 'FF', 'O1': 42, 'N2': '0', 'O2': 10,
    'N3': '7FFFFFFF', 'O3': 99,
}
ERROR_INPUTS = {'B6': '#N/A'}
# Independently observed Excel 16.113.3 spelling changes in this exact fixture.
# This is a closed address/formula map, not a general normalization rule.
FORMULA_NORMALIZATIONS = {
    'E106': {'authored': 'DEC2HEX(-0)', 'saved': 'DEC2HEX(0)'},
    'E109': {'authored': 'DEC2HEX(+1)', 'saved': 'DEC2HEX(1)'},
}


def add(row, label, formula, *, array=False):
    ref = f'E{row}'
    LABELS[f'A{row}'] = label
    FORMULAS[ref] = formula
    if array:
        ARRAY_MASTERS.add(ref)
    return ref


for row, value in zip([1, 4, 7, 10, 13, 16], [0, 1, 15, 16, 255, 2147483647], strict=True):
    ref = add(row, f'one-argument integer {value}', f'DEC2HEX({value})')
    FORMULAS[f'F{row}'] = f'LEN({ref})'
    FORMULAS[f'G{row}'] = f'ISTEXT({ref})'
    FORMULAS[f'H{row}'] = f'{ref}&":hex"'
    FORMULAS[f'I{row}'] = f'VLOOKUP({ref},$N$1:$O$3,2,FALSE)'

for row, label, formula in [
    (19, 'numeric reference', 'DEC2HEX($B$1)'),
    (22, 'numeric reference plus expression', 'DEC2HEX($B$2+1)'),
    (25, 'numeric expression', 'DEC2HEX(128+127)'),
    (28, 'negative number', 'DEC2HEX(-1)'),
    (31, 'fraction', 'DEC2HEX(1.5)'),
    (34, 'outside proposed native bound', 'DEC2HEX(2147483648)'),
    (37, 'Excel positive bound', 'DEC2HEX(549755813887)'),
    (40, 'above Excel positive bound', 'DEC2HEX(549755813888)'),
    (43, 'optional places padding', 'DEC2HEX(255,4)'),
    (46, 'optional places too short', 'DEC2HEX(255,1)'),
    (49, 'quoted numeric text', 'DEC2HEX("255")'),
    (52, 'referenced numeric text', 'DEC2HEX($B$3)'),
    (55, 'literal boolean', 'DEC2HEX(TRUE)'),
    (58, 'referenced boolean', 'DEC2HEX($B$4)'),
    (61, 'referenced blank', 'DEC2HEX($B$5)'),
    (64, 'referenced error', 'DEC2HEX($B$6)'),
    (67, 'formula error input', 'DEC2HEX(1/0)'),
    (70, 'caught formula error', 'IFERROR(DEC2HEX(1/0),"caught")'),
    (73, 'negative input is not a missing-subset error', 'IFERROR(DEC2HEX(-1),"incorrect-catch")'),
    (76, 'native LEN expression dependent', 'LEN(DEC2HEX(255))'),
    (79, 'native ISTEXT expression dependent', 'ISTEXT(DEC2HEX(255))'),
    (82, 'native concatenation expression dependent', 'DEC2HEX(255)&":hex"'),
    (85, 'compatibility VLOOKUP expression dependent', 'VLOOKUP(DEC2HEX($B$1),$N$1:$O$3,2,FALSE)'),
    (88, 'LEN reference dependent', 'LEN(E13)'),
    (91, 'ISTEXT reference dependent', 'ISTEXT(E13)'),
    (94, 'concatenation reference dependent', 'E13&":hex"'),
    (97, 'inactive fractional branch', 'IF(TRUE,DEC2HEX(255),DEC2HEX(1.5))'),
    (100, 'selected qualified branch', 'IF(FALSE,DEC2HEX(1.5),DEC2HEX(255))'),
    (103, 'qualified result through IFERROR', 'IFERROR(DEC2HEX(255),"caught")'),
    (106, 'negative zero', 'DEC2HEX(-0)'),
    (109, 'positive unary input', 'DEC2HEX(+1)'),
    (112, 'invalid text', 'DEC2HEX("not-a-number")'),
    (115, 'zero places', 'DEC2HEX(255,0)'),
    (118, 'static numeric defined name', 'DEC2HEX(InputNumber)'),
]:
    add(row, label, formula)

for row, label, formula in [
    (124, 'vertical numeric range with free destination', 'DEC2HEX($B$7:$B$8)'),
    (130, 'horizontal array with free destination', 'DEC2HEX({1,2})'),
    (136, 'vertical range with literal blocker', 'DEC2HEX($B$7:$B$8)'),
    (142, 'vertical range with styled blank destination', 'DEC2HEX($B$7:$B$8)'),
    (148, 'vertical range with zero blocker', 'DEC2HEX($B$7:$B$8)'),
    (154, 'horizontal array with merged destination', 'DEC2HEX({15,16})'),
    (160, 'horizontal padded array with free destination', 'DEC2HEX({255,16},4)'),
]:
    ref = add(row, label, formula, array=True)
    FORMULAS[f'I{row}'] = f'SUM(_xlfn.ANCHORARRAY({ref}))'
    ARRAY_MASTERS.add(f'I{row}')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def generate(path):
    # Reuse the exact style, XLDAPR and relationship envelope accepted by Excel
    # for the retained dynamic reference. Its worksheet is replaced completely.
    with tempfile.TemporaryDirectory(prefix='xlsx-dec2hex-envelope-') as directory:
        base = pathlib.Path(directory) / 'base.xlsx'
        support.generate(base)
        with zipfile.ZipFile(base) as package:
            parts = {name: package.read(name) for name in package.namelist()}
    rows = {}
    literals = {**LITERALS, **LABELS}
    for ref in sorted(set(FORMULAS) | set(literals) | set(ERROR_INPUTS), key=support.position):
        style = f' s="{STYLES[ref]}"' if ref in STYLES else ''
        if ref in FORMULAS:
            dynamic = f' cm="1"><f t="array" ref="{ref}">' if ref in ARRAY_MASTERS else '><f>'
            cell = f'<c r="{ref}"{style}{dynamic}{escape(FORMULAS[ref])}</f></c>'
        elif ref in ERROR_INPUTS:
            cell = f'<c r="{ref}" t="e"><v>{ERROR_INPUTS[ref]}</v></c>'
        else:
            value = literals[ref]
            if value is None:
                cell = f'<c r="{ref}"{style}/>'
            elif isinstance(value, bool):
                cell = f'<c r="{ref}" t="b"><v>{int(value)}</v></c>'
            elif isinstance(value, str):
                cell = f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
            else:
                cell = f'<c r="{ref}"><v>{value}</v></c>'
        rows.setdefault(support.position(ref)[0], []).append(cell)
    data = ''.join(f'<row r="{row}">' + ''.join(cells) + '</row>' for row, cells in rows.items())
    merges = ''.join(f'<mergeCell ref="{ref}"/>' for ref in MERGES)
    parts['xl/worksheets/sheet1.xml'] = (
        f'<worksheet xmlns="{NS}"><dimension ref="A1:O164"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<cols><col min="1" max="1" width="48" customWidth="1"/>'
        '<col min="2" max="15" width="18" customWidth="1"/></cols>'
        f'<sheetData>{data}</sheetData><mergeCells count="{len(MERGES)}">{merges}</mergeCells></worksheet>').encode()
    parts['xl/workbook.xml'] = parts['xl/workbook.xml'].decode().replace(
        '<calcPr ', '<definedNames><definedName name="InputNumber">Cases!$B$1</definedName></definedNames><calcPr ').encode()
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as package:
        for name, content in parts.items():
            info = zipfile.ZipInfo(name, (2026, 9, 30, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            package.writestr(info, content)


def extract(source, oracle, output, excel, method):
    if source.resolve() == oracle.resolve():
        raise ValueError('The external recalculated oracle must be a separate workbook')
    if output.resolve() in [source.resolve(), oracle.resolve()]:
        raise ValueError('Oracle metadata cannot replace either workbook')
    with zipfile.ZipFile(source) as package:
        cells = {cell.attrib['r']: cell for cell in ET.fromstring(package.read('xl/worksheets/sheet1.xml')).iter(f'{{{NS}}}c')}
    for ref, formula in FORMULAS.items():
        cell = cells.get(ref)
        if cell is None or cell.findtext(f'{{{NS}}}f') != formula or cell.find(f'{{{NS}}}v') is not None:
            raise ValueError(f'Source formula/cache is not the uncached generated template at {ref}')
    support.FORMULAS = FORMULAS
    with tempfile.TemporaryDirectory(prefix='xlsx-dec2hex-extract-') as directory:
        result = support.extract(source, oracle, pathlib.Path(directory) / 'pending.json', excel, method)
    for ref, formula in FORMULAS.items():
        translation = FORMULA_NORMALIZATIONS.get(ref)
        if translation is not None and formula != translation['authored']:
            raise ValueError(f'Normalization source no longer matches its qualified template at {ref}')
        saved_formula = translation['saved'] if translation else formula
        if result['cells'][ref].get('formula') != saved_formula:
            raise ValueError(f'Unqualified saved formula translation at {ref}')
    result.update(schema='workbook.xlsx.excel-dec2hex-oracle.v1', labels=LABELS,
                  sourceArrayMasters=sorted(ARRAY_MASTERS, key=support.position),
                  savedFormulas={ref: result['cells'][ref]['formula'] for ref in FORMULAS},
                  formulaNormalizations=FORMULA_NORMALIZATIONS)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2) + '\n')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    gen = commands.add_parser('generate')
    gen.add_argument('output', type=pathlib.Path)
    ext = commands.add_parser('extract')
    for name in ['source', 'oracle', 'output']:
        ext.add_argument(name, type=pathlib.Path)
    ext.add_argument('--excel', required=True)
    ext.add_argument('--method', required=True)
    args = parser.parse_args()
    if args.command == 'generate':
        generate(args.output)
        print(json.dumps({'source': str(args.output), 'formulas': len(FORMULAS), 'sha256': sha(args.output), 'uncached': True}))
    else:
        result = extract(args.source, args.oracle, args.output, args.excel, args.method)
        print(json.dumps({'oracle': str(args.output), 'formulas': len(FORMULAS), 'cells': len(result['cells'])}))
