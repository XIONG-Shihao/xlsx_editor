"""Uncached dynamic-array inputs and independent saved-Excel geometry extraction."""
import argparse
import datetime
import hashlib
import json
import pathlib
import platform
import re
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
FORMULAS = {
    'A1': '_xlfn.SEQUENCE(3,2)', 'D1': 'SUM(A1#)', 'D2': 'INDEX(A1#,3,2)',
    'A6': '_xlfn.SEQUENCE(2,2,10,2)', 'D6': 'SUM(A6#)',
    'A10': '_xlfn.SEQUENCE(1,3,-2,0.5)',
    'A13': 'TRANSPOSE({1,2,3})',
    'A18': '_xlfn._xlws.FILTER(F1:F5,G1:G5=1,"none")', 'D18': 'SUM(A18#)',
    'A24': '_xlfn.UNIQUE(F1:F5)', 'D24': 'SUM(A24#)',
    'A30': '_xlfn._xlws.SORT(F1:F5,1,-1)',
    'A38': '_xlfn.SEQUENCE(0)', 'A41': '_xlfn.SEQUENCE(2,2)',
    'A44': '_xlfn.SEQUENCE(-1)', 'A47': 'SUM(_xlfn.SEQUENCE(3,1))',
    'A50': 'IF(TRUE,_xlfn.SEQUENCE(2),0)', 'D50': 'SUM(A50#)',
    'A53': 'IF(FALSE,_xlfn.SEQUENCE(2),7)',
    'A56': '_xlfn.SEQUENCE(2,2)', 'A60': '_xlfn.SEQUENCE(2,2)',
    'A64': '_xlfn.SEQUENCE(2,2)', 'A69': '_xlfn.SEQUENCE(2,2)',
}
FORMULAS = {ref: re.sub(r'\b([A-Z]+[1-9]\d*)#', r'_xlfn.ANCHORARRAY(\1)', expr)
            for ref, expr in FORMULAS.items()}
LITERALS = {'B42': 'occupied', 'B61': 0, 'B70': '=not a formula'}
for i, v in enumerate([1, 2, 2, 4, 5], 1):
    LITERALS[f'F{i}'] = v
for i, v in enumerate([1, 0, 1, 0, 1], 1):
    LITERALS[f'G{i}'] = v


def position(ref):
    m = re.fullmatch(r'([A-Z]+)([1-9]\d*)', ref)
    column = 0
    for char in m[1]:
        column = column * 26 + ord(char) - 64
    return int(m[2]), column


def generate(path):
    rows = {}
    for ref in sorted(set(FORMULAS) | set(LITERALS) | {'B56'}, key=position):
        row = position(ref)[0]
        if ref in FORMULAS:
            data = f'<c r="{ref}" cm="1"><f t="array" ref="{ref}">{escape(FORMULAS[ref])}</f></c>'
        elif ref == 'B56':
            data = f'<c r="{ref}" s="1"/>'
        else:
            value = LITERALS[ref]
            data = (f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
                    if isinstance(value, str) else f'<c r="{ref}"><v>{value}</v></c>')
        rows.setdefault(row, []).append(data)
    data = ''.join(f'<row r="{r}">' + ''.join(cells) + '</row>' for r, cells in rows.items())
    sheet = (f'<worksheet xmlns="{NS}"><dimension ref="A1:G72"/>'
             '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
             '<cols><col min="1" max="7" width="18" customWidth="1"/></cols>'
             f'<sheetData>{data}</sheetData><mergeCells count="1"><mergeCell ref="B65:C65"/></mergeCells></worksheet>')
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="0"/><bookViews><workbookView/></bookViews><sheets><sheet name="Cases" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': sheet,
        'xl/metadata.xml': f'<metadata xmlns="{NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>',
        'xl/styles.xml': f'<styleSheet xmlns="{NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    }
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, (2026, 9, 30, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode())


def extract(source, oracle, output, excel, method):
    with zipfile.ZipFile(oracle) as z:
        sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        metadata = {n: z.read(n).decode() for n in z.namelist() if n == 'xl/metadata.xml' or n.startswith('xl/richData/') and n.endswith('.xml')}
    rich = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
    modern_errors = []
    if 'xl/richData/rdrichvalue.xml' in metadata:
        structures = list(ET.fromstring(metadata['xl/richData/rdrichvaluestructure.xml']))
        values = list(ET.fromstring(metadata['xl/richData/rdrichvalue.xml']))
        for item in values:
            structure = structures[int(item.attrib['s'])]
            fields = {k.attrib['n']: int(v.text) for k, v in zip(structure, item, strict=True)}
            modern_errors.append(fields if structure.attrib['t'] == '_error' else None)
        meta = ET.fromstring(metadata['xl/metadata.xml'])
        future = meta.find(f'{{{NS}}}futureMetadata[@name="XLRICHVALUE"]')
        metadata_values = meta.find(f'{{{NS}}}valueMetadata')
        if future is None or metadata_values is None:
            raise ValueError('Rich error has no value metadata mapping')
    cells = {}
    for cell in sheet.iter(f'{{{NS}}}c'):
        ref = cell.attrib['r']
        formula = cell.find(f'{{{NS}}}f')
        raw = cell.findtext(f'{{{NS}}}v')
        typ = cell.attrib.get('t', 'n')
        if typ == 's':
            value, datatype = strings[int(raw)], 'text'
        elif typ == 'inlineStr':
            value, datatype = ''.join(cell.find(f'{{{NS}}}is').itertext()), 'text'
        elif raw is None:
            value, datatype = None, 'blank'
        else:
            value = float(raw) if typ == 'n' else raw == '1' if typ == 'b' else raw
            datatype = {'n': 'number', 'b': 'boolean', 'e': 'error', 'str': 'text'}[typ]
        cells[ref] = {'expected': value, 'dataType': datatype, 'rawValue': raw,
                      'cellAttributes': cell.attrib,
                      **({'formula': formula.text, 'formulaAttributes': formula.attrib} if formula is not None else {})}
        if 'vm' in cell.attrib:
            record = list(metadata_values)[int(cell.attrib['vm']) - 1]
            rc = record.find(f'{{{NS}}}rc')
            block = list(future)[int(rc.attrib['v'])]
            entry = next(block.iter(f'{{{rich}}}rvb'))
            details = modern_errors[int(entry.attrib['i'])]
            if details is None:
                raise ValueError(f'Unqualified rich value at {ref}')
            cells[ref]['richError'] = details
            cells[ref]['rawCachedValue'] = value
            # MS-XLSX2.3.6.1.3.7 identifies rich errorType8 as #SPILL!.
            # Unknown modern error codes stay explicit, never guessed from the
            # legacy #VALUE! cache. Code13 is retained as an unqualified label.
            cells[ref]['expected'] = '#SPILL!' if details['errorType'] == 8 else None
            cells[ref]['semanticStatus'] = 'qualified-rich-error' if details['errorType'] == 8 else 'unqualified-rich-error-label'
    for ref in FORMULAS:
        if ref not in cells or 'formula' not in cells[ref] or cells[ref]['rawValue'] is None:
            raise ValueError(f'External formula/cache missing at {ref}')
    result = {'schema': 'workbook.xlsx.excel-dynamic-array-oracle.v1', 'excel': excel,
              'sheet': 'Cases', 'source': source.name, 'oracleWorkbook': oracle.name,
              'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
              'oracleWorkbookSha256': hashlib.sha256(oracle.read_bytes()).hexdigest(),
              'capturedAt': datetime.datetime.now(datetime.timezone.utc).isoformat(),
              'platform': platform.platform(), 'method': method,
              'authoredFormulas': FORMULAS, 'cells': cells, 'metadata': metadata,
              'specificationSources': ['https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7a27b5fb-48a8-4a1c-b67a-2ca5edc07514', 'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1d44a793-c90a-47da-8943-08802dfad1fd'],
              'merges': [node.attrib['ref'] for node in sheet.iter(f'{{{NS}}}mergeCell')]}
    output.write_text(json.dumps(result, indent=2) + '\n')
    return result


if __name__ == '__main__':
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='command', required=True)
    gen = sub.add_parser('generate')
    gen.add_argument('output', type=pathlib.Path)
    ext = sub.add_parser('extract')
    for arg in ['source', 'oracle', 'output']:
        ext.add_argument(arg, type=pathlib.Path)
    ext.add_argument('--excel', required=True)
    ext.add_argument('--method', required=True)
    args = p.parse_args()
    if args.command == 'generate':
        generate(args.output)
        print(json.dumps({'source': str(args.output), 'formulas': len(FORMULAS)}))
    else:
        data = extract(args.source, args.oracle, args.output, args.excel, args.method)
        print(json.dumps({'oracle': str(args.output), 'formulas': len(FORMULAS), 'cells': len(data['cells'])}))
