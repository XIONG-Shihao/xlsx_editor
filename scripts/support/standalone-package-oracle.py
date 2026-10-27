"""Reuse the independent package oracle and verify the saved Bold action."""
import json
import math
from pathlib import Path
import posixpath
import runpy
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile


def cache_checks(report, expected):
    """Check saved cache XML type and exact configured scalar, not approximate equality."""
    cells = {item['target']: item.get('actual') for item in report['checks']}
    results = []
    for target, key in expected.items():
        value, actual = key['value'], cells.get(target)
        passed, reason, actual_bits = False, 'Missing formula/cache cell', None
        if actual is not None and actual.get('formula') is not None:
            if isinstance(value, bool):
                passed = actual['type'] == 'b' and type(actual['value']) is bool and actual['value'] is value
                reason = 'Expected boolean cache t=b with exact boolean value'
            elif isinstance(value, str):
                passed = actual['type'] == 'str' and type(actual['value']) is str and actual['value'] == value
                reason = 'Expected string formula cache t=str with exact text'
            elif type(value) in (int, float) and math.isfinite(value):
                wanted = key.get('binary64Bits')
                configured = struct.pack('>d', float(value)).hex()
                # JSON preserves the separate bits field even when JavaScript's
                # number serializer changes -0 to 0. Other mismatches are invalid.
                valid = wanted == configured or value == 0 and wanted in ('0000000000000000', '8000000000000000')
                # The general inventory treats integer lexical text as an int,
                # which loses -0. Decode the numeric cache's own XML text for bits.
                try:
                    cached = float(actual['rawValue'])
                    numeric = actual['rawValue'] is not None and math.isfinite(cached)
                except (TypeError, ValueError, OverflowError):
                    numeric = False
                actual_bits = struct.pack('>d', cached).hex() if numeric else None
                passed = valid and actual['type'] == 'n' and numeric and actual_bits == wanted
                reason = 'Expected numeric cache t=n with exact configured binary64 bits'
            else:
                reason = 'Unsupported configured formula cache scalar'
        results.append({'target': target, 'passed': passed, 'expected': key, 'actual': actual, 'actualBinary64Bits': actual_bits, 'reason': reason})
    return results


def bold_check(saved, expected):
    x = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    p = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    with zipfile.ZipFile(saved) as package:
        workbook = ET.fromstring(package.read('xl/workbook.xml'))
        sheet = next(item for item in workbook.findall(x + 'sheets/' + x + 'sheet')
                     if item.get('name') == expected['sheet'])
        relationships = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
        link = next(item for item in relationships.findall(p + 'Relationship')
                    if item.get('Id') == sheet.get(r + 'id'))
        target = link.get('Target')
        part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        style = None
        with package.open(part) as stream:
            for _, element in ET.iterparse(stream, events=('end',)):
                if element.tag == x + 'c' and element.get('r') == expected['cell']:
                    style = int(element.get('s', '0'))
                    break
                if element.tag == x + 'row':
                    element.clear()
        if style is None:
            return {'passed': False, 'expected': expected, 'reason': 'Formatted cell missing'}
        styles = ET.fromstring(package.read('xl/styles.xml'))
        xf = styles.find(x + 'cellXfs')[style]
        font = styles.find(x + 'fonts')[int(xf.get('fontId', '0'))]
        bold = font.find(x + 'b')
        actual = bold is not None and bold.get('val', '1') not in ('0', 'false', 'False')
        return {'passed': actual == expected['value'], 'expected': expected, 'actual': actual, 'styleId': style}


def shared_region_checks(saved, regions):
    """Inspect every declared cache/formula/owner in bounded uniform regions."""
    x = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    p = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    results = []
    with zipfile.ZipFile(saved) as package:
        workbook = ET.fromstring(package.read('xl/workbook.xml'))
        relationships = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
        for expected in regions:
            first, last = expected['first'], expected['last']
            ordinary = expected.get('formulaKind') == 'ordinary'
            required = {'sheet', 'first', 'last', 'column', 'type', 'value', 'master', 'formula', 'reference'} | ({'formulaKind'} if ordinary else {'sharedId'})
            if set(expected) != required or (ordinary and expected['formulaKind'] != 'ordinary') or (not ordinary and (not isinstance(expected['sharedId'], str) or not expected['sharedId'].isdigit())):
                raise ValueError('Invalid formula region kind or fields')
            if not (type(first) is int and type(last) is int and 1 <= first <= last <= 1048576 and last-first < 100000 and expected['type'] == 'str' and expected['master'] == expected['column']+str(first) and expected['reference'] == expected['column']+str(first)+':'+expected['column']+str(last)):
                raise ValueError('Invalid bounded shared-region expectation')
            sheet = next(item for item in workbook.findall(x + 'sheets/' + x + 'sheet') if item.get('name') == expected['sheet'])
            link = next(item for item in relationships.findall(p + 'Relationship') if item.get('Id') == sheet.get(r + 'id'))
            if link.get('TargetMode') == 'External':
                raise ValueError('External worksheet relationship')
            target = link.get('Target')
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
            seen, failures = set(), []
            with package.open(part) as stream:
                for _, element in ET.iterparse(stream, events=('end',)):
                    if element.tag == x + 'c':
                        reference = element.get('r', '')
                        if reference.startswith(expected['column']) and reference[len(expected['column']):].isdigit():
                            row = int(reference[len(expected['column']):])
                            if first <= row <= last:
                                formula, caches = element.findall(x+'f'), element.findall(x+'v')
                                owner = {} if ordinary else {'t':'shared', 'si':expected['sharedId']}
                                wanted_formula = expected['formula'] if ordinary else None
                                if not ordinary and reference == expected['master']:
                                    owner['ref'] = expected['reference']
                                    wanted_formula = expected['formula']
                                passed = reference not in seen and element.get('t') == expected['type'] and len(caches) == 1 and caches[0].text == expected['value'] and len(formula) == 1 and formula[0].attrib == owner and (formula[0].text or None) == wanted_formula
                                seen.add(reference)
                                if not passed and len(failures) < 20:
                                    failures.append(reference)
                    elif element.tag == x+'row':
                        element.clear()
            passed = len(seen) == last-first+1 and not failures
            results.append({'passed':passed,'expected':expected,'checkedCells':len(seen),'failureExamples':failures,'method':'independent streaming XML, every typed cache/owner/formula'})
    return results


if __name__ == '__main__':
    request = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
    oracle = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'test' / 'xlsx-package-check.py'))
    report = oracle['check'](request)
    report['formulaCacheChecks'] = cache_checks(report, request.get('expectedFormulaCaches', {}))
    report['formulaRegionChecks'] = shared_region_checks(request['saved'], request.get('expectedFormulaRegions', []))
    for result in report['formulaRegionChecks']:
        if not result['passed']:
            report['failures'].append({'kind':'shared-formula-region','check':result})
            report['passed'] = False
    for result in report['formulaCacheChecks']:
        if not result['passed']:
            report['failures'].append({'kind': 'expected-saved-formula-cache', 'check': result})
            report['passed'] = False
    if request.get('expectedBold'):
        report['boldCheck'] = bold_check(request['saved'], request['expectedBold'])
        if not report['boldCheck']['passed']:
            report['failures'].append({'kind': 'expected-saved-bold', 'check': report['boldCheck']})
            report['passed'] = False
    print(json.dumps(report, ensure_ascii=False))
    sys.exit(0 if report['passed'] else 1)
