import { DOMParser } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
/** Independent OOXML/CSV oracle shared by component and host integration tests. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
const hash = (file) => createHash('sha256').update(fs.readFileSync(file)).digest('hex');
export function checkXlsxPackage(source, saved, scenario) {
  if (scenario.realTargets) {
    const report = path.join(path.dirname(saved), `.${path.basename(saved)}-independent.json`);
    const checked = spawnSync(
      'python3',
      [
        fileURLToPath(new URL('./xlsx-real-sample.py', import.meta.url)),
        'verify',
        '--source',
        source,
        '--saved',
        saved,
        '--edits',
        scenario.realTargets,
        '--report',
        report,
        ...(scenario.format === 'csv' ? [] : ['--allow-calculation-metadata']),
      ],
      { encoding: 'utf8', timeout: 180000 }
    );
    assert.equal(checked.status, 0, checked.stderr || 'Independent real-file verification failed');
    const result = JSON.parse(fs.readFileSync(report));
    assert(result.passed);
    return result;
  }
  if (scenario.format === 'csv') {
    assert.deepEqual(
      fs.readFileSync(saved),
      Buffer.from(scenario.expectedBytes),
      'CSV must match the complete independently specified expected bytes'
    );
    return {
      passed: true,
      method: 'complete-literal-byte-oracle',
      sourceSha256: hash(source),
      outputSha256: hash(saved),
      utf8BomPreserved: true,
      quotedAndMultilineTextPreserved: true,
      allUnchangedFieldsPreserved: true,
      formulasRemainLiteral: true,
    };
  }
  const result = spawnSync('python3', [fileURLToPath(new URL('./xlsx-package-check.py', import.meta.url))], {
    input: JSON.stringify({ ...scenario, source, saved }),
    encoding: 'utf8',
    timeout: 30000,
    maxBuffer: 8 * 1024 * 1024,
  });
  assert(result.stdout, result.stderr || String(result.error));
  const report = JSON.parse(result.stdout);
  assert(report.passed, JSON.stringify(report.failures));
  return report;
}

/** Preserve members exactly, permitting only closed derived calculation/extent metadata.
 * Authored XML, styles, links and opaque payloads remain fully compared. */
export function assertRetainedPart(name, before, after) {
  if (Buffer.from(before).equals(Buffer.from(after))) return;
  assert(name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]+\.xml$/.test(name), name);
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const a = new DOMParser().parseFromString(strFromU8(before), 'application/xml'),
    b = new DOMParser().parseFromString(strFromU8(after), 'application/xml');
  const direct = (d, tag) =>
    Array.from(d.documentElement.childNodes).filter(
      (n) => n.nodeType === 1 && n.namespaceURI === ns && n.localName === tag
    );
  const attrs = (n, bindings = true) =>
    Object.fromEntries(
      Array.from(n.attributes)
        .filter((a) => bindings || a.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
        .map((a) => [a.name, a.value])
    );
  const closed = (n) => assert.equal(n.childNodes.length, 0, 'Derived metadata cannot own child payloads');
  if (name === 'xl/workbook.xml') {
    const old = direct(a, 'calcPr'),
      next = direct(b, 'calcPr');
    assert(old.length <= 1 && next.length === 1, 'Unique calculation properties');
    for (const n of [...old, ...next]) closed(n);
    const original = old[0] ? attrs(old[0], false) : {},
      result = attrs(next[0], false),
      allowed = new Set(['calcMode', 'calcOnSave', 'fullCalcOnLoad', 'forceFullCalc']);
    for (const key of new Set([...Object.keys(original), ...Object.keys(result)]))
      if (original[key] !== result[key]) {
        assert(allowed.has(key), 'Unowned calcPr attribute: ' + key);
        assert(
          key === 'calcMode' ? ['auto', 'manual'].includes(result[key]) : ['0', '1'].includes(result[key]),
          'Invalid calculation property: ' + key
        );
      }
    for (const n of [...old, ...next]) n.parentNode.removeChild(n);
  } else {
    const old = direct(a, 'dimension'),
      next = direct(b, 'dimension');
    assert(old.length <= 1 && next.length === 1, 'Unique derived dimension');
    for (const n of [...old, ...next]) {
      closed(n);
      assert.deepEqual(Object.keys(attrs(n, false)), ['ref'], 'Opaque dimensions remain protected');
      assert(/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(n.getAttribute('ref')));
    }
    for (const n of [...old, ...next]) n.parentNode.removeChild(n);
  }
  const tree = (n) => [
    n.nodeType,
    n.namespaceURI ?? null,
    n.localName ?? n.nodeName,
    n.nodeType === 1 ? Object.entries(attrs(n)).sort() : n.nodeValue,
    Array.from(n.childNodes ?? []).map(tree),
  ];
  assert.deepEqual(tree(a.documentElement), tree(b.documentElement), name + ' authored XML must remain');
}
