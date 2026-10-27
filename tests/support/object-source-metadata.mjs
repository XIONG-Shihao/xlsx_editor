/** Independently complete the hand-authored object fixtures' ordinary envelopes.
 * Their original sparse XML intentionally omitted dimensions/calculation flags.
 * Real object preservation checks compare this declared input, never a product
 * result, so initialization bookkeeping cannot hide later object/data damage. */
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function objectSourceMetadata(bytes) {
  const parts = unzipSync(bytes);
  for (const [file, data] of Object.entries(parts)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(file)) continue;
    let raw = strFromU8(data);
    const doc = new DOMParser().parseFromString(raw, 'application/xml');
    if (doc.documentElement.namespaceURI !== S) continue;
    if (doc.getElementsByTagNameNS(S, 'dimension').length) continue;
    const refs = Array.from(doc.getElementsByTagNameNS(S, 'c')).map((node) => node.getAttribute('r'));
    let firstRow = 1048576,
      lastRow = 1,
      firstCol = 16384,
      lastCol = 1;
    for (const ref of refs) {
      const m = /^([A-Z]{1,3})([1-9]\d*)$/.exec(ref);
      assert(m);
      const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
        row = Number(m[2]);
      firstCol = Math.min(firstCol, col);
      lastCol = Math.max(lastCol, col);
      firstRow = Math.min(firstRow, row);
      lastRow = Math.max(lastRow, row);
    }
    const name = (col, row) => {
        let letters = '';
        for (let n = col; n > 0; n = Math.floor((n - 1) / 26))
          letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
        return letters + row;
      },
      first = refs.length ? name(firstCol, firstRow) : 'A1',
      last = refs.length ? name(lastCol, lastRow) : 'A1';
    raw = raw.replace(
      /(<(?:\w+:)?worksheet\b[^>]*>)/,
      '$1<dimension ref="' + (first === last ? first : first + ':' + last) + '" xmlns="' + S + '"/>'
    );
    parts[file] = strToU8(raw);
  }
  const wb = strFromU8(parts['xl/workbook.xml']);
  if (!new DOMParser().parseFromString(wb, 'application/xml').getElementsByTagNameNS(S, 'calcPr').length)
    parts['xl/workbook.xml'] = strToU8(
      wb.replace(
        /(<\/(?:\w+:)?workbook>)/,
        '<calcPr xmlns="' + S + '" fullCalcOnLoad="0" forceFullCalc="0" calcMode="auto" calcOnSave="1"/>$1'
      )
    );
  return Buffer.from(zipSync(parts));
}
