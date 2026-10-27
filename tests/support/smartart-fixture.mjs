/** Hand-authored workbook envelopes plus repository sample DiagramML models. */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
import {
  smartartPackage,
  smartartAnchorXml,
  smartartEdgeIds,
  smartartSampleModel,
} from '../../core/smartart-model.mjs';
import { objectSourceMetadata } from './object-source-metadata.mjs';
export const SMARTART_IDS = [
  '11111111-1111-4111-8111-ffffffffffff',
  '22222222-2222-4222-8222-222222222222',
  '33333333-3333-4333-8333-333333333333',
];
export function smartartFixture(source = createWorkbook(), layouts = ['process', 'hierarchy', 'cycle']) {
  const parts = unzipSync(objectSourceMetadata(source)),
    drawing = 'xl/drawings/workbook-drawing-' + SMARTART_IDS[0] + '.xml',
    anchors = [],
    relations = [],
    types = [];
  layouts.forEach((layout, i) => {
    const id = SMARTART_IDS[i],
      anchor = {
        kind: 'oneCell',
        cell: ['E2', 'E22', 'Q2'][i],
        dxEmu: 0,
        dyEmu: 0,
        widthEmu: 5715000,
        heightEmu: 3429000,
      },
      graph = smartartPackage(id, smartartSampleModel(layout), anchor),
      edges = smartartEdgeIds(id);
    anchors.push(smartartAnchorXml(id, i + 1, anchor, 'Workbook ' + layout + ' sample'));
    for (const [path, text] of Object.entries(graph.parts)) parts[path] = strToU8(text);
    for (const [part, type] of Object.entries(graph.types))
      types.push('<Override PartName="/' + part + '" ContentType="' + type + '"/>');
    for (const [attr, role, type] of [
      ['dm', 'data', 'Data'],
      ['lo', 'layout', 'Layout'],
      ['qs', 'style', 'QuickStyle'],
      ['cs', 'colors', 'Colors'],
    ])
      relations.push(
        '<Relationship Id="' +
          edges[attr] +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagram' +
          type +
          '" Target="../diagrams/' +
          graph.names[role].split('/').at(-1) +
          '"/>'
      );
  });
  parts[drawing] = strToU8(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
      anchors.join('') +
      '</xdr:wsDr>'
  );
  parts[drawing.replace(/([^/]+)$/, '_rels/$1.rels')] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relations.join('') +
      '</Relationships>'
  );
  const sheet = 'xl/worksheets/sheet1.xml';
  parts[sheet] = strToU8(
    strFromU8(parts[sheet]).replace(
      '</worksheet>',
      '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="workbookSmartArtDrawing"/></worksheet>'
    )
  );
  const rels = 'xl/worksheets/_rels/sheet1.xml.rels',
    edge =
      '<Relationship Id="workbookSmartArtDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/' +
      drawing.split('/').at(-1) +
      '"/>';
  parts[rels] = strToU8(
    parts[rels]
      ? strFromU8(parts[rels]).replace('</Relationships>', edge + '</Relationships>')
      : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          edge +
          '</Relationships>'
  );
  types.push(
    '<Override PartName="/' + drawing + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace('</Types>', types.join('') + '</Types>')
  );
  parts['custom/smartart-opaque.bin'] = Uint8Array.from([2, 7, 128, 255]);
  return Buffer.from(zipSync(parts));
}
