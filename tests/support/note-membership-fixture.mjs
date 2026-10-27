/** Authored package facts; not an observed Excel Notes-selection reference. */
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
export const noteThreadId = '{12345678-1234-4234-8234-123456789ABC}';
export const notePersonId = '{23456789-2345-4345-8345-23456789ABCD}';
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const opc = 'http://schemas.openxmlformats.org/package/2006/relationships';
const thread = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const extensionRel = 'http://schemas.microsoft.com/office/2017/10/relationships';
export function noteMembershipFixture({
  mixed = true,
  protectedSheet = false,
  missingPerson = false,
  nonstandard = false,
} = {}) {
  const parts = unzipSync(createWorkbook());
  const commentPath = nonstandard ? 'xl/annotations/legacy-data.xml' : 'xl/comments1.xml';
  const threadedPath = nonstandard ? 'xl/annotations/discussion.xml' : 'xl/threadedComments/threadedComment1.xml';
  const personPath = nonstandard ? 'xl/annotations/people.xml' : 'xl/persons/person.xml';
  const vmlPath = 'xl/drawings/vmlDrawing1.vml';
  const annotations = [commentPath, vmlPath, ...(mixed ? [threadedPath, personPath] : [])];
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/' +
        commentPath +
        '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' +
        (mixed
          ? '<Override PartName="/' +
            threadedPath +
            '" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/' +
            personPath +
            '" ContentType="application/vnd.ms-excel.person+xml"/>'
          : '') +
        '</Types>'
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="' +
      main +
      '" xmlns:r="' +
      rel +
      '"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="4" hidden="1"/></sheetData>' +
      (protectedSheet ? '<sheetProtection sheet="1" objects="1" scenarios="1"/>' : '') +
      '<legacyDrawing r:id="vml"/></worksheet>'
  );
  parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    '<Relationships xmlns="' +
      opc +
      '"><Relationship Id="notes" Type="' +
      rel +
      '/comments" Target="/' +
      commentPath +
      '"/><Relationship Id="vml" Type="' +
      rel +
      '/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
      (mixed
        ? '<Relationship Id="thread" Type="' + extensionRel + '/threadedComment" Target="/' + threadedPath + '"/>'
        : '') +
      '</Relationships>'
  );
  parts[commentPath] = strToU8(
    '<comments xmlns="' +
      main +
      '"><authors><author>Legacy author</author>' +
      (mixed ? '<author>Thread author tc=' + noteThreadId + '</author>' : '') +
      '</authors><commentList>' +
      '<comment ref="D4" authorId="0"><text><r><rPr><b/></rPr><t>Note without a cell record</t></r></text></comment>' +
      '<comment ref="XFD1048576" authorId="0"><text><t/></text></comment>' +
      (mixed
        ? '<comment ref="B2" authorId="1"><text><t>This compatibility text does not define Notes membership.</t></text></comment>'
        : '') +
      '</commentList></comments>'
  );
  parts[vmlPath] = strToU8(
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>3</x:Row><x:Column>3</x:Column></x:ClientData></v:shape></xml>'
  );
  if (mixed) {
    parts[threadedPath] = strToU8(
      '<ThreadedComments xmlns="' +
        thread +
        '"><threadedComment ref="B2" dT="2026-10-04T00:00:00Z" personId="' +
        notePersonId +
        '" id="' +
        noteThreadId +
        '"><text>A threaded discussion</text></threadedComment></ThreadedComments>'
    );
    parts[personPath] = strToU8(
      '<personList xmlns="' +
        thread +
        '"><person displayName="Local author" id="' +
        notePersonId +
        '" userId="Local author" providerId="None"/></personList>'
    );
    if (missingPerson) delete parts[personPath];
    else
      parts['xl/_rels/workbook.xml.rels'] = strToU8(
        strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
          '</Relationships>',
          '<Relationship Id="people" Type="' + extensionRel + '/person" Target="/' + personPath + '"/></Relationships>'
        )
      );
  }
  return {
    bytes: zipSync(parts, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') }),
    annotations,
    notes: ['D4', 'XFD1048576'],
    threadedPlaceholders: mixed ? ['B2'] : [],
  };
}
