/** Source package membership. Fixtures are authored, not Excel selector baselines. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { indexedNoteMembership } from '../core/indexed-notes.mjs';
import { openNoteZipFile } from '../core/note-zip-reader.mjs';
import { noteMembershipFixture, noteThreadId, notePersonId } from './support/note-membership-fixture.mjs';
const SOURCE = { version: 7, sourceRevision: 'a'.repeat(64), sheetPath: 'xl/worksheets/sheet1.xml' },
  VIEW = { sheets: [{ name: 'Sheet1', path: SOURCE.sheetPath }] };
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  OPC = 'http://schemas.openxmlformats.org/package/2006/relationships',
  OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  THREAD = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const parts = (options = {}) => unzipSync(noteMembershipFixture(options).bytes);
const xml = (parts, name, update) => (parts[name] = strToU8(update(strFromU8(parts[name]))));
const encode = (parts) => zipSync(parts, { level: 6, mtime: new Date('2020-01-01T00:00:00Z') });
async function read(input, options = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'workbook-note-reader-')),
    file = path.join(dir, 'metadata.xlsx');
  try {
    fs.writeFileSync(file, input instanceof Uint8Array ? input : encode(input));
    return await indexedNoteMembership({ metadataFile: file, view: VIEW, sheet: 'Sheet1', source: SOURCE, ...options });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
function incomplete(result, code) {
  assert.equal(result.status, 'incomplete');
  assert.equal(result.notes, null);
  assert.equal(result.threadedPlaceholders, null);
  assert.ok(result.reasons.length);
  if (code) assert.equal(result.reasons[0].code, code);
}
const prefix = (source) => source.replace(/<(\/?)([A-Za-z][A-Za-z0-9]*)/g, '<$1n:$2').replace(/xmlns=/, 'xmlns:n=');
test('legacy notes include rich/empty text and note-only edge cells independent of protections and records', async () => {
  const result = await read(parts({ protectedSheet: true }));
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
  assert.deepEqual(result.threadedPlaceholders, ['B2']);
  assert.deepEqual(result.source, SOURCE);
  const source = parts({ mixed: false });
  delete source['xl/drawings/vmlDrawing1.vml'];
  assert.equal((await read(source)).status, 'complete', 'membership does not qualify the VML display graph');
  source['xl/drawings/vmlDrawing1.vml'] = new Uint8Array();
  assert.equal((await read(source)).status, 'complete', 'native metadata empty VML members do not mean source absence');
});
test('relative/rooted/nonstandard targets and namespace prefixes resolve the same graph', async () => {
  const source = parts({ nonstandard: true });
  xml(source, 'xl/worksheets/_rels/sheet1.xml.rels', (s) =>
    s
      .replace('Target="/xl/annotations/legacy-data.xml"', 'Target="../annotations/legacy-data.xml"')
      .replace('Target="/xl/annotations/discussion.xml"', 'Target="../annotations/discussion.xml"')
  );
  for (const name of [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/_rels/sheet1.xml.rels',
    'xl/annotations/legacy-data.xml',
    'xl/annotations/discussion.xml',
    'xl/annotations/people.xml',
  ])
    xml(source, name, prefix);
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
  assert.deepEqual(result.threadedPlaceholders, ['B2']);
});
test('compatibility identity is the author token; absent UID and differing thread ref are valid', async () => {
  const source = parts();
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) => s.replace('ref="B2"', 'ref="C3"'));
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.threadedPlaceholders, ['B2']);
  xml(source, 'xl/comments1.xml', (s) =>
    s.replace('ref="B2" authorId="1"', 'ref="B2" authorId="1" uid="' + noteThreadId + '"')
  );
  assert.equal((await read(source)).status, 'complete');
  xml(source, 'xl/comments1.xml', (s) =>
    s.replace(
      ' uid="' + noteThreadId + '"',
      ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="' +
        noteThreadId.toLowerCase() +
        '"'
    )
  );
  assert.equal((await read(source)).status, 'complete');
});
test('optional legacy UID mismatch and empty thread parts remain spec-supported membership profiles', async () => {
  const mismatch = parts();
  xml(mismatch, 'xl/comments1.xml', (s) =>
    s.replace('ref="B2" authorId="1"', 'ref="B2" authorId="1" uid="' + notePersonId + '"')
  );
  assert.deepEqual((await read(mismatch)).threadedPlaceholders, ['B2']);
  const empty = parts();
  xml(empty, 'xl/comments1.xml', (s) => s.replace(/<comment ref="B2"[^>]*>.*?<\/comment>/, ''));
  empty['xl/threadedComments/threadedComment1.xml'] = strToU8('<ThreadedComments xmlns="' + THREAD + '"/>');
  delete empty['xl/persons/person.xml'];
  xml(empty, '[Content_Types].xml', (s) => s.replace(/<Override PartName="\/xl\/persons\/person.xml"[^>]*\/>/, ''));
  xml(empty, 'xl/_rels/workbook.xml.rels', (s) => s.replace(/<Relationship Id="people"[^>]*\/>/, ''));
  const result = await read(empty);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
  assert.deepEqual(result.threadedPlaceholders, []);
});
test('ordinary wording never becomes a placeholder without the required author token', async () => {
  const source = parts({ mixed: false });
  xml(source, 'xl/comments1.xml', (s) =>
    s.replace('Note without a cell record', '[Threaded comment] compatibility warning')
  );
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
});
test('a qualified package with no annotations completes with an empty inventory', async () => {
  const source = parts({ mixed: false });
  for (const name of ['xl/comments1.xml', 'xl/worksheets/_rels/sheet1.xml.rels']) delete source[name];
  xml(source, '[Content_Types].xml', (s) => s.replace(/<Override PartName="\/xl\/comments1.xml"[^>]*\/>/, ''));
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, []);
  assert.deepEqual(result.threadedPlaceholders, []);
});
test('invalid legacy references and author indexes fail the whole inventory', async () => {
  for (const patch of ['ref="D4:E5"', 'ref="XFE1"', 'ref="A1048577"', 'ref="A01"', 'ref="Other!D4"']) {
    const source = parts();
    xml(source, 'xl/comments1.xml', (s) => s.replace('ref="D4"', patch));
    incomplete(await read(source), 'invalid-reference');
  }
  for (const patch of [
    'authorId="99"',
    'authorId="-1"',
    'authorId="0.5"',
    'authorId="4294967296"',
    'authorId=""',
    '',
  ]) {
    const source = parts();
    xml(source, 'xl/comments1.xml', (s) => s.replace('authorId="0"', patch));
    incomplete(await read(source), 'invalid-author');
  }
});
test('single-cell source aliases normalize before duplicate checks and unsignedInt source lexical forms are valid', async () => {
  const source = parts();
  xml(source, 'xl/comments1.xml', (s) =>
    s
      .replace('ref="D4" authorId="0"', 'ref="$d$4" authorId=" +0000 " shapeId="4294967295"')
      .replace('ref="B2" authorId="1"', 'ref="b2" authorId="0001"')
  );
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) => s.replace('ref="B2"', 'ref="$b$2"'));
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
  assert.deepEqual(result.threadedPlaceholders, ['B2']);
  xml(source, 'xl/comments1.xml', (s) => s.replace('ref="XFD1048576"', 'ref="D4"'));
  incomplete(await read(source), 'duplicate-annotation');
});
test('malformed, duplicate, missing and foreign roots or content types never return partial success', async () => {
  const cases = [
    ['xl/comments1.xml', (s) => s.replace('ref="XFD1048576"', 'ref="D4"'), 'duplicate-annotation'],
    ['xl/comments1.xml', (s) => s.replace('</commentList>', '<extLst/></commentList>'), 'foreign-graph'],
    ['xl/comments1.xml', (s) => s.replace(MAIN, 'urn:foreign'), 'wrong-root'],
    ['xl/comments1.xml', (s) => s.replace('</comments>', ''), 'malformed-xml'],
    ['xl/comments1.xml', (s) => 'bad' + s, 'malformed-xml'],
    ['xl/comments1.xml', (s) => s + 'bad', 'malformed-xml'],
    [
      'xl/comments1.xml',
      (s) => s.replace('<comments ', '<!DOCTYPE comments [<!ENTITY x "x">]><comments '),
      'unqualified-extension',
    ],
    [
      '[Content_Types].xml',
      (s) => s.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml', 'application/xml'),
      'wrong-content-type',
    ],
    ['xl/persons/person.xml', (s) => s.replaceAll('personList', 'foreign'), 'wrong-root'],
    [
      'xl/persons/person.xml',
      (s) => s.replace('</personList>', '<person id="' + notePersonId + '" displayName="Duplicate"/></personList>'),
      'duplicate-person',
    ],
    ['xl/threadedComments/threadedComment1.xml', (s) => s.replace(THREAD, 'urn:foreign'), 'wrong-root'],
    ['xl/worksheets/_rels/sheet1.xml.rels', (s) => s.replace('Id="notes"', 'Id="vml"'), 'duplicate-relationship'],
  ];
  for (const [name, update, code] of cases) {
    const source = parts();
    xml(source, name, update);
    incomplete(await read(source), code);
  }
  for (const name of ['xl/comments1.xml', 'xl/persons/person.xml', '[Content_Types].xml']) {
    const source = parts();
    delete source[name];
    incomplete(await read(source), 'missing-part');
  }
});
test('relationship content types qualify both defaults and overriding declarations', async () => {
  const correct = parts();
  xml(correct, '[Content_Types].xml', (s) =>
    s.replace(
      '</Types>',
      '<Override PartName="/xl/worksheets/_rels/sheet1.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'
    )
  );
  assert.equal((await read(correct)).status, 'complete');
  const wrongDefault = parts();
  xml(wrongDefault, '[Content_Types].xml', (s) =>
    s.replace('application/vnd.openxmlformats-package.relationships+xml', 'application/xml')
  );
  incomplete(await read(wrongDefault), 'wrong-content-type');
  const wrongOverride = parts();
  xml(wrongOverride, '[Content_Types].xml', (s) =>
    s.replace(
      '</Types>',
      '<Override PartName="/xl/worksheets/_rels/sheet1.xml.rels" ContentType="application/xml"/></Types>'
    )
  );
  incomplete(await read(wrongOverride), 'wrong-content-type');
  const empty = parts({ mixed: false });
  delete empty['xl/comments1.xml'];
  delete empty['xl/worksheets/_rels/sheet1.xml.rels'];
  xml(empty, '[Content_Types].xml', (s) =>
    s
      .replace(/<Override PartName="\/xl\/comments1.xml"[^>]*\/>/, '')
      .replace('application/vnd.openxmlformats-package.relationships+xml', 'application/xml')
  );
  incomplete(await read(empty), 'wrong-content-type');
});
test('qualified UID presence is distinct from absence even when its value is empty', async () => {
  const source = parts();
  xml(source, 'xl/comments1.xml', (s) =>
    s.replace(
      'ref="B2" authorId="1"',
      'ref="B2" authorId="1" xmlns:alias="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" alias:uid=""'
    )
  );
  incomplete(await read(source), 'invalid-thread-id');
  const valid = parts();
  xml(valid, 'xl/comments1.xml', (s) =>
    s.replace(
      'ref="B2" authorId="1"',
      'ref="B2" authorId="1" xmlns:alias="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" alias:uid="' +
        noteThreadId.toLowerCase() +
        '"'
    )
  );
  assert.deepEqual((await read(valid)).threadedPlaceholders, ['B2']);
});
test('the compatibility author marker is literal lowercase tc while GUID case can vary', async () => {
  const lowerGuid = parts();
  xml(lowerGuid, 'xl/comments1.xml', (s) => s.replace('tc=' + noteThreadId, 'tc=' + noteThreadId.toLowerCase()));
  assert.deepEqual((await read(lowerGuid)).threadedPlaceholders, ['B2']);
  for (const marker of ['TC=', 'Tc=', 'tC=']) {
    const source = parts();
    xml(source, 'xl/comments1.xml', (s) => s.replace('tc=', marker));
    incomplete(await read(source), 'ambiguous-placeholder');
  }
});
test('legal literal DTD-looking text and comments do not change source membership', async () => {
  const source = parts();
  xml(source, 'xl/comments1.xml', (s) =>
    s
      .replace('Note without a cell record', '<![CDATA[<!DOCTYPE notes [<!ENTITY note "plain text">]>]]>')
      .replace('</comments>', '<!-- <!DOCTYPE ignored [<!ENTITY x "text">]> --></comments>')
  );
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) =>
    s
      .replace('A threaded discussion', '<![CDATA[<!DOCTYPE text><!ENTITY visible "words">]]>')
      .replace('</ThreadedComments>', '<?audit literal="<!DOCTYPE visible>"?></ThreadedComments>')
  );
  const result = await read(source);
  assert.equal(result.status, 'complete');
  assert.deepEqual(result.notes, ['D4', 'XFD1048576']);
  assert.deepEqual(result.threadedPlaceholders, ['B2']);
});
test('actual DTD/entity declarations and document-level CDATA never authorize membership', async () => {
  for (const update of [
    (s) => '<!DOCTYPE comments [<!ENTITY x "text">]>' + s,
    (s) => s.replace('<comments ', '<!ENTITY x "text"><comments '),
  ]) {
    const source = parts();
    xml(source, 'xl/comments1.xml', update);
    incomplete(await read(source), 'unqualified-extension');
  }
  for (const update of [(s) => '<![CDATA[ ]]>' + s, (s) => s + '<![CDATA[ ]]>', (s) => '<![CDATA[visible]]>' + s]) {
    const source = parts();
    xml(source, 'xl/comments1.xml', update);
    incomplete(await read(source), 'malformed-xml');
  }
});
test('unqualified extensions fail even with empty legacy inventories', async () => {
  for (const [name, update] of [
    ['xl/comments1.xml', (s) => s.replace('</comments>', '<extLst/></comments>')],
    [
      'xl/threadedComments/threadedComment1.xml',
      (s) => s.replace('</ThreadedComments>', '<extLst/></ThreadedComments>'),
    ],
    ['xl/persons/person.xml', (s) => s.replace('/></personList>', '><extLst/></person></personList>')],
    ['xl/comments1.xml', (s) => s.replace('ref="D4"', 'ref="D4" foreign="yes"')],
  ]) {
    const source = parts();
    xml(source, name, update);
    incomplete(await read(source));
  }
});
test('placeholder tokens require unique top-level thread identity', async () => {
  for (const [name, update] of [
    ['xl/comments1.xml', (s) => s.replace('tc=' + noteThreadId, 'tc={bad}')],
    ['xl/comments1.xml', (s) => s.replace('tc=' + noteThreadId, 'tc=' + noteThreadId + ' tc=' + noteThreadId)],
    [
      'xl/threadedComments/threadedComment1.xml',
      (s) => s.replace('id="' + noteThreadId + '"', 'id="' + notePersonId + '"'),
    ],
    [
      'xl/threadedComments/threadedComment1.xml',
      (s) => s.replace(' id="' + noteThreadId + '"', ' id="' + noteThreadId + '" parentId="' + notePersonId + '"'),
    ],
    ['xl/comments1.xml', (s) => s.replace('Thread author tc=' + noteThreadId, 'Thread author')],
  ]) {
    const source = parts();
    xml(source, name, update);
    incomplete(await read(source));
  }
});
test('qualified reply/person graphs are accepted and missing/cyclic/duplicate thread graphs fail', async () => {
  const source = parts(),
    reply = '{34567890-3456-4456-8456-34567890ABCD}';
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) =>
    s.replace(
      '</ThreadedComments>',
      '<threadedComment personId="' +
        notePersonId +
        '" id="' +
        reply +
        '" parentId="' +
        noteThreadId +
        '"><text>Reply</text></threadedComment></ThreadedComments>'
    )
  );
  assert.equal((await read(source)).status, 'complete');
  for (const update of [
    (s) => s.replace(reply, noteThreadId),
    (s) => s.replace('parentId="' + noteThreadId + '"', 'parentId="' + reply + '"'),
    (s) => s.replace(notePersonId, '{99999999-9999-4999-8999-999999999999}'),
  ]) {
    const altered = { ...source };
    xml(altered, 'xl/threadedComments/threadedComment1.xml', update);
    incomplete(await read(altered));
  }
});
test('long reply chains validate in linear work and cycles fail without recursion', async () => {
  const source = parts();
  let previous = noteThreadId;
  const replies = [];
  for (let i = 0; i < 10000; i++) {
    const id = '{' + i.toString(16).padStart(8, '0') + '-3456-4456-8456-34567890ABCD}';
    replies.push('<threadedComment personId="' + notePersonId + '" id="' + id + '" parentId="' + previous + '"/>');
    previous = id;
  }
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) =>
    s.replace('</ThreadedComments>', replies.reverse().join('') + '</ThreadedComments>')
  );
  assert.equal((await read(source)).status, 'complete');
  xml(source, 'xl/threadedComments/threadedComment1.xml', (s) =>
    s.replace('id="' + noteThreadId + '"', 'id="' + noteThreadId + '" parentId="' + previous + '"')
  );
  incomplete(await read(source), 'malformed-thread');
});
test('shared, orphaned, external, and foreign-owner annotation parts are incomplete', async () => {
  const source = parts();
  source['xl/drawings/_rels/foreign.xml.rels'] = strToU8(
    '<Relationships xmlns="' +
      OPC +
      '"><Relationship Id="foreign" Type="' +
      OFFICE +
      '/image" Target="/xl/comments1.xml"/></Relationships>'
  );
  incomplete(await read(source), 'shared-part');
  const foreign = parts();
  xml(foreign, 'xl/worksheets/_rels/sheet1.xml.rels', (s) =>
    s.replace('Type="' + OFFICE + '/comments"', 'Type="' + OFFICE + '/image"')
  );
  incomplete(await read(foreign), 'foreign-graph');
  const external = parts();
  xml(external, 'xl/worksheets/_rels/sheet1.xml.rels', (s) =>
    s.replace('Id="notes"', 'Id="notes" TargetMode="External"')
  );
  incomplete(await read(external), 'foreign-graph');
  const owner = parts();
  xml(owner, 'xl/_rels/workbook.xml.rels', (s) => s.replace('/person"', '/threadedComment"'));
  incomplete(await read(owner), 'foreign-graph');
});
test('source worksheet identity and graph MIME cannot be caller-controlled', async () => {
  incomplete(await read(parts(), { source: { ...SOURCE, sheetPath: 'xl/worksheets/other.xml' } }), 'source-mismatch');
  const wrong = parts();
  xml(wrong, '[Content_Types].xml', (s) =>
    s.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml', 'application/xml')
  );
  incomplete(await read(wrong), 'ambiguous-graph');
});
function patchDeclaredRaw(bytes, name, raw) {
  const copy = Buffer.from(bytes);
  for (let offset = 0; offset < copy.length - 46; offset++)
    if (copy.readUInt32LE(offset) === 0x02014b50) {
      const length = copy.readUInt16LE(offset + 28);
      if (copy.subarray(offset + 46, offset + 46 + length).toString() === name) {
        copy.writeUInt32LE(raw, offset + 24);
        const local = copy.readUInt32LE(offset + 42);
        copy.writeUInt32LE(raw, local + 22);
        return copy;
      }
    }
  throw Error('Fixture member not found.');
}
test('selected declared expanded sizes are bounded before inflation; huge worksheets stay unread', async () => {
  incomplete(await read(patchDeclaredRaw(encode(parts()), 'xl/comments1.xml', 17 * 1024 * 1024)), 'resource-limit');
  const result = await read(patchDeclaredRaw(encode(parts()), SOURCE.sheetPath, 512 * 1024 * 1024));
  assert.equal(result.status, 'complete', 'worksheet inflation is never needed for note membership');
});
test('selected compressed bytes have their own bound before payload allocation', async () => {
  const source = parts();
  source['xl/comments1.xml'] = new Uint8Array(17 * 1024 * 1024);
  const bytes = zipSync(source, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') });
  incomplete(await read(patchDeclaredRaw(bytes, 'xl/comments1.xml', 16)), 'resource-limit');
});
test('selected parts share a cumulative 64 MiB resource budget', async () => {
  const data = new Uint8Array(13 * 1024 * 1024),
    members = Object.fromEntries(Array.from({ length: 5 }, (_, i) => ['part' + i + '.xml', data]));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'workbook-note-total-')),
    file = path.join(dir, 'metadata.xlsx');
  fs.writeFileSync(file, zipSync(members, { level: 6, mtime: new Date('2020-01-01T00:00:00Z') }));
  let zip;
  try {
    zip = await openNoteZipFile(file);
    for (let i = 0; i < 4; i++) await zip.read('part' + i + '.xml');
    await assert.rejects(
      () => zip.read('part4.xml'),
      (error) => error.noteMembershipCode === 'resource-limit'
    );
  } finally {
    await zip?.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
test('declared undersized expansion cannot authorize truncated XML', async () => {
  const source = parts();
  incomplete(await read(patchDeclaredRaw(encode(source), 'xl/comments1.xml', 32)), 'malformed-package');
});
test('annotation, XML-element, and nesting resource limits are explicit incomplete outcomes', async () => {
  const source = parts({ mixed: false });
  source['xl/comments1.xml'] = strToU8(
    '<comments xmlns="' +
      MAIN +
      '"><authors><author>A</author></authors><commentList>' +
      Array.from({ length: 100001 }, (_, i) => '<comment ref="A' + (i + 1) + '" authorId="0"><text/></comment>').join(
        ''
      ) +
      '</commentList></comments>'
  );
  incomplete(await read(source), 'resource-limit');
  const deep = parts({ mixed: false });
  xml(deep, 'xl/comments1.xml', (s) => s.replace('<t/>', '<r>'.repeat(130) + '<t/>' + '</r>'.repeat(130)));
  incomplete(await read(deep), 'resource-limit');
});
test('abort/retirement checks reject without partial authority and close newly opened descriptors', async () => {
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(() => read(parts(), { signal: controller.signal }), /abort/i);
  let checks = 0;
  await assert.rejects(
    () =>
      read(parts(), {
        check: () => {
          if (++checks === 4) throw Error('retired');
        },
      }),
    /retired/
  );
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'workbook-note-close-')),
    file = path.join(dir, 'metadata.xlsx');
  fs.writeFileSync(file, encode(parts()));
  const original = fs.promises.open;
  let opened,
    closed = false;
  fs.promises.open = async (...args) => {
    const handle = await original(...args);
    opened = handle;
    const close = handle.close.bind(handle);
    handle.close = async () => {
      closed = true;
      return close();
    };
    return handle;
  };
  try {
    let guardCount = 0;
    await assert.rejects(
      () =>
        openNoteZipFile(file, () => {
          if (++guardCount === 2) throw Error('retired at open');
        }),
      /retired at open/
    );
    assert.ok(opened);
    assert.equal(closed, true);
  } finally {
    fs.promises.open = original;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
test('malformed ZIP directory and duplicate/shared member payloads are incomplete', async () => {
  const bytes = Buffer.from(encode(parts()));
  bytes[bytes.length - 22] = 0;
  incomplete(await read(bytes), 'malformed-package');
  const directory = parts();
  directory['../bad/'] = new Uint8Array();
  incomplete(await read(encode(directory)), 'malformed-package');
});
