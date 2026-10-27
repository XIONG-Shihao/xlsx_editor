/** Source-backed membership only. No selector geometry, text-prefix inference, or writes. */
import path from 'node:path';
import { DOMParser } from '@xmldom/xmldom';
import { openNoteZipFile } from './note-zip-reader.mjs';
import {
  noteMembershipSheet,
  noteMembershipSource,
  noteMembershipCell,
  noteMembershipInner,
  NOTE_MEMBERSHIP_MAX_ANNOTATIONS,
} from './note-membership.mjs';
import { refusal } from './refusal-codes.mjs';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  STRICT = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships',
  CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  STRICT_OFFICE = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const THREAD = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
  EXT_REL = 'http://schemas.microsoft.com/office/2017/10/relationships';
const REV = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  XML = 'http://www.w3.org/XML/1998/namespace';
const COMMENTS_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  THREAD_MIME = 'application/vnd.ms-excel.threadedcomments+xml',
  PERSON_MIME = 'application/vnd.ms-excel.person+xml';
const SHEET_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  RELS_MIME = 'application/vnd.openxmlformats-package.relationships+xml';
const WORKBOOK_MIMES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  'application/vnd.ms-excel.template.macroEnabled.main+xml',
];
const role = (type) =>
  [OFFICE + '/comments', STRICT_OFFICE + '/comments'].includes(type)
    ? 'comments'
    : type === EXT_REL + '/threadedComment'
      ? 'thread'
      : type === EXT_REL + '/person'
        ? 'person'
        : null;
const fail = (code, message) => {
  const error = refusal('package-malformed', message);
  error.noteMembershipCode = code;
  throw error;
};
const children = (node) => Array.from(node.childNodes).filter((child) => child.nodeType === 1);
const attr = (node, name) => (node.hasAttribute(name) ? node.getAttribute(name) : null);
const same = (node, namespace, name) => node.namespaceURI === namespace && node.localName === name;
const guid = (value) => {
  if (typeof value !== 'string' || !/^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/i.test(value))
    fail('invalid-thread-id', 'A threaded comment or person has an invalid GUID.');
  return value.toUpperCase();
};
const uint = (value) =>
  typeof value === 'string' &&
  /^[+-]?[0-9]+$/.test(value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '')) &&
  Number(value) >= 0 &&
  Number(value) <= 4294967295;
function sourceRef(value) {
  if (typeof value !== 'string') fail('invalid-reference', 'An annotation has no single-cell reference.');
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(value);
  if (!match) fail('invalid-reference', 'An annotation has an invalid single-cell reference.');
  const canonical = match[1].toUpperCase() + match[2];
  try {
    noteMembershipCell(canonical);
  } catch {
    fail('invalid-reference', 'An annotation reference exceeds worksheet limits.');
  }
  return canonical;
}
function attributes(node, allowed, { uid = false, space = false, ignorable = false } = {}) {
  for (const item of Array.from(node.attributes ?? [])) {
    if (item.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
    if (!item.namespaceURI && allowed.includes(item.name)) continue;
    if (uid && item.namespaceURI === REV && item.localName === 'uid') continue;
    if (
      space &&
      item.namespaceURI === XML &&
      item.localName === 'space' &&
      ['default', 'preserve'].includes(item.value)
    )
      continue;
    if (
      ignorable &&
      item.namespaceURI === MC &&
      item.localName === 'Ignorable' &&
      item.value
        .trim()
        .split(/\s+/)
        .every((prefix) => node.lookupNamespaceURI(prefix) === REV)
    )
      continue;
    fail('unqualified-extension', 'Annotation metadata has an unqualified attribute or extension.');
  }
}
function elementChildren(node, namespace, allowed) {
  const result = children(node);
  if (result.some((child) => child.namespaceURI !== namespace || !allowed.includes(child.localName)))
    fail('foreign-graph', 'Annotation metadata has a foreign or unqualified element.');
  for (const child of Array.from(node.childNodes))
    if ([3, 4].includes(child.nodeType) && child.data.trim())
      fail('malformed-xml', 'Annotation structural metadata has unexpected text.');
  return result;
}
function xmlSource(bytes) {
  let encoding = 'utf-8';
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0)) encoding = 'utf-16le';
  else if ((bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c)) encoding = 'utf-16be';
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch {
    fail('malformed-xml', 'Note metadata has invalid XML text encoding.');
  }
}
function parseXml(bytes, budget) {
  const source = xmlSource(bytes);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(source))
    fail('malformed-xml', 'Note metadata contains invalid XML characters.');
  // Bound element count and depth before DOM allocation. Skip quoted attributes,
  // comments, CDATA, and processing instructions so literal tag text has no role.
  let count = 0,
    depth = 0,
    cursor = 0;
  while (true) {
    const next = source.indexOf('<', cursor);
    if (next < 0) {
      if (depth === 0 && source.slice(cursor).trim())
        fail('malformed-xml', 'Note metadata has text outside its XML root.');
      break;
    }
    if (depth === 0 && source.slice(cursor, next).trim())
      fail('malformed-xml', 'Note metadata has text outside its XML root.');
    cursor = next;
    let terminal;
    if (source.startsWith('<!--', cursor)) {
      terminal = source.indexOf('-->', cursor + 4);
      if (terminal < 0) fail('malformed-xml', 'Note metadata has an unterminated XML comment.');
      cursor = terminal + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', cursor)) {
      if (depth === 0) fail('malformed-xml', 'CDATA is not allowed outside the XML root.');
      terminal = source.indexOf(']]>', cursor + 9);
      if (terminal < 0) fail('malformed-xml', 'Note metadata has unterminated CDATA.');
      cursor = terminal + 3;
      continue;
    }
    if (source.startsWith('<?', cursor)) {
      terminal = source.indexOf('?>', cursor + 2);
      if (terminal < 0) fail('malformed-xml', 'Note metadata has an unterminated XML instruction.');
      cursor = terminal + 2;
      continue;
    }
    if (/^<!(?:DOCTYPE|ENTITY)\b/i.test(source.slice(cursor, cursor + 10)))
      fail('unqualified-extension', 'DTD and custom entities are unqualified note metadata.');
    let quote = null;
    terminal = cursor + 1;
    for (; terminal < source.length; terminal++) {
      const char = source[terminal];
      if (quote) {
        if (char === quote) quote = null;
      } else if (char === '"' || char === "'") quote = char;
      else if (char === '>') break;
    }
    if (terminal === source.length) fail('malformed-xml', 'Note metadata has an unterminated XML tag.');
    if (source[cursor + 1] === '/') {
      if (--depth < 0) fail('malformed-xml', 'Note metadata XML is unbalanced.');
    } else {
      if (++count > 300000 || budget.elements + count > 600000)
        fail('resource-limit', 'Selected note metadata exceeds its XML element resource limit.');
      if (!/\/\s*$/.test(source.slice(cursor + 1, terminal)) && ++depth > 128)
        fail('resource-limit', 'Selected note metadata exceeds its XML depth resource limit.');
    }
    cursor = terminal + 1;
  }
  if (depth !== 0) fail('malformed-xml', 'Note metadata XML is unbalanced.');
  budget.elements += count;
  let doc;
  try {
    doc = new DOMParser({
      errorHandler: {
        warning() {
          throw refusal('note-malformed', 'Invalid XML.');
        },
        error() {
          throw refusal('note-malformed', 'Invalid XML.');
        },
        fatalError() {
          throw refusal('note-malformed', 'Invalid XML.');
        },
      },
    }).parseFromString(source, 'application/xml');
  } catch {
    fail('malformed-xml', 'Note metadata XML is malformed.');
  }
  if (
    !doc?.documentElement ||
    children(doc).length !== 1 ||
    Array.from(doc.childNodes).some((node) => [3, 4].includes(node.nodeType) && node.data.trim())
  )
    fail('malformed-xml', 'Note metadata must contain one XML root.');
  const stack = [[doc.documentElement, 1]];
  while (stack.length) {
    const [node, depth] = stack.pop();
    if (depth > 128) fail('resource-limit', 'Selected note metadata exceeds its XML depth resource limit.');
    for (const child of children(node)) stack.push([child, depth + 1]);
  }
  return doc.documentElement;
}
function canonicalPart(value) {
  if (
    typeof value !== 'string' ||
    !value ||
    value.length > 1024 ||
    value.includes('\\') ||
    /[\0?#]/.test(value) ||
    value.startsWith('//') ||
    /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value)
  )
    fail('invalid-relationship', 'A note relationship has an invalid internal target.');
  const normalized = path.posix.normalize(value);
  if (
    normalized === '.' ||
    normalized === '..' ||
    normalized.startsWith('../') ||
    normalized.split('/').some((segment) => !segment)
  )
    fail('invalid-relationship', 'A note relationship escapes the package.');
  return normalized;
}
function resolvePart(owner, target, zip) {
  if (typeof target !== 'string' || target.startsWith('//') || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target))
    fail('invalid-relationship', 'A note relationship has an invalid internal target URI.');
  const rooted = target.startsWith('/'),
    part = canonicalPart(rooted ? target.slice(1) : path.posix.join(path.posix.dirname(owner || '.'), target));
  let decoded;
  try {
    decoded = decodeURIComponent(part);
  } catch {
    fail('invalid-relationship', 'A note relationship has malformed URI escapes.');
  }
  if (decoded !== part) {
    canonicalPart(decoded);
    if (/%2f|%5c/i.test(part)) fail('invalid-relationship', 'A note relationship has escaped path separators.');
    if (zip.has(part) && zip.has(decoded))
      fail('ambiguous-graph', 'A note relationship has ambiguous escaped part names.');
    if (zip.has(decoded)) return decoded;
  }
  return part;
}
const relationshipOwner = (name) =>
  name === '_rels/.rels'
    ? ''
    : name.includes('/_rels/') && name.endsWith('.rels')
      ? name.replace('/_rels/', '/').slice(0, -5)
      : name.startsWith('_rels/') && name.endsWith('.rels')
        ? name.slice(6, -5)
        : null;
function relationshipGraph(root, owner, zip) {
  if (!same(root, REL, 'Relationships')) fail('wrong-root', 'A relationship part has the wrong XML root or namespace.');
  attributes(root, []);
  const ids = new Set();
  return elementChildren(root, REL, ['Relationship']).map((node) => {
    attributes(node, ['Id', 'Type', 'Target', 'TargetMode']);
    if (children(node).length) fail('malformed-graph', 'A relationship contains nested metadata.');
    const id = attr(node, 'Id'),
      type = attr(node, 'Type'),
      target = attr(node, 'Target'),
      mode = attr(node, 'TargetMode');
    if (!id || ids.has(id) || !type || !target || (mode !== null && !['Internal', 'External'].includes(mode)))
      fail('duplicate-relationship', 'A relationship has invalid or duplicate fields.');
    ids.add(id);
    return {
      owner,
      id,
      type,
      external: mode === 'External',
      target: mode === 'External' ? target : resolvePart(owner, target, zip),
    };
  });
}
function contentTypes(root, zip) {
  if (!same(root, CT, 'Types')) fail('wrong-root', 'Content types have the wrong XML root or namespace.');
  attributes(root, []);
  const defaults = new Map(),
    overrides = new Map();
  for (const node of elementChildren(root, CT, ['Default', 'Override'])) {
    const isDefault = node.localName === 'Default';
    attributes(node, isDefault ? ['Extension', 'ContentType'] : ['PartName', 'ContentType']);
    if (children(node).length) fail('malformed-graph', 'A content type contains nested metadata.');
    const type = attr(node, 'ContentType'),
      raw = attr(node, isDefault ? 'Extension' : 'PartName');
    if (!type || !raw || (isDefault && !/^[A-Za-z0-9_-]+$/.test(raw)) || (!isDefault && !raw.startsWith('/')))
      fail('invalid-content-type', 'A content type has invalid fields.');
    const key = isDefault ? raw.toLowerCase() : resolvePart('', raw, zip),
      map = isDefault ? defaults : overrides;
    if (map.has(key)) fail('duplicate-content-type', 'A content type is declared more than once.');
    map.set(key, type);
  }
  // OPC's package relationship part is named .rels; filesystem extname treats
  // that name as extensionless, so resolve extensions from the final dot.
  return {
    overrides,
    type: (part) => {
      const name = path.posix.basename(part),
        dot = name.lastIndexOf('.');
      return overrides.get(part) ?? defaults.get(dot < 0 ? '' : name.slice(dot + 1).toLowerCase());
    },
  };
}
function legacyComments(root, namespace, budget) {
  if (!same(root, namespace, 'comments'))
    fail('wrong-root', 'A legacy comments part has the wrong XML root or namespace.');
  attributes(root, [], { ignorable: true });
  const lists = elementChildren(root, namespace, ['authors', 'commentList']);
  if (lists.length !== 2 || lists[0].localName !== 'authors' || lists[1].localName !== 'commentList')
    fail('malformed-graph', 'Legacy comments require one authors list and one comment list.');
  attributes(lists[0], []);
  attributes(lists[1], []);
  const authors = elementChildren(lists[0], namespace, ['author']).map((author) => {
    attributes(author, [], { space: true });
    if (children(author).length) fail('invalid-author', 'A legacy comment author has nested metadata.');
    return author.textContent;
  });
  const refs = new Set(),
    result = [];
  for (const node of elementChildren(lists[1], namespace, ['comment'])) {
    if (++budget.annotations > NOTE_MEMBERSHIP_MAX_ANNOTATIONS)
      fail('resource-limit', 'Note graph exceeds its 100000 annotation resource limit.');
    attributes(node, ['ref', 'authorId', 'shapeId', 'uid'], { uid: true });
    const ref = sourceRef(attr(node, 'ref')),
      authorId = attr(node, 'authorId');
    if (refs.has(ref)) fail('duplicate-annotation', 'Legacy comments repeat a cell reference.');
    refs.add(ref);
    if (!uint(authorId) || Number(authorId) >= authors.length)
      fail('invalid-author', 'A legacy comment has an invalid author index.');
    if (attr(node, 'shapeId') !== null && !uint(attr(node, 'shapeId')))
      fail('invalid-author', 'A legacy comment has an invalid shape identifier.');
    const content = elementChildren(node, namespace, ['text']);
    if (content.length !== 1) fail('malformed-graph', 'A legacy comment requires exactly one text element.');
    // Rich and empty text both authorize membership; their wording has no role.
    const text = content[0];
    attributes(text, []);
    const fonts = [
      'rFont',
      'charset',
      'family',
      'b',
      'i',
      'strike',
      'outline',
      'shadow',
      'condense',
      'extend',
      'color',
      'sz',
      'u',
      'vertAlign',
      'scheme',
    ];
    const validateRich = (item, parent) => {
      if (
        item.namespaceURI !== namespace ||
        !({ text: ['t', 'r', 'rPh', 'phoneticPr'], r: ['rPr', 't'], rPr: fonts, rPh: ['t'] }[parent] ?? []).includes(
          item.localName
        )
      )
        fail('unqualified-extension', 'Legacy comment text contains an unqualified extension.');
      if (item.localName === 't') {
        attributes(item, [], { space: true });
        if (children(item).length) fail('malformed-graph', 'Legacy comment text has nested text.');
        return;
      }
      if (item.localName === 'rPh') attributes(item, ['sb', 'eb']);
      else if (item.localName === 'phoneticPr') attributes(item, ['fontId', 'type', 'alignment']);
      else if (fonts.includes(item.localName))
        attributes(item, item.localName === 'color' ? ['rgb', 'indexed', 'theme', 'tint', 'auto'] : ['val']);
      else attributes(item, []);
      for (const child of children(item)) validateRich(child, item.localName);
      if (!['r', 'rPr', 'rPh'].includes(item.localName) && children(item).length)
        fail('malformed-graph', 'Legacy comment formatting contains nested metadata.');
    };
    for (const child of elementChildren(text, namespace, ['t', 'r', 'rPh', 'phoneticPr'])) validateRich(child, 'text');
    const author = authors[Number(authorId)],
      tokens = [
        ...author.matchAll(/tc=(\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\})/g),
      ];
    if (/tc=/i.test(author) && (tokens.length !== 1 || author.replace(tokens[0]?.[0] ?? '', '').match(/tc=/i)))
      fail('ambiguous-placeholder', 'A compatibility author has a malformed or ambiguous thread token.');
    const unqualified = attr(node, 'uid'),
      qualified = node.hasAttributeNS(REV, 'uid') ? node.getAttributeNS(REV, 'uid') : null;
    if (unqualified !== null && qualified !== null)
      fail('ambiguous-placeholder', 'A legacy comment has multiple UID attributes.');
    result.push({
      ref,
      token: tokens.length ? guid(tokens[0][1]) : null,
      uid: unqualified !== null ? guid(unqualified) : qualified !== null ? guid(qualified) : null,
    });
  }
  return result;
}
function persons(root, budget) {
  if (!same(root, THREAD, 'personList')) fail('wrong-root', 'The persons part has the wrong XML root or namespace.');
  attributes(root, []);
  const ids = new Set();
  for (const node of elementChildren(root, THREAD, ['person'])) {
    if (++budget.people > NOTE_MEMBERSHIP_MAX_ANNOTATIONS)
      fail('resource-limit', 'Note graph exceeds its person resource limit.');
    attributes(node, ['displayName', 'id', 'userId', 'providerId']);
    if (attr(node, 'displayName') === null || children(node).length)
      fail('unqualified-extension', 'A person has missing or extended metadata.');
    const id = guid(attr(node, 'id'));
    if (ids.has(id)) fail('duplicate-person', 'The persons graph has duplicate IDs.');
    ids.add(id);
  }
  return ids;
}
function threads(root, budget, people) {
  if (!same(root, THREAD, 'ThreadedComments'))
    fail('wrong-root', 'The threaded comments part has the wrong XML root or namespace.');
  attributes(root, []);
  const result = [];
  for (const node of elementChildren(root, THREAD, ['threadedComment'])) {
    if (++budget.annotations > NOTE_MEMBERSHIP_MAX_ANNOTATIONS)
      fail('resource-limit', 'Note graph exceeds its 100000 annotation resource limit.');
    attributes(node, ['ref', 'dT', 'personId', 'id', 'parentId', 'done']);
    const id = guid(attr(node, 'id')),
      personId = guid(attr(node, 'personId')),
      parentId = attr(node, 'parentId'),
      ref = attr(node, 'ref') === null ? null : sourceRef(attr(node, 'ref'));
    if (!people.has(personId))
      fail('missing-person', 'A threaded comment author is absent from the qualified persons graph.');
    if (attr(node, 'done') !== null && !['0', '1', 'false', 'true'].includes(attr(node, 'done').trim()))
      fail('malformed-graph', 'A threaded comment has an invalid resolved flag.');
    const content = elementChildren(node, THREAD, ['text', 'mentions']);
    if (
      content.filter((child) => child.localName === 'text').length > 1 ||
      content.filter((child) => child.localName === 'mentions').length > 1
    )
      fail('duplicate-annotation', 'A threaded comment repeats content elements.');
    for (const child of content) {
      attributes(child, [], { space: child.localName === 'text' });
      if (child.localName === 'text') {
        if (children(child).length) fail('unqualified-extension', 'Thread text contains extended metadata.');
        continue;
      }
      const mentionIds = new Set();
      for (const mention of elementChildren(child, THREAD, ['mention'])) {
        attributes(mention, ['mentionpersonId', 'mentionId', 'startIndex', 'length']);
        const mentionId = guid(attr(mention, 'mentionId'));
        if (
          mentionIds.has(mentionId) ||
          children(mention).length ||
          !uint(attr(mention, 'startIndex')) ||
          !uint(attr(mention, 'length')) ||
          !people.has(guid(attr(mention, 'mentionpersonId')))
        )
          fail('malformed-graph', 'A thread mention has invalid or missing person metadata.');
        mentionIds.add(mentionId);
      }
    }
    result.push({ id, parentId: parentId === null ? null : guid(parentId), ref });
  }
  return result;
}

export async function indexedNoteMembership({ metadataFile, view, sheet, source, signal, check = () => {} }) {
  noteMembershipSheet(sheet);
  noteMembershipSource(source);
  const identity = { ...source };
  let guardFailure;
  const guard = () => {
    try {
      signal?.throwIfAborted();
      check();
    } catch (error) {
      guardFailure = error;
      throw error;
    }
  };
  guard();
  let zip;
  const base = { schema: 1, policy: 'xlsx-legacy-note-membership-v1', sheet, source: identity };
  try {
    zip = await openNoteZipFile(metadataFile, guard);
    guard();
    const budget = { elements: 0, annotations: 0, people: 0 };
    const root = async (name) => {
      guard();
      const bytes = await zip.read(name);
      guard();
      const parsed = parseXml(bytes, budget);
      guard();
      return parsed;
    };
    const types = contentTypes(await root('[Content_Types].xml'), zip),
      relationships = new Map(),
      all = [];
    for (const name of zip.names)
      if (name.endsWith('.rels')) {
        if (types.type(name) !== RELS_MIME)
          fail('wrong-content-type', 'A relationship part has the wrong content type.');
        const owner = relationshipOwner(name);
        if (owner === null) fail('invalid-relationship', 'A relationship part has an invalid package location.');
        const list = relationshipGraph(await root(name), owner, zip);
        relationships.set(owner, list);
        all.push(...list);
      }
    const office = (relationships.get('') ?? []).filter((rel) =>
      [OFFICE + '/officeDocument', STRICT_OFFICE + '/officeDocument'].includes(rel.type)
    );
    if (office.length !== 1 || office[0].external)
      fail('ambiguous-graph', 'The workbook requires one internal office-document relationship.');
    const workbookPath = office[0].target;
    if (!WORKBOOK_MIMES.includes(types.type(workbookPath)))
      fail('wrong-content-type', 'The workbook part has an unqualified content type.');
    const workbook = await root(workbookPath),
      namespace = workbook.namespaceURI;
    if (![MAIN, STRICT].includes(namespace) || workbook.localName !== 'workbook')
      fail('wrong-root', 'The workbook has the wrong XML root or namespace.');
    const sheetLists = children(workbook).filter((node) => same(node, namespace, 'sheets'));
    if (sheetLists.length !== 1) fail('ambiguous-graph', 'The workbook requires one worksheet inventory.');
    const worksheetRels = (relationships.get(workbookPath) ?? []).filter((rel) =>
        [OFFICE + '/worksheet', STRICT_OFFICE + '/worksheet'].includes(rel.type)
      ),
      byId = new Map(worksheetRels.map((rel) => [rel.id, rel])),
      sheets = new Map(),
      names = new Set();
    for (const node of elementChildren(sheetLists[0], namespace, ['sheet'])) {
      const name = attr(node, 'name'),
        id = node.getAttributeNS(namespace === STRICT ? STRICT_OFFICE : OFFICE, 'id'),
        rel = byId.get(id);
      noteMembershipSheet(name);
      if (
        !rel ||
        rel.external ||
        names.has(name) ||
        sheets.has(rel.target) ||
        types.type(rel.target) !== SHEET_MIME ||
        !zip.has(rel.target)
      )
        fail('ambiguous-graph', 'Workbook worksheet paths or relationships are invalid or shared.');
      names.add(name);
      sheets.set(rel.target, name);
    }
    if (
      worksheetRels.length !== sheets.size ||
      sheets.get(source.sheetPath) !== sheet ||
      !Array.isArray(view?.sheets) ||
      view.sheets.filter((item) => item.name === sheet && item.path === source.sheetPath).length !== 1
    )
      fail('source-mismatch', 'The source worksheet identity disagrees with the package graph.');
    const annotationRels = all.filter((rel) => role(rel.type)),
      parts = new Map(),
      incoming = new Map();
    for (const rel of all) if (!rel.external) incoming.set(rel.target, (incoming.get(rel.target) ?? 0) + 1);
    for (const rel of annotationRels) {
      const kind = role(rel.type);
      if (rel.external || (kind === 'person' ? rel.owner !== workbookPath : !sheets.has(rel.owner)))
        fail('foreign-graph', 'An annotation relationship has a foreign or external owner.');
      const expectedMime = kind === 'comments' ? COMMENTS_MIME : kind === 'thread' ? THREAD_MIME : PERSON_MIME;
      if (!zip.has(rel.target)) fail('missing-part', 'An annotation relationship points to a missing part.');
      if (types.type(rel.target) !== expectedMime)
        fail('wrong-content-type', 'An annotation part has the wrong content type.');
      if (parts.has(rel.target) || incoming.get(rel.target) !== 1)
        fail('shared-part', 'An annotation part is duplicated or shared.');
      parts.set(rel.target, rel);
    }
    for (const [part, type] of types.overrides)
      if ([COMMENTS_MIME, THREAD_MIME, PERSON_MIME].includes(type) && !parts.has(part))
        fail('foreign-graph', 'An annotation content type has no qualified owner relationship.');
    // Detect annotation MIME via defaults too, including nonstandard filenames.
    for (const part of zip.names)
      if ([COMMENTS_MIME, THREAD_MIME, PERSON_MIME].includes(types.type(part)) && !parts.has(part))
        fail('foreign-graph', 'An annotation part has no qualified owner relationship.');
    const peopleParts = [...parts.values()].filter((rel) => role(rel.type) === 'person');
    if (peopleParts.length > 1) fail('ambiguous-graph', 'The workbook has multiple persons graphs.');
    const people = peopleParts.length ? persons(await root(peopleParts[0].target), budget) : new Set(),
      legacy = new Map(),
      threadGraphs = new Map(),
      threadIds = new Set();
    for (const [part, rel] of parts) {
      const kind = role(rel.type);
      if (kind === 'person') continue;
      const map = kind === 'comments' ? legacy : threadGraphs;
      if (map.has(rel.owner)) fail('duplicate-part', 'A worksheet has multiple annotation parts of one kind.');
      const parsed =
        kind === 'comments'
          ? legacyComments(await root(part), namespace, budget)
          : threads(await root(part), budget, people);
      map.set(rel.owner, parsed);
      if (kind === 'thread') {
        for (const item of parsed) {
          if (threadIds.has(item.id)) fail('duplicate-thread', 'Thread IDs repeat in the workbook.');
          threadIds.add(item.id);
        }
      }
    }
    const notes = [],
      placeholders = [];
    for (const [owner, name] of sheets) {
      const comments = legacy.get(owner) ?? [],
        thread = threadGraphs.get(owner) ?? [],
        byThreadId = new Map(thread.map((item) => [item.id, item])),
        used = new Set();
      const parentState = new Map();
      for (const item of thread) {
        if (parentState.get(item.id) === 2) continue;
        const chain = [];
        let current = item;
        while (parentState.get(current.id) !== 2) {
          if (parentState.get(current.id) === 1) fail('malformed-thread', 'A thread reply has a cyclic parent.');
          parentState.set(current.id, 1);
          chain.push(current.id);
          if (current.parentId === null) break;
          current = byThreadId.get(current.parentId);
          if (!current) fail('malformed-thread', 'A thread reply has a missing or foreign parent.');
        }
        for (const id of chain) parentState.set(id, 2);
      }
      for (const comment of comments) {
        if (comment.token) {
          const top = byThreadId.get(comment.token);
          if (!top || top.parentId !== null || used.has(top.id))
            fail('ambiguous-placeholder', 'A compatibility placeholder lacks a unique matching top-level thread.');
          used.add(top.id);
          if (name === sheet) placeholders.push(comment.ref);
        } else {
          if (comment.uid !== null && threadIds.has(comment.uid))
            fail('ambiguous-placeholder', 'A legacy comment UID refers to a thread without its required author token.');
          if (name === sheet) notes.push(comment.ref);
        }
      }
      if (thread.some((item) => item.parentId === null && !used.has(item.id)))
        fail('missing-placeholder', 'A top-level thread has no qualified legacy compatibility placeholder.');
    }
    const sort = (refs) => refs.sort((a, b) => noteMembershipCell(a) - noteMembershipCell(b));
    guard();
    return noteMembershipInner(
      { ...base, status: 'complete', notes: sort(notes), threadedPlaceholders: sort(placeholders), reasons: [] },
      { sheet, source: identity }
    );
  } catch (error) {
    if (guardFailure) throw guardFailure;
    guard();
    return noteMembershipInner(
      {
        ...base,
        status: 'incomplete',
        notes: null,
        threadedPlaceholders: null,
        reasons: [
          {
            code: error.noteMembershipCode ?? 'unreadable-source',
            message: error.noteMembershipCode
              ? error.message
              : 'The note membership source could not be read completely.',
          },
        ],
      },
      { sheet, source: identity }
    );
  } finally {
    if (zip) await zip.close();
    if (guardFailure) throw guardFailure;
    guard();
  }
}
