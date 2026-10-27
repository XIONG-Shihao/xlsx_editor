import { lineStarts } from './xml-position.mjs';
import { DOMParser, DOMImplementation } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { cellPosition, columnName } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const decode = (value) =>
  value.replace(/_x([0-9a-fA-F]{4})_/g, (_match, code) => String.fromCharCode(parseInt(code, 16)));

/** Consume XML events without retaining its element tree. The same pinned XML
 * parser and line-ending semantics are used by the preserving DOM writer. */
function visitXml(bytes, visitor) {
  const source = strFromU8(bytes).replace(/\r\n?/g, '\n');
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('package-malformed', 'XML entities are unsupported');
  const doc = new DOMImplementation().createDocument(null, null, null);
  const lines = visitor.locations ? lineStarts(source) : undefined;
  const stack = [];
  const builder = {
    doc,
    startDocument() {},
    endDocument() {},
    startPrefixMapping() {},
    endPrefixMapping() {},
    setDocumentLocator(locator) {
      this.locator = locator;
      locator.lineNumber = 0;
    },
    startCDATA() {
      this.cdata = true;
    },
    endCDATA() {
      this.cdata = false;
    },
    processingInstruction(target, data) {
      visitor.instruction?.(target, data, doc);
    },
    comment(chars, start, length) {
      visitor.comment?.(String(chars).substr(start, length), doc);
    },
    ignorableWhitespace() {},
    startElement(uri, local, name, attributes) {
      const attrs = Object.create(null);
      for (let i = 0; i < attributes.length; i++) attrs[attributes.getQName(i)] = attributes.getValue(i);
      const entry = { uri, local, name, attrs };
      if (lines) entry.start = lines[this.locator.lineNumber - 1] + this.locator.columnNumber - 1;
      visitor.start?.(entry, stack, attributes, doc);
      stack.push(entry);
    },
    endElement() {
      const entry = stack.pop();
      visitor.end?.(entry, stack, doc);
    },
    characters(chars, start, length) {
      visitor.text?.(String(chars).substr(start, length), stack, doc, this.cdata);
    },
  };
  visitor.source = source;
  new DOMParser({
    domBuilder: builder,
    ...(lines ? { locator: {} } : {}),
    normalizeLineEndings: (value) => value,
    errorHandler: {
      warning(message) {
        throw refusal('package-malformed', message);
      },
      error(message) {
        throw refusal('package-malformed', message);
      },
      fatalError(message) {
        throw refusal('package-malformed', message);
      },
    },
  }).parseFromString(source, 'application/xml');
  return { source, document: doc };
}

/** Find a selected element's closing boundary while respecting quoted attributes,
 * comments and CDATA. XML validity/namespaces are checked by the SAX pass. */
export function xmlElementEnd(source, start) {
  let depth = 0,
    at = start;
  while (at >= 0 && at < source.length) {
    if (source.startsWith('<!--', at)) {
      at = source.indexOf('-->', at + 4);
      if (at < 0) break;
      at = source.indexOf('<', at + 3);
      continue;
    }
    if (source.startsWith('<![CDATA[', at)) {
      at = source.indexOf(']]>', at + 9);
      if (at < 0) break;
      at = source.indexOf('<', at + 3);
      continue;
    }
    if (source.startsWith('<?', at)) {
      at = source.indexOf('?>', at + 2);
      if (at < 0) break;
      at = source.indexOf('<', at + 2);
      continue;
    }
    let end = at + 1,
      quote;
    for (; end < source.length; end++) {
      const char = source[end];
      if (quote) {
        if (char === quote) quote = undefined;
      } else if (char === '"' || char === "'") quote = char;
      else if (char === '>') break;
    }
    if (end === source.length) break;
    if (source[at + 1] === '/') depth--;
    else if (!/\/\s*$/.test(source.slice(at, end))) depth++;
    if (!depth) return end + 1;
    at = source.indexOf('<', end + 1);
  }
  throw refusal('package-unsupported', 'Incomplete selected XML element');
}

const escapeAttribute = (value) => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
function namespaces(stack) {
  const result = { xml: 'http://www.w3.org/XML/1998/namespace' };
  for (const entry of stack)
    for (const [name, value] of Object.entries(entry.attrs)) {
      if (name === 'xmlns') result[''] = value;
      else if (name.startsWith('xmlns:')) result[name.slice(6)] = value;
    }
  return result;
}

export function readSharedStringsXml(bytes, { indices } = {}) {
  const strings = [];
  let current,
    text = '',
    index = 0;
  visitXml(bytes, {
    start(entry, stack) {
      if (!stack.length && (entry.uri !== NS || entry.local !== 'sst'))
        throw refusal('package-unsupported', 'Unsupported shared string namespace or root');
      if (entry.uri === NS && entry.local === 'si' && stack.length === 1)
        current = !indices || indices.has(index) ? '' : undefined;
      if (entry.uri === NS && entry.local === 't') text = '';
    },
    text(value, stack) {
      if (
        current !== undefined &&
        stack.at(-1)?.uri === NS &&
        stack.at(-1)?.local === 't' &&
        !stack.some((entry) => entry.uri === NS && entry.local === 'rPh')
      )
        text += value;
    },
    end(entry, stack) {
      if (
        entry.uri === NS &&
        entry.local === 't' &&
        current !== undefined &&
        !stack.some((item) => item.uri === NS && item.local === 'rPh')
      )
        current += decode(text);
      if (entry.uri === NS && entry.local === 'si' && stack.length === 1) {
        if (current !== undefined) strings[index] = current;
        index++;
        current = undefined;
      }
    },
  });
  strings.length = index;
  return strings;
}

/** Keep worksheet metadata and cell values, not millions of DOM nodes. */
export function readWorksheetXml(
  bytes,
  readCell,
  { includeCell, onCell, capture = false, includeFormulas = false, targetRows = new Set() } = {}
) {
  const cells = {},
    hiddenRows = [],
    visibleRows = [],
    rowHeights = [];
  const cellLocations = new Map(),
    rowLocations = new Map(),
    selectedRows = new Map(),
    rowCells = new Map(),
    rowStarts = [];
  let sheetDataStart, dimension, rootStart;
  let current,
    text = '',
    cellCount = 0,
    formulaCount = 0,
    specialFormulas = false,
    extended = false;
  let metadataParent,
    skipDepth = 0,
    inferredRow = 0;
  const visitor = {
    locations: capture,
    start(entry, stack, attrs, doc) {
      if (capture && !stack.length) rootStart = entry.start;
      if (capture && entry.uri === NS && entry.local === 'sheetData' && stack.length === 1)
        sheetDataStart = entry.start;
      if (capture && entry.uri === NS && entry.local === 'dimension' && stack.length === 1)
        dimension = { start: entry.start, end: xmlElementEnd(visitor.source, entry.start) };
      if (entry.uri === NS && entry.local === 'extLst') extended = true;
      if (
        entry.uri === NS &&
        entry.local === 'row' &&
        stack.length === 2 &&
        stack.at(-1)?.uri === NS &&
        stack.at(-1)?.local === 'sheetData'
      ) {
        const row = entry.attrs.r === undefined ? inferredRow + 1 : Number(entry.attrs.r);
        if (!Number.isSafeInteger(row) || row < 1 || row > 1048576)
          throw refusal('internal-invalid-field', 'Invalid worksheet row coordinate');
        entry.rowNumber = inferredRow = row;
        entry.column = 0;
        if (capture) rowStarts.push(row, entry.start);
        if (/^(1|true)$/.test(entry.attrs.hidden ?? '')) hiddenRows.push(row - 1);
        if (/^(0|false)$/.test(entry.attrs.hidden ?? '')) visibleRows.push(row - 1);
        if (entry.attrs.ht !== undefined) rowHeights.push({ row: row - 1, height: Number(entry.attrs.ht) });
        skipDepth = stack.length + 1;
      }
      if (entry.uri === NS && entry.local === 'c' && stack.length === 3 && stack.at(-1)?.rowNumber !== undefined) {
        const row = stack.at(-1),
          reference = entry.attrs.r?.replaceAll('$', '').toUpperCase() ?? `${columnName(row.column)}${row.rowNumber}`;
        const position = cellPosition(reference);
        if (position.row + 1 !== row.rowNumber)
          throw refusal('internal-mismatch', 'Cell coordinate disagrees with its row');
        row.column = position.col + 1;
        current = {
          reference,
          attrs: entry.attrs,
          valueText: '',
          inlineText: '',
          valuePresent: false,
          formula: undefined,
          formulaAttrs: undefined,
          start: entry.start,
        };
      } else if (current && entry.uri === NS && entry.local === 'v' && stack.length === 4) current.valuePresent = true;
      else if (current && entry.uri === NS && entry.local === 't' && stack[4]?.uri === NS && stack[4]?.local === 'is')
        text = '';
      else if (current && entry.uri === NS && entry.local === 'f' && stack.length === 4) {
        current.formula = '';
        current.formulaAttrs = entry.attrs;
        formulaCount++;
        if (entry.attrs.t !== undefined) specialFormulas = true;
      }
      if (!skipDepth) {
        const node = doc.createElementNS(entry.uri, entry.name);
        for (let i = 0; i < attrs.length; i++)
          node.setAttributeNS(attrs.getURI(i), attrs.getQName(i), attrs.getValue(i));
        (metadataParent ?? doc).appendChild(node);
        metadataParent = node;
      }
    },
    text(value, stack, doc, cdata) {
      const parent = stack.at(-1);
      if (current && parent?.uri === NS) {
        if (parent.local === 'v' && stack.length === 5) current.valueText += value;
        else if (parent.local === 'f' && stack.length === 5) current.formula += value;
        else if (
          parent.local === 't' &&
          stack[4]?.uri === NS &&
          stack[4]?.local === 'is' &&
          !stack.some((entry) => entry.uri === NS && entry.local === 'rPh')
        )
          text += value;
      }
      if (!skipDepth && metadataParent)
        metadataParent.appendChild(cdata ? doc.createCDATASection(value) : doc.createTextNode(value));
    },
    comment(value, doc) {
      if (!skipDepth) (metadataParent ?? doc).appendChild(doc.createComment(value));
    },
    instruction(target, data, doc) {
      if (!skipDepth) (metadataParent ?? doc).appendChild(doc.createProcessingInstruction(target, data));
    },
    end(entry, stack, doc) {
      if (
        current &&
        entry.uri === NS &&
        entry.local === 't' &&
        stack[4]?.uri === NS &&
        stack[4]?.local === 'is' &&
        !stack.some((item) => item.uri === NS && item.local === 'rPh')
      )
        current.inlineText += decode(text);
      if (entry.uri === NS && entry.local === 'c' && current && stack.length === 3) {
        const ref = current.reference;
        const rowNumber = stack.at(-1)?.rowNumber;
        onCell?.(current, { rowHidden: /^(1|true)$/.test(stack.at(-1)?.attrs.hidden ?? ''), document: doc });
        if (capture && targetRows.has(rowNumber)) {
          let list = rowCells.get(rowNumber);
          if (!list) rowCells.set(rowNumber, (list = []));
          list.push({ ref, start: current.start });
        }
        if (!includeCell || includeCell(ref) || (includeFormulas && current.formulaAttrs)) {
          if (Object.hasOwn(cells, ref)) throw refusal('package-duplicate', `Duplicate worksheet cell: ${ref}`);
          cells[ref] = readCell(current);
          if (capture) {
            const rowEntry = stack.at(-1),
              scope = namespaces([...stack, entry]);
            let row = selectedRows.get(rowEntry.attrs.r);
            if (!row) {
              row = doc.createElementNS(rowEntry.uri, rowEntry.name);
              for (const [name, value] of Object.entries(rowEntry.attrs))
                row.setAttributeNS(
                  name === 'xmlns' || name.startsWith('xmlns:')
                    ? 'http://www.w3.org/2000/xmlns/'
                    : name.includes(':')
                      ? scope[name.split(':')[0]]
                      : null,
                  name,
                  value
                );
              metadataParent.appendChild(row);
              selectedRows.set(rowEntry.attrs.r, row);
            }
            const end = xmlElementEnd(visitor.source, current.start);
            const declarations = Object.entries(scope)
              .filter(([prefix]) => prefix !== 'xml')
              .map(([prefix, value]) => ` xmlns${prefix ? ':' + prefix : ''}="${escapeAttribute(value)}"`)
              .join('');
            const fragment = new DOMParser({
              errorHandler: {
                warning() {},
                error(message) {
                  throw refusal('package-malformed', message);
                },
                fatalError(message) {
                  throw refusal('package-malformed', message);
                },
              },
            }).parseFromString(
              `<container${declarations}>${visitor.source.slice(current.start, end)}</container>`,
              'application/xml'
            );
            row.appendChild(doc.importNode(fragment.documentElement.firstChild, true));
            cellLocations.set(ref, { start: current.start, end });
          }
        }
        cellCount++;
        current = undefined;
      }
      if (
        capture &&
        entry.uri === NS &&
        entry.local === 'row' &&
        stack.length === 2 &&
        entry.rowNumber !== undefined &&
        targetRows.has(entry.rowNumber)
      )
        rowLocations.set(entry.rowNumber, {
          start: entry.start,
          end: xmlElementEnd(visitor.source, entry.start),
          name: entry.name,
        });
      if (skipDepth && stack.length < skipDepth) skipDepth = 0;
      else if (!skipDepth) metadataParent = metadataParent?.parentNode;
    },
  };
  const parsed = visitXml(bytes, visitor);
  return {
    cells,
    hiddenRows,
    visibleRows,
    rowHeights,
    cellCount,
    formulaCount,
    specialFormulas,
    extended,
    document: parsed.document,
    ...(capture
      ? {
          source: parsed.source,
          cellLocations,
          rowLocations,
          rowCells,
          rowStarts,
          sheetDataStart,
          dimension,
          rootStart,
        }
      : {}),
  };
}
