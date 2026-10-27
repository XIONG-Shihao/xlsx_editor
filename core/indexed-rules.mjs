import { formatHashSalt, FORMAT } from './format.mjs';
/** Guarded collection replacement. Existing DXFs are immutable and retained. */
import { lineStarts } from './xml-position.mjs';
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8, strToU8 } from 'fflate';
import { openZipParts } from './zip-parts.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import {
  validateConditionalRules,
  validateValidationRules,
  rangesOverlap,
  acceptedRuleValue,
  duplicateRuleKey,
} from './rules-model.mjs';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import {
  ruleChildren,
  readConditionalRules,
  readValidationRules,
  conditionalRulesXml,
  validationRulesXml,
  differentialRulesXml,
} from './rules-xml.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function parse(bytes, root) {
  if (!bytes || bytes.byteLength > 16 * 1024 ** 2) throw refusal('rules-limit', 'Rule metadata exceeds its bound');
  const source = typeof bytes === 'string' ? bytes : strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('rules-malformed', 'Rule XML entities are protected');
  const errors = [],
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (value) => value,
      errorHandler: {
        warning: (value) => errors.push(value),
        error: (value) => errors.push(value),
        fatalError: (value) => errors.push(value),
      },
    }).parseFromString(source, 'application/xml');
  if (errors.length || doc.documentElement.namespaceURI !== NS || doc.documentElement.localName !== root)
    throw refusal('rules-malformed', 'Invalid rule XML');
  const lines = lineStarts(source);
  const raw = (node) => {
    const start = lines[node.lineNumber - 1] + node.columnNumber - 1;
    if (!Number.isSafeInteger(start) || !source.startsWith('<' + node.tagName, start))
      throw refusal('rules-unsupported', 'Unqualified rule raw boundary');
    const end = xmlElementEnd(source, start);
    return { start, end, raw: source.slice(start, end) };
  };
  return { source, doc, raw };
}
function context(bytes, view, sheetName, kind, stylesXml) {
  const sheet = view.sheets.find((item) => item.name === sheetName);
  if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
  const scope = view.sheets.indexOf(sheet),
    parts = openZipParts(bytes, 16 * 1024 ** 2),
    worksheet = parse(parts[sheet.path], 'worksheet'),
    styles = parts['xl/styles.xml'] ? parse(parts['xl/styles.xml'], 'styleSheet') : null,
    tag = kind === 'conditional' ? 'conditionalFormatting' : 'dataValidations',
    nodes = ruleChildren(worksheet.doc.documentElement).filter((node) => node.localName === tag);
  if (nodes.some((node) => node.namespaceURI !== NS)) throw refusal('rules-protected', 'Foreign rules are protected');
  const dxfs = styles ? ruleChildren(styles.doc.documentElement).filter((node) => node.localName === 'dxfs') : [];
  if (dxfs.length > 1) throw refusal('rules-ambiguous', 'Ambiguous differential style collection');
  const raw = [
      ...nodes.map((node) => worksheet.raw(node).raw),
      ...(kind === 'conditional' && dxfs.length ? [styles.raw(dxfs[0]).raw] : []),
    ],
    revision = nodes.length
      ? bytesToHex(
          sha256(strToU8([formatHashSalt('RULES_V1'), kind, sheet.path, sheet.name, String(scope), ...raw].join('\0')))
        )
      : null;
  return { sheet, scope, kind, worksheet, styles, nodes, revision, stylesXml };
}
export function indexedRulesDefinition(bytes, view, sheetName, kind) {
  const ctx = context(bytes, view, sheetName, kind);
  try {
    if (ruleChildren(ctx.worksheet.doc.documentElement).some((node) => node.localName === 'extLst'))
      throw refusal('rules-unsupported', 'Extended worksheet rules require qualification');
    const rules =
      kind === 'conditional'
        ? readConditionalRules(ctx.worksheet.doc, ctx.styles?.doc)
        : readValidationRules(ctx.worksheet.doc);
    return { sheet: sheetName, revision: ctx.revision, canEdit: true, protectedReasons: [], rules };
  } catch (error) {
    let protectedRanges;
    try {
      const containers = ruleChildren(ctx.worksheet.doc.documentElement).filter(
        (node) => node.localName === 'dataValidations'
      );
      if (kind === 'validation' && containers.length === 1 && containers[0].namespaceURI === NS) {
        const nodes = ruleChildren(containers[0]);
        if (nodes.length > 128) throw refusal('rules-limit', 'Too many protected ranges');
        protectedRanges = nodes.flatMap((node) => {
          if (node.namespaceURI !== NS || node.localName !== 'dataValidation')
            throw refusal('rules-unsupported', 'Foreign validation');
          return node
            .getAttribute('sqref')
            .split(/\s+/)
            .map((range) => {
              rangeBounds(range);
              return range;
            });
        });
      }
    } catch {
      protectedRanges = undefined;
    }
    return {
      sheet: sheetName,
      revision: ctx.revision,
      canEdit: false,
      protectedReasons: [String(error.message ?? error)],
      ...(protectedRanges ? { protectedRanges } : {}),
      rules: [],
    };
  }
}
export function isIndexedRulesCommand(command) {
  return ['setConditionalRules', 'setValidationRules'].includes(command?.operation);
}
function appendDxfs(styles, rules) {
  const root = styles.doc.documentElement,
    prefix = root.prefix ? root.prefix + ':' : '',
    containers = ruleChildren(root).filter((node) => node.localName === 'dxfs');
  if (containers.length > 1) throw refusal('rules-ambiguous', 'Ambiguous differential style collection');
  const container = containers[0],
    old = container ? ruleChildren(container).length : 0;
  if (old + rules.length > 4096) throw refusal('rules-limit', 'Differential style collection exceeds its bound');
  if (
    container &&
    (container.namespaceURI !== NS ||
      Array.from(container.attributes).some(
        (attr) => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/' && (attr.namespaceURI || attr.name !== 'count')
      ) ||
      ruleChildren(container).some((node) => node.localName !== 'dxf' || node.namespaceURI !== NS))
  )
    throw refusal('rules-protected', 'Extended differential style collection is protected');
  if (container?.hasAttribute('count') && Number(container.getAttribute('count')) !== old)
    throw refusal('internal-invalid-field', 'Invalid differential style count');
  if (!rules.length) return { xml: styles.source, old };
  const added = differentialRulesXml(rules, prefix);
  let xml;
  if (container) {
    const item = styles.raw(container),
      raw = item.raw;
    const start = raw
      .slice(0, raw.indexOf('>') + 1)
      .replace(/\s+count\s*=\s*(['"])[\s\S]*?\1/, '')
      .replace(/\/?\s*>$/, ' count="' + (old + rules.length) + '">');
    const body = /\/\s*>$/.test(raw) ? '' : raw.slice(raw.indexOf('>') + 1, raw.lastIndexOf('</'));
    xml =
      styles.source.slice(0, item.start) +
      start +
      body +
      added +
      '</' +
      container.tagName +
      '>' +
      styles.source.slice(item.end);
  } else {
    const following = ruleChildren(root).find((node) => ['tableStyles', 'colors', 'extLst'].includes(node.localName)),
      at = following ? styles.raw(following).start : styles.source.lastIndexOf('</');
    xml =
      styles.source.slice(0, at) +
      '<' +
      prefix +
      'dxfs count="' +
      rules.length +
      '">' +
      added +
      '</' +
      prefix +
      'dxfs>' +
      styles.source.slice(at);
  }
  return { xml, old };
}
export function indexedRulesPlan(bytes, view, command, { stylesXml } = {}) {
  if (
    !command ||
    Object.keys(command).some((key) => !['operation', 'sheet', 'expectedRulesRevision', 'rules'].includes(key)) ||
    !isIndexedRulesCommand(command)
  )
    throw refusal('internal-invalid-field', 'Invalid rule command');
  const kind = command.operation === 'setConditionalRules' ? 'conditional' : 'validation',
    ctx = context(bytes, view, command.sheet, kind),
    previous = indexedRulesDefinition(bytes, view, command.sheet, kind);
  if (!previous.canEdit)
    throw refusal('rules-protected', 'Rules are protected: ' + previous.protectedReasons.join('; '));
  if (command.expectedRulesRevision !== ctx.revision) throw refusal('rules-stale', 'Conflict: rule definition changed');
  const rules =
      kind === 'conditional' ? validateConditionalRules(command.rules) : validateValidationRules(command.rules),
    prefix = ctx.worksheet.doc.documentElement.prefix ? ctx.worksheet.doc.documentElement.prefix + ':' : '',
    replacements = {};
  let fragment;
  if (kind === 'conditional') {
    if (rules.length && !ctx.styles)
      throw refusal('rules-unsupported', 'Conditional authoring requires an existing styles table');
    const styles = stylesXml === undefined ? ctx.styles : parse(stylesXml, 'styleSheet'),
      result = styles ? appendDxfs(styles, rules) : { old: 0 };
    if (rules.length) replacements['xl/styles.xml'] = result.xml;
    fragment = conditionalRulesXml(rules, prefix, result.old);
  } else fragment = validationRulesXml(rules, prefix);
  const plan = {
    kind: command.operation,
    sheetPart: ctx.sheet.path,
    sheetName: ctx.sheet.name,
    scope: ctx.scope,
    priorRulesRevision: ctx.revision,
    rules,
    replacements,
    removed: [],
    copies: {},
    fragments: {
      [ctx.sheet.path]: { [kind === 'conditional' ? 'conditionalFormatting' : 'dataValidations']: fragment },
    },
  };
  if (strToU8(JSON.stringify(plan)).byteLength > 2 * 1024 ** 2)
    throw refusal('rules-limit', 'Rule plan exceeds its transport bound');
  return {
    plan,
    command: { operation: command.operation, sheet: ctx.sheet.name, expectedRulesRevision: ctx.revision, rules },
  };
}
export function indexedValidationMutationGuard(bytes, view, command) {
  const sheet = command.sheet ?? command.shape?.sheet;
  if (!sheet) return null;
  const definition = indexedRulesDefinition(bytes, view, sheet, 'validation');
  const references = ['renameSheet', 'duplicateSheet', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'];
  if (references.includes(command.operation)) return null;
  if (command.operation === 'filterRows') return null;
  // Validation describes an area; a sort leaves it in place, as in Excel.
  if (command.operation === 'sortRows') return null;
  if (
    command.type === 'formatRange' ||
    (command.type === 'fillAcrossSheets' && command.mode === 'formats') ||
    isIndexedRulesCommand(command) ||
    [
      'setPrintSettings',
      'createChart',
      'updateChart',
      'insertImage',
      'deleteObject',
      'defineName',
      'removeName',
      'resizeRow',
      'resizeColumn',
      'freeze',
      'mergeCells',
      'unmergeCells',
    ].includes(command.operation)
  )
    return null;
  // A dropdown or validation rule the editor cannot evaluate is kept exactly as
  // it is and does not stop editing. Excel treats validation the same way for
  // values that arrive by paste or fill: the rule stays, the value is accepted.
  if (!definition.canEdit) return null;
  if (!definition.rules.length) return null;
  if (
    command.type === 'clearRange' &&
    definition.rules.filter((rule) => rangesOverlap(rule.range, command.range)).every((rule) => rule.allowBlank)
  )
    return null;
  if (
    [
      'importCells',
      'fillRange',
      'fillAcrossSheets',
      'pasteFormulas',
      'fillSeries',
      'pasteValues',
      'clearRange',
    ].includes(command.type)
  ) {
    const range = command.range ?? command.shape?.start + ':' + command.shape?.start;
    if (command.type === 'importCells' || definition.rules.some((rule) => rangesOverlap(rule.range, range)))
      return 'Bulk editing in validated areas is not qualified; use explicit literal edits or clear validation first';
  }
  return null;
}
/** Complete duplicate counts use accepted source pages, never viewport cells.
 * Bound both source positions and distinct serialized keys before retaining. */
export async function indexedRuleDuplicates(rules, sheet, readRange, check = () => {}) {
  const result = {};
  let positions = 0,
    bytes = 0;
  for (const rule of rules) {
    if (rule.kind !== 'duplicateValues' || result[rule.range]) continue;
    const bounds = rangeBounds(rule.range);
    positions += bounds.size;
    if (positions > 10000)
      throw refusal('rules-duplicate', 'Duplicate rule reads are limited to 10,000 total source positions');
    const counts = {},
      columns = Math.min(1000, bounds.columns),
      rows = Math.max(1, Math.floor(1000 / columns));
    for (let row = bounds.first.row; row <= bounds.last.row; row += rows)
      for (let col = bounds.first.col; col <= bounds.last.col; col += columns) {
        check();
        const lastRow = Math.min(bounds.last.row, row + rows - 1),
          lastCol = Math.min(bounds.last.col, col + columns - 1),
          reply = await readRange(sheet, columnName(col) + (row + 1) + ':' + columnName(lastCol) + (lastRow + 1));
        check();
        for (let r = row; r <= lastRow; r++)
          for (let c = col; c <= lastCol; c++) {
            const record = acceptedRuleValue(reply.cells[columnName(c) + (r + 1)]);
            if (!record.qualified)
              throw refusal('rules-duplicate', 'Duplicate rule source contains unresolved, special or unsafe values');
            const key = duplicateRuleKey(record.value);
            if (!Object.hasOwn(counts, key)) {
              bytes += strToU8(key).byteLength + 24;
              if (bytes > 1024 * 1024)
                throw refusal('rules-duplicate', 'Duplicate values exceed the bounded metadata reply');
            }
            counts[key] = (counts[key] ?? 0) + 1;
          }
      }
    result[rule.range] = Object.fromEntries(Object.entries(counts).filter(([, count]) => count > 1));
  }
  return result;
}
