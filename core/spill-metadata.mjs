import { FORMAT, formatTemplate } from './format.mjs';
/** Closed OOXML writer for qualified native spill ownership. This module cannot
 * qualify a formula or infer followers; its groups must come from the native
 * accepted version. Unrelated metadata records/parts are retained. */
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { rangeBounds } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XDA = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const RICH = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const xml = (source) => {
  if (typeof source !== 'string' || source.length > 8 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(source))
    throw refusal('spill-malformed', 'Invalid spill metadata XML');
  return new DOMParser({
    errorHandler: {
      warning() {},
      error(message) {
        throw refusal('package-malformed', message);
      },
      fatalError(message) {
        throw refusal('package-malformed', message);
      },
    },
  }).parseFromString(source, 'application/xml');
};
const children = (parent, name, ns = parent.namespaceURI) =>
  Array.from(parent.childNodes).filter(
    (child) => child.nodeType === 1 && child.namespaceURI === ns && (!name || child.localName === name)
  );
const integer = (value, label) => {
  if (!/^(?:0|[1-9]\d*)$/.test(value)) throw refusal('internal-invalid-field', `Invalid ${label}`);
  const n = Number(value);
  if (!Number.isSafeInteger(n)) throw refusal('internal-invalid-field', `Invalid ${label}`);
  return n;
};
const signedInteger = (value) => {
  if (!/^-?(?:0|[1-9]\d*)$/.test(value) || !Number.isSafeInteger(Number(value)))
    throw refusal('internal-invalid-field', 'Invalid rich integer');
  return Number(value);
};
const node = (document, ns, name, attributes = {}) => {
  const item = document.createElementNS(ns, name);
  for (const [key, value] of Object.entries(attributes))
    if (key === 'xmlns' || key.startsWith('xmlns:'))
      item.setAttributeNS('http://www.w3.org/2000/xmlns/', key, String(value));
    else item.setAttribute(key, String(value));
  return item;
};
function counted(parent, name, ns = parent.namespaceURI) {
  const items = children(parent, name, ns);
  if (
    children(parent).length !== items.length ||
    Array.from(parent.childNodes).some(
      (child) => child.nodeType === 1 && (child.namespaceURI !== ns || child.localName !== name)
    ) ||
    integer(parent.getAttribute('count'), 'metadata count') !== items.length
  )
    throw refusal('spill-unsupported', 'Inconsistent spill metadata count/children');
  return items;
}
function append(parent, child, name) {
  parent.appendChild(child);
  parent.setAttribute('count', children(parent, name).length);
}
function one(parent, name) {
  const found = children(parent, name);
  if (found.length > 1) throw refusal('spill-ambiguous', `Ambiguous ${name} metadata`);
  return found[0];
}
function extensionLeaf(block, uri, name, namespace) {
  const exact = (parent, local, ns) => {
    const all = Array.from(parent.childNodes).filter((child) => child.nodeType === 1);
    if (
      all.length !== 1 ||
      all[0].localName !== local ||
      all[0].namespaceURI !== ns ||
      Array.from(parent.childNodes).some((child) => child.nodeType === 3 && child.data.trim())
    )
      throw refusal('spill-unsupported', 'Extended spill metadata graph requires qualification');
    return all[0];
  };
  const attributes = (element, allowed) => {
    if (
      Array.from(element.attributes).some(
        (attribute) => attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/' && !allowed.includes(attribute.name)
      )
    )
      throw refusal('spill-unsupported', 'Extended spill metadata attributes require qualification');
  };
  attributes(block, []);
  const list = exact(block, 'extLst', NS);
  attributes(list, []);
  const extension = exact(list, 'ext', NS);
  attributes(extension, ['uri']);
  if (extension.getAttribute('uri') !== uri)
    throw refusal('spill-unsupported', 'Unqualified spill metadata extension URI');
  const leaf = exact(extension, name, namespace);
  attributes(leaf, name === 'rvb' ? ['i'] : ['fDynamic', 'fCollapsed']);
  if (Array.from(leaf.childNodes).some((child) => child.nodeType === 1 || (child.nodeType === 3 && child.data.trim())))
    throw refusal('spill-unsupported', 'Extended spill metadata leaf requires qualification');
  return leaf;
}
const METADATA_TYPE = {
  minSupportedVersion: 120000,
  copy: 1,
  pasteAll: 1,
  pasteValues: 1,
  merge: 1,
  splitFirst: 1,
  rowColShift: 1,
  clearFormats: 1,
  clearComments: 1,
  assign: 1,
  coerce: 1,
};
const RICH_TYPES =
  '<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><global><keyFlags>' +
  [
    '_Self',
    '_DisplayString',
    '_Flags',
    '_Format',
    '_SubLabel',
    '_Attribution',
    '_Icon',
    '_Display',
    '_CanonicalPropertyNames',
    '_ClassificationId',
  ]
    .map(
      (name) =>
        `<key name="${name}">${name === '_Self' ? '<flag name="ExcludeFromFile" value="1"/>' : ''}<flag name="ExcludeFromCalcComparison" value="1"/></key>`
    )
    .join('') +
  '</keyFlags></global></rvTypesInfo>';

export function buildSpillMetadata(metadataBytes, groups) {
  if (!Array.isArray(groups)) throw refusal('internal-missing-field', 'Missing native spill export groups');
  if (!groups.length) return { replacements: {}, masters: {} };
  const original = unzipSync(metadataBytes),
    documents = new Map(),
    changed = new Set(),
    masters = {},
    seen = new Set(),
    rectangles = [];
  function document(part, rootName, namespace) {
    if (!documents.has(part)) {
      const doc = original[part]
        ? xml(strFromU8(original[part]))
        : xml(
            `<${rootName} xmlns="${namespace}"${['metadata', 'rvData', 'rvStructures'].includes(rootName) ? '' : ''}/>`
          );
      if (doc.documentElement.namespaceURI !== namespace || doc.documentElement.localName !== rootName)
        throw refusal('spill-malformed', `Invalid ${part} root`);
      documents.set(part, doc);
      if (!original[part]) changed.add(part);
    }
    return documents.get(part);
  }
  const metadataPart = 'xl/metadata.xml',
    metadata = document(metadataPart, 'metadata', NS),
    root = metadata.documentElement;
  function container(name) {
    let item = one(root, name);
    if (item) return item;
    item = node(metadata, NS, name, { count: 0 });
    const order = ['metadataTypes', 'futureMetadata', 'cellMetadata', 'valueMetadata', 'extLst'],
      rank = order.indexOf(name);
    root.insertBefore(item, children(root).find((child) => order.indexOf(child.localName) > rank) ?? null);
    changed.add(metadataPart);
    return item;
  }
  const types = container('metadataTypes');
  counted(types, 'metadataType');
  function type(name, cellMeta = false) {
    const list = counted(types, 'metadataType'),
      matches = list.map((item, index) => ({ item, index })).filter(({ item }) => item.getAttribute('name') === name);
    if (matches.length > 1) throw refusal('spill-ambiguous', 'Ambiguous spill metadata type');
    if (matches.length) {
      if (cellMeta && matches[0].item.getAttribute('cellMeta') !== '1')
        throw refusal('spill-bad-input', 'XLDAPR is not cell metadata');
      if (!cellMeta && matches[0].item.hasAttribute('cellMeta') && matches[0].item.getAttribute('cellMeta') !== '0')
        throw refusal('spill-unsupported', 'XLRICHVALUE is not qualified value metadata');
      return matches[0].index + 1;
    }
    append(
      types,
      node(metadata, NS, 'metadataType', { name, ...METADATA_TYPE, ...(cellMeta ? { cellMeta: 1 } : {}) }),
      'metadataType'
    );
    changed.add(metadataPart);
    return list.length + 1;
  }
  function future(name) {
    const matches = children(root, 'futureMetadata').filter((item) => item.getAttribute('name') === name);
    if (matches.length > 1) throw refusal('spill-ambiguous', 'Ambiguous future spill metadata');
    if (matches.length) {
      counted(matches[0], 'bk');
      return matches[0];
    }
    const item = node(metadata, NS, 'futureMetadata', { name, count: 0 });
    root.insertBefore(
      item,
      children(root).find((child) => ['cellMetadata', 'valueMetadata', 'extLst'].includes(child.localName)) ?? null
    );
    changed.add(metadataPart);
    return item;
  }
  const dynamicType = type('XLDAPR', true),
    dynamicFuture = future('XLDAPR'),
    cellMetadata = container('cellMetadata');
  counted(cellMetadata, 'bk');
  let dynamicBlock = counted(dynamicFuture, 'bk').findIndex((block) => {
    const flags = extensionLeaf(block, '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}', 'dynamicArrayProperties', XDA);
    if (
      !['0', '1'].includes(flags.getAttribute('fDynamic')) ||
      (flags.hasAttribute('fCollapsed') && !['0', '1'].includes(flags.getAttribute('fCollapsed')))
    )
      throw refusal('internal-invalid-field', 'Invalid XLDAPR dynamic flag');
    return flags.getAttribute('fDynamic') === '1' && flags.getAttribute('fCollapsed') !== '1';
  });
  if (dynamicBlock < 0) {
    const block = node(metadata, NS, 'bk'),
      list = node(metadata, NS, 'extLst'),
      extension = node(metadata, NS, 'ext', { uri: '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}' });
    extension.appendChild(
      node(metadata, XDA, 'xda:dynamicArrayProperties', { 'xmlns:xda': XDA, fDynamic: 1, fCollapsed: 0 })
    );
    list.appendChild(extension);
    block.appendChild(list);
    dynamicBlock = counted(dynamicFuture, 'bk').length;
    append(dynamicFuture, block, 'bk');
    changed.add(metadataPart);
  }
  let cm =
    counted(cellMetadata, 'bk').findIndex((block) => {
      const records = children(block, 'rc');
      return (
        records.length === 1 &&
        integer(records[0].getAttribute('t'), 'cell type') === dynamicType &&
        integer(records[0].getAttribute('v'), 'cell future index') === dynamicBlock
      );
    }) + 1;
  if (!cm) {
    const block = node(metadata, NS, 'bk');
    block.appendChild(node(metadata, NS, 'rc', { t: dynamicType, v: dynamicBlock }));
    append(cellMetadata, block, 'bk');
    cm = counted(cellMetadata, 'bk').length;
    changed.add(metadataPart);
  }
  let rich;
  function spillError(rows, columns) {
    if (!rich) {
      const structurePart = 'xl/richData/rdrichvaluestructure.xml',
        valuePart = 'xl/richData/rdrichvalue.xml';
      const structures = document(structurePart, 'rvStructures', RICH),
        values = document(valuePart, 'rvData', RICH);
      for (const doc of [structures, values])
        if (!doc.documentElement.hasAttribute('count')) {
          if (original[doc === structures ? structurePart : valuePart])
            throw refusal('spill-unsupported', 'Rich data has no count');
          doc.documentElement.setAttribute('count', '0');
        }
      const fields = ['colOffset', 'errorType', 'rwOffset', 'subType'];
      let structure = counted(structures.documentElement, 's').findIndex(
        (item) =>
          item.getAttribute('t') === '_error' &&
          children(item, 'k').length === fields.length &&
          children(item, 'k').every(
            (key, index) => key.getAttribute('n') === fields[index] && key.getAttribute('t') === 'i'
          )
      );
      if (structure < 0) {
        const item = node(structures, RICH, 's', { t: '_error' });
        for (const name of fields) item.appendChild(node(structures, RICH, 'k', { n: name, t: 'i' }));
        structure = counted(structures.documentElement, 's').length;
        append(structures.documentElement, item, 's');
        changed.add(structurePart);
      }
      counted(values.documentElement, 'rv');
      const richType = type('XLRICHVALUE'),
        richFuture = future('XLRICHVALUE'),
        valueMetadata = container('valueMetadata');
      counted(valueMetadata, 'bk');
      if (!original['xl/richData/rdRichValueTypes.xml']) {
        documents.set('xl/richData/rdRichValueTypes.xml', xml(RICH_TYPES));
        changed.add('xl/richData/rdRichValueTypes.xml');
      }
      rich = { structures, values, structure, richType, richFuture, valueMetadata, structurePart, valuePart };
    }
    const expected = [columns - 1, 8, rows - 1, 1],
      { values, structure, richType, richFuture, valueMetadata, valuePart } = rich;
    let valueIndex = counted(values.documentElement, 'rv').findIndex(
      (item) =>
        integer(item.getAttribute('s'), 'rich structure') === structure &&
        children(item, 'v').length === expected.length &&
        children(item, 'v').every((v, index) => signedInteger(v.textContent) === expected[index])
    );
    if (valueIndex < 0) {
      const item = node(values, RICH, 'rv', { s: structure });
      for (const value of expected) {
        const child = node(values, RICH, 'v');
        child.appendChild(values.createTextNode(String(value)));
        item.appendChild(child);
      }
      valueIndex = counted(values.documentElement, 'rv').length;
      append(values.documentElement, item, 'rv');
      changed.add(valuePart);
    }
    let futureIndex = counted(richFuture, 'bk').findIndex(
      (block) =>
        integer(
          extensionLeaf(block, '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}', 'rvb', RICH).getAttribute('i'),
          'rich value index'
        ) === valueIndex
    );
    if (futureIndex < 0) {
      const block = node(metadata, NS, 'bk'),
        list = node(metadata, NS, 'extLst'),
        extension = node(metadata, NS, 'ext', { uri: '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}' });
      extension.appendChild(node(metadata, RICH, 'xlrd:rvb', { 'xmlns:xlrd': RICH, i: valueIndex }));
      list.appendChild(extension);
      block.appendChild(list);
      futureIndex = counted(richFuture, 'bk').length;
      append(richFuture, block, 'bk');
      changed.add(metadataPart);
    }
    let vm =
      counted(valueMetadata, 'bk').findIndex((block) => {
        const records = children(block, 'rc');
        return (
          records.length === 1 &&
          integer(records[0].getAttribute('t'), 'value type') === richType &&
          integer(records[0].getAttribute('v'), 'value future index') === futureIndex
        );
      }) + 1;
    if (!vm) {
      const block = node(metadata, NS, 'bk');
      block.appendChild(node(metadata, NS, 'rc', { t: richType, v: futureIndex }));
      append(valueMetadata, block, 'bk');
      vm = counted(valueMetadata, 'bk').length;
      changed.add(metadataPart);
    }
    return vm;
  }
  for (const group of groups) {
    if (
      group.schema !== 1 ||
      typeof group.sheet !== 'string' ||
      typeof group.master !== 'string' ||
      !['calculated', 'error'].includes(group.status)
    )
      throw refusal('internal-invalid-field', 'Invalid native spill export group');
    const key = `${group.sheet}\0${group.master}`;
    if (seen.has(key)) throw refusal('spill-duplicate', 'Duplicate native spill master');
    seen.add(key);
    const bounds = rangeBounds(group.range);
    if (bounds.size > 1000 || group.master !== group.range.split(':')[0])
      throw refusal('internal-invalid-field', 'Invalid spill output geometry');
    if (group.status === 'calculated') {
      for (const prior of rectangles)
        if (
          prior.sheet === group.sheet &&
          prior.bounds.first.row <= bounds.last.row &&
          bounds.first.row <= prior.bounds.last.row &&
          prior.bounds.first.col <= bounds.last.col &&
          bounds.first.col <= prior.bounds.last.col
        )
          throw refusal('spill-conflict', 'Overlapping native spill output geometry');
      rectangles.push({ sheet: group.sheet, bounds });
      masters[key] = { cm, range: group.range };
    } else {
      if (group.error !== '#SPILL!' || bounds.size !== 1)
        throw refusal('spill-unsupported', 'Unqualified modern spill error');
      const intended = rangeBounds(group.intendedRange);
      if (intended.size > 1000 || intended.first.row !== bounds.first.row || intended.first.col !== bounds.first.col)
        throw refusal('internal-invalid-field', 'Invalid blocked spill geometry');
      masters[key] = {
        cm,
        vm: spillError(intended.rows, intended.columns),
        range: group.master,
        intendedRange: group.intendedRange,
        aca: '1',
        ca: '1',
      };
    }
  }
  function relationship(part, type, target) {
    const doc = document('xl/_rels/workbook.xml.rels', 'Relationships', REL),
      list = children(doc.documentElement, 'Relationship');
    const existingIds = list.map((item) => item.getAttribute('Id'));
    if (existingIds.some((id) => !id) || new Set(existingIds).size !== existingIds.length)
      throw refusal('spill-ambiguous', 'Ambiguous spill relationship identifiers');
    const matches = list.filter((item) => item.getAttribute('Type') === type);
    if (
      matches.length > 1 ||
      (matches.length &&
        ((matches[0].getAttribute('TargetMode') || 'Internal') !== 'Internal' ||
          matches[0].getAttribute('Target') !== target))
    )
      throw refusal('spill-unsupported', 'Unsupported spill relationship mapping');
    if (original[part] && !matches.length)
      throw refusal('spill-unsupported', 'Existing spill part has no qualified relationship');
    if (!matches.length) {
      const ids = new Set(list.map((item) => item.getAttribute('Id')));
      let index = 1;
      while (ids.has(formatTemplate([FORMAT.SPILL, FORMAT.FORMAT_E3B0C4], [index]))) index++;
      doc.documentElement.appendChild(
        node(doc, REL, 'Relationship', {
          Id: formatTemplate([FORMAT.SPILL, FORMAT.FORMAT_E3B0C4], [index]),
          Type: type,
          Target: target,
        })
      );
      changed.add('xl/_rels/workbook.xml.rels');
    }
    const types = document('[Content_Types].xml', 'Types', CT),
      allOverrides = children(types.documentElement, 'Override'),
      overrideNames = allOverrides.map((item) => item.getAttribute('PartName'));
    if (overrideNames.some((name) => !name) || new Set(overrideNames).size !== overrideNames.length)
      throw refusal('spill-ambiguous', 'Ambiguous spill content type identifiers');
    const overrides = allOverrides.filter((item) => item.getAttribute('PartName') === `/${part}`),
      contentType = {
        'xl/metadata.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
        'xl/richData/rdrichvalue.xml': 'application/vnd.ms-excel.rdrichvalue+xml',
        'xl/richData/rdrichvaluestructure.xml': 'application/vnd.ms-excel.rdrichvaluestructure+xml',
        'xl/richData/rdRichValueTypes.xml': 'application/vnd.ms-excel.rdrichvaluetypes+xml',
      }[part];
    if (overrides.length > 1 || (overrides.length && overrides[0].getAttribute('ContentType') !== contentType))
      throw refusal('spill-unsupported', 'Unsupported spill content type');
    if (original[part] && !overrides.length)
      throw refusal('spill-unsupported', 'Existing spill part has no qualified content type');
    if (!overrides.length) {
      types.documentElement.appendChild(
        node(types, CT, 'Override', { PartName: `/${part}`, ContentType: contentType })
      );
      changed.add('[Content_Types].xml');
    }
  }
  relationship(
    'xl/metadata.xml',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata',
    'metadata.xml'
  );
  if (rich)
    for (const [part, type] of [
      ['rdrichvalue.xml', 'rdRichValue'],
      ['rdrichvaluestructure.xml', 'rdRichValueStructure'],
      ['rdRichValueTypes.xml', 'rdRichValueTypes'],
    ])
      relationship(
        `xl/richData/${part}`,
        `http://schemas.microsoft.com/office/2017/06/relationships/${type}`,
        `richData/${part}`
      );
  const serializer = new XMLSerializer(),
    replacements = {};
  for (const part of changed) replacements[part] = serializer.serializeToString(documents.get(part));
  return { replacements, masters };
}
