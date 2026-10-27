/** Shared offline workbook adapter for the editor and MCP. Untouched ZIP members
 * remain byte-identical after decompression; edits touch only relevant XML parts.
 * No workbook is passed through a lossy generic import/export representation. */
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import { createHash } from 'node:crypto'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const encoder = new XMLSerializer()
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex')
const elements = (node, name) =>
  Array.from(node.getElementsByTagNameNS(NS, name))
const children = (node, name) =>
  Array.from(node.childNodes).filter(
    (n) => n.nodeType === 1 && (!name || n.localName === name)
  )
const text = (node, name) => elements(node, name)[0]?.textContent ?? ''
function xml(bytes) {
  const source = strFromU8(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(source))
    throw new Error('XML entities are unsupported')
  return new DOMParser({
    errorHandler: {
      warning() {},
      error(message) {
        throw new Error(message)
      },
      fatalError(message) {
        throw new Error(message)
      }
    }
  }).parseFromString(source, 'application/xml')
}
function node(doc, name, attrs = {}, value) {
  const n = doc.createElementNS(NS, name)
  for (const [key, val] of Object.entries(attrs))
    n.setAttribute(key, String(val))
  if (value !== undefined) n.appendChild(doc.createTextNode(String(value)))
  return n
}
export function columnName(index) {
  let name = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name
  return name
}
export function cellPosition(ref) {
  const m = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(ref)
  if (!m) throw new Error(`Invalid cell address: ${ref}`)
  const col =
    [...m[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) -
    1
  const row = Number(m[2]) - 1
  if (col > 16383 || row > 1048575)
    throw new Error('Cell is outside workbook limits')
  return { row, col }
}
function unpack(bytes) {
  if (bytes.length > 10 * 1024 * 1024)
    throw new Error('Workbook exceeds the 10 MiB editing limit')
  let inflated = 0
  const parts = unzipSync(bytes, {
    filter(entry) {
      inflated += entry.originalSize
      if (inflated > 100 * 1024 * 1024)
        throw new Error('Expanded workbook exceeds the 100 MiB editing limit')
      return true
    }
  })
  if (!parts['xl/workbook.xml'] || !parts['xl/_rels/workbook.xml.rels'])
    throw new Error('Not a supported XLSX workbook')
  return parts
}
function inspect(parts) {
  const workbook = xml(parts['xl/workbook.xml'])
  const rels = xml(parts['xl/_rels/workbook.xml.rels'])
  const targets = Object.fromEntries(
    Array.from(rels.getElementsByTagName('Relationship')).map((r) => {
      const target = r.getAttribute('Target')
      if (target.includes('..') || r.getAttribute('TargetMode') === 'External')
        return [r.getAttribute('Id'), null]
      return [
        r.getAttribute('Id'),
        target.startsWith('/') ? target.slice(1) : `xl/${target}`
      ]
    })
  )
  const reasons = []
  if (
    Object.keys(parts).some((p) =>
      /vbaProject|externalLinks|activeX|embeddings|pivotTables|slicer|connections\.xml|_xmlsignatures/.test(
        p
      )
    )
  )
    reasons.push(
      'Macros, linked data, embedded objects, pivots, or signed workbooks are read-only'
    )
  if (elements(workbook, 'workbookProtection').length)
    reasons.push('Workbook protection is enabled')
  const strings = parts['xl/sharedStrings.xml']
    ? elements(xml(parts['xl/sharedStrings.xml']), 'si').map((s) =>
        elements(s, 't')
          .map((t) => t.textContent)
          .join('')
      )
    : []
  const styles = parts['xl/styles.xml'] ? xml(parts['xl/styles.xml']) : null
  const formats = styles
    ? Object.fromEntries(
        elements(styles, 'numFmt').map((n) => [
          n.getAttribute('numFmtId'),
          n.getAttribute('formatCode')
        ])
      )
    : {}
  const xfs = styles
    ? children(elements(styles, 'cellXfs')[0] ?? styles, 'xf').map((n) => ({
        numFmtId: Number(n.getAttribute('numFmtId')),
        fontId: Number(n.getAttribute('fontId')),
        fillId: Number(n.getAttribute('fillId'))
      }))
    : []
  const fonts = styles
    ? elements(styles, 'font').map((f) => ({
        bold: elements(f, 'b').length > 0,
        italic: elements(f, 'i').length > 0
      }))
    : []
  let cellCount = 0
  const date1904 = ['1', 'true'].includes(
    elements(workbook, 'workbookPr')[0]?.getAttribute('date1904')
  )
  const sheets = elements(workbook, 'sheet').map((sheet) => {
    const path = targets[sheet.getAttribute('r:id')]
    if (!path || !parts[path])
      throw new Error('Workbook contains a missing or external sheet')
    const document = xml(parts[path])
    if (elements(document, 'sheetProtection').length)
      reasons.push(`Sheet “${sheet.getAttribute('name')}” is protected`)
    if (elements(document, 'f').some((f) => f.hasAttribute('t')))
      reasons.push(
        'Shared, array, or data-table formulas require a full spreadsheet application'
      )
    if (elements(document, 'extLst').length)
      reasons.push('Extended worksheet features are read-only')
    const cells = {}
    let rows = 20,
      columns = 8
    for (const c of elements(document, 'c')) {
      const ref = c.getAttribute('r')
      const position = cellPosition(ref)
      rows = Math.max(rows, position.row + 1)
      columns = Math.max(columns, position.col + 1)
      if (++cellCount > 100000)
        throw new Error('Workbook exceeds the 100,000-cell editing limit')
      const type = c.getAttribute('t') || 'n'
      const raw = text(c, 'v')
      const value =
        type === 's'
          ? (strings[Number(raw)] ?? '')
          : type === 'inlineStr'
            ? elements(c, 't')
                .map((t) => t.textContent)
                .join('')
            : type === 'b'
              ? raw === '1'
              : type === 'n' && raw !== ''
                ? Number(raw)
                : raw
      const styleId = Number(c.getAttribute('s') || 0)
      const style = xfs[styleId] ?? {}
      cells[ref] = {
        value,
        formula: text(c, 'f') || undefined,
        styleId,
        format: formats[style.numFmtId] ?? style.numFmtId,
        ...fonts[style.fontId]
      }
    }
    return {
      name: sheet.getAttribute('name'),
      date1904,
      state: sheet.getAttribute('state') || 'visible',
      path,
      cells,
      rows,
      columns,
      merges: elements(document, 'mergeCell').map((m) => m.getAttribute('ref')),
      widths: elements(document, 'col').map((c) => ({
        min: Number(c.getAttribute('min')),
        max: Number(c.getAttribute('max')),
        width: Number(c.getAttribute('width')),
        hidden: c.getAttribute('hidden') === '1'
      })),
      document
    }
  })
  const structuralReason =
    sheets.some((s) =>
      [
        'f',
        'mergeCell',
        'tableParts',
        'drawing',
        'conditionalFormatting',
        'dataValidations',
        'hyperlink'
      ].some((name) => elements(s.document, name).length)
    ) || elements(workbook, 'definedName').length
      ? 'Row/column changes are unavailable for workbooks with formulas, merges, tables, drawings, validation, or named ranges. Cell editing remains available.'
      : undefined
  return {
    workbook,
    sheets,
    styles,
    structuralReason,
    reasons: [...new Set(reasons)]
  }
}
export function readWorkbook(bytes) {
  const state = inspect(unpack(bytes))
  return {
    revision: hash(bytes),
    readOnly: state.reasons.length > 0,
    limitations: state.reasons,
    structuralReason: state.structuralReason,
    calculation:
      'Basic arithmetic and SUM/AVERAGE/MIN/MAX/COUNT display locally. Other formulas retain their expression and recalculate in a compatible spreadsheet application.',
    sheets: state.sheets.map(({ document, ...sheet }) => sheet)
  }
}
export function editWorkbook(bytes, edits, expectedRevision) {
  if (expectedRevision && hash(bytes) !== expectedRevision)
    throw new Error('Conflict: this workbook changed. Reload or save a copy.')
  if (!Array.isArray(edits) || edits.length > 10000)
    throw new Error('Invalid workbook edit batch')
  const parts = unpack(bytes)
  const state = inspect(parts)
  if (state.reasons.length)
    throw new Error(`Read-only workbook: ${state.reasons.join('; ')}`)
  for (const edit of edits) {
    const sheet = state.sheets.find((s) => s.name === edit.sheet)
    if (!sheet) throw new Error(`Sheet not found: ${edit.sheet}`)
    const doc = sheet.document
    if (
      edit.operation === 'insertRow' ||
      edit.operation === 'deleteRow' ||
      edit.operation === 'insertColumn' ||
      edit.operation === 'deleteColumn'
    ) {
      if (state.structuralReason) throw new Error(state.structuralReason)
      const isRow = edit.operation.endsWith('Row'),
        remove = edit.operation.startsWith('delete'),
        index = edit.index
      if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= (isRow ? 1048576 : 16384)
      )
        throw new Error('Invalid row or column')
      for (const c of [...elements(doc, 'c')]) {
        const p = cellPosition(c.getAttribute('r')),
          coordinate = isRow ? p.row : p.col
        if (remove && coordinate === index) {
          c.parentNode.removeChild(c)
          continue
        }
        if (coordinate >= index) {
          if (isRow) p.row += remove ? -1 : 1
          else p.col += remove ? -1 : 1
          c.setAttribute('r', `${columnName(p.col)}${p.row + 1}`)
        }
      }
      if (isRow)
        for (const row of [...elements(doc, 'row')]) {
          const n = Number(row.getAttribute('r')) - 1
          if (remove && n === index) row.parentNode.removeChild(row)
          else if (n >= index)
            row.setAttribute('r', String(n + (remove ? 0 : 2)))
        }
      else
        for (const col of elements(doc, 'col')) {
          for (const key of ['min', 'max']) {
            const n = Number(col.getAttribute(key))
            if (n >= index + 1)
              col.setAttribute(key, String(Math.max(1, n + (remove ? -1 : 1))))
          }
        }
      continue
    }
    const ref = String(edit.cell).toUpperCase(),
      position = cellPosition(ref)
    let data = elements(doc, 'sheetData')[0]
    if (!data) {
      data = node(doc, 'sheetData')
      doc.documentElement.appendChild(data)
    }
    let row = children(data, 'row').find(
      (r) => Number(r.getAttribute('r')) === position.row + 1
    )
    if (!row) {
      row = node(doc, 'row', { r: position.row + 1 })
      data.insertBefore(
        row,
        children(data, 'row').find(
          (r) => Number(r.getAttribute('r')) > position.row + 1
        ) ?? null
      )
    }
    let cell = children(row, 'c').find((c) => c.getAttribute('r') === ref)
    if (!cell) {
      cell = node(doc, 'c', { r: ref })
      row.insertBefore(
        cell,
        children(row, 'c').find(
          (c) => cellPosition(c.getAttribute('r')).col > position.col
        ) ?? null
      )
    }
    if (edit.value !== undefined || edit.formula !== undefined) {
      for (const child of children(cell))
        if (['f', 'v', 'is'].includes(child.localName)) cell.removeChild(child)
      cell.removeAttribute('t')
      if (edit.formula !== undefined) {
        const formula = String(edit.formula).replace(/^=/, '')
        if (
          !formula ||
          formula.length > 8192 ||
          /\[|\]|https?:|file:|DDE\(|WEBSERVICE\(/i.test(formula)
        )
          throw new Error('External or oversized formulas are unsupported')
        cell.appendChild(node(doc, 'f', {}, formula))
      } else if (typeof edit.value === 'number' && Number.isFinite(edit.value))
        cell.appendChild(node(doc, 'v', {}, edit.value))
      else if (typeof edit.value === 'boolean') {
        cell.setAttribute('t', 'b')
        cell.appendChild(node(doc, 'v', {}, edit.value ? '1' : '0'))
      } else {
        cell.setAttribute('t', 'inlineStr')
        const inline = node(doc, 'is')
        const t = node(doc, 't', {}, String(edit.value ?? ''))
        t.setAttribute('xml:space', 'preserve')
        inline.appendChild(t)
        cell.appendChild(inline)
      }
    }
    if (
      edit.bold !== undefined ||
      edit.italic !== undefined ||
      edit.numberFormat !== undefined
    ) {
      if (!state.styles)
        throw new Error('This workbook has no editable style table')
      const styles = state.styles,
        xfs = elements(styles, 'cellXfs')[0],
        fonts = elements(styles, 'fonts')[0]
      const xf = children(xfs, 'xf')[
        Number(cell.getAttribute('s') || 0)
      ]?.cloneNode(true)
      if (!xf) throw new Error('Unsupported cell style')
      if (edit.bold !== undefined || edit.italic !== undefined) {
        const font = children(fonts, 'font')[
          Number(xf.getAttribute('fontId') || 0)
        ]?.cloneNode(true)
        if (!font) throw new Error('Unsupported font style')
        for (const [key, tag] of [
          ['bold', 'b'],
          ['italic', 'i']
        ])
          if (edit[key] !== undefined) {
            for (const old of children(font, tag)) font.removeChild(old)
            if (edit[key]) font.appendChild(node(styles, tag))
          }
        xf.setAttribute('fontId', String(children(fonts, 'font').length))
        xf.setAttribute('applyFont', '1')
        fonts.appendChild(font)
        fonts.setAttribute('count', String(children(fonts, 'font').length))
      }
      if (edit.numberFormat !== undefined) {
        const formats = {
          general: 0,
          number: 2,
          percent: 10,
          date: 14,
          text: 49
        }
        if (!(edit.numberFormat in formats))
          throw new Error('Unsupported number format')
        xf.setAttribute('numFmtId', String(formats[edit.numberFormat]))
        xf.setAttribute('applyNumberFormat', '1')
      }
      cell.setAttribute('s', String(children(xfs, 'xf').length))
      xfs.appendChild(xf)
      xfs.setAttribute('count', String(children(xfs, 'xf').length))
    }
  }
  // Do not display stale calculated results after inputs change. Excel
  // can calculate every formula, while the app explicitly supports a small subset.
  for (const sheet of state.sheets) {
    for (const cell of elements(sheet.document, 'c'))
      if (children(cell, 'f').length)
        for (const v of children(cell, 'v')) cell.removeChild(v)
    for (const dim of elements(sheet.document, 'dimension'))
      dim.parentNode.removeChild(dim)
    parts[sheet.path] = strToU8(encoder.serializeToString(sheet.document))
  }
  let calc = elements(state.workbook, 'calcPr')[0]
  if (!calc) {
    calc = node(state.workbook, 'calcPr')
    state.workbook.documentElement.appendChild(calc)
  }
  calc.setAttribute('fullCalcOnLoad', '1')
  calc.setAttribute('forceFullCalc', '1')
  calc.setAttribute('calcMode', 'auto')
  // Cached dependency order is invalid after a formula edit. Remove the chain
  // together with its package references; the workbook requests full recalc.
  const relationships = xml(parts['xl/_rels/workbook.xml.rels'])
  const contentTypes = xml(parts['[Content_Types].xml'])
  for (const relation of Array.from(
    relationships.getElementsByTagName('Relationship')
  )) {
    if (!relation.getAttribute('Type').endsWith('/calcChain')) continue
    const target = relation.getAttribute('Target')
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`
    delete parts[path]
    relation.parentNode.removeChild(relation)
    for (const type of Array.from(
      contentTypes.getElementsByTagName('Override')
    )) {
      if (type.getAttribute('PartName') === `/${path}`)
        type.parentNode.removeChild(type)
    }
  }
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    encoder.serializeToString(relationships)
  )
  parts['[Content_Types].xml'] = strToU8(
    encoder.serializeToString(contentTypes)
  )
  parts['xl/workbook.xml'] = strToU8(encoder.serializeToString(state.workbook))
  if (state.styles)
    parts['xl/styles.xml'] = strToU8(encoder.serializeToString(state.styles))
  return zipSync(parts)
}
export function createWorkbook() {
  return zipSync(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml':
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        '_rels/.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/_rels/workbook.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${NS}"><sheetData/></worksheet>`,
        'xl/styles.xml': `<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
      }).map(([name, value]) => [name, strToU8(value)])
    )
  )
}
