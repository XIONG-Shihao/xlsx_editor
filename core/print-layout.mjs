import { chartSeriesAppearance } from './chart-styles.mjs';
import { chartPreviewGeometry } from './chart-layout.mjs';
/** Physical page composition from a bounded immutable accepted snapshot. */
import { rangeBounds, columnName, cellPosition } from './workbook-commands.mjs';
import { PRINT_LIMITS, printPaperSize, validatePrintSettings } from './print-model.mjs';
import { refusal } from './refusal-codes.mjs';
const bytes = (text) => new TextEncoder().encode(text).byteLength;
const sum = (items, key) => items.reduce((total, item) => total + item[key], 0);
function refuse(code, details = {}) {
  const error = Error(code.replaceAll('_', ' '));
  error.code = code;
  Object.assign(error, details);
  throw error;
}
function chunks(items, key, capacity, repeat) {
  const header = sum(repeat, key);
  if (header >= capacity - 0.001) refuse('PRINT_REPEAT_HEADER_TOO_LARGE');
  const repeated = new Set(repeat.map((item) => item.index)),
    body = items.filter((item) => !repeated.has(item.index)),
    parts = [];
  let part = [],
    used = header;
  for (const item of body) {
    if (item[key] > capacity - header + 0.001) refuse('PRINT_CELL_PAGE_CUT', { index: item.index });
    if (part.length && used + item[key] > capacity + 0.001) {
      parts.push([...repeat, ...part]);
      part = [];
      used = header;
    }
    part.push(item);
    used += item[key];
  }
  if (part.length || !parts.length) parts.push([...repeat, ...part]);
  return parts;
}
const intersects = (a, b) => a[0] < b[0] + b[2] && a[0] + a[2] > b[0] && a[1] < b[1] + b[3] && a[1] + a[3] > b[1];
export function composePrint(snapshot) {
  if (
    snapshot?.schema !== 1 ||
    typeof snapshot.fingerprint !== 'string' ||
    snapshot.cells.length > PRINT_LIMITS.sourcePositions ||
    bytes(JSON.stringify(snapshot)) > PRINT_LIMITS.snapshotBytes
  )
    refuse('PRINT_SNAPSHOT_BOUND');
  const settings = validatePrintSettings(snapshot.settings),
    bounds = rangeBounds(settings.area),
    size = printPaperSize(settings),
    margin = settings.marginsPt,
    usable = [size[0] - margin.left - margin.right, size[1] - margin.top - margin.bottom],
    rows = snapshot.rows.filter((item) => !item.hidden),
    columns = snapshot.columns.filter((item) => !item.hidden && item.widthPt > 0);
  if (!rows.length || !columns.length) refuse('PRINT_NO_VISIBLE_CELLS');
  for (const item of rows)
    if (!Number.isSafeInteger(item.index) || !Number.isFinite(item.heightPt) || item.heightPt <= 0)
      refuse('PRINT_INVALID_GEOMETRY');
  for (const item of columns)
    if (!Number.isSafeInteger(item.index) || !Number.isFinite(item.widthPt) || item.widthPt <= 0)
      refuse('PRINT_INVALID_GEOMETRY');
  const repeatRows = settings.repeatRows
      ? rows.filter((item) => item.index + 1 >= settings.repeatRows.first && item.index + 1 <= settings.repeatRows.last)
      : [],
    repeatColumns = settings.repeatColumns
      ? columns.filter(
          (item) =>
            item.index >= cellPosition(settings.repeatColumns.first + '1').col &&
            item.index <= cellPosition(settings.repeatColumns.last + '1').col
        )
      : [],
    totalWidth = sum(columns, 'widthPt'),
    totalHeight = sum(rows, 'heightPt');
  let scale;
  if (settings.scaling.mode === 'scale') scale = settings.scaling.percent / 100;
  else {
    const fit = settings.scaling,
      fits = (value) => {
        try {
          return (
            chunks(columns, 'widthPt', usable[0] / value, repeatColumns).length <= fit.width &&
            (!fit.height || chunks(rows, 'heightPt', usable[1] / value, repeatRows).length <= fit.height)
          );
        } catch (error) {
          if (['PRINT_REPEAT_HEADER_TOO_LARGE', 'PRINT_CELL_PAGE_CUT'].includes(error.code)) return false;
          throw error;
        }
      };
    // Whole rows/columns can waste the fractional space implied by total-area
    // division. Solve actual pagination, preserving repeats, at one uniform scale.
    // Fit shrinks an area when required; explicit percent mode can enlarge it.
    if (!fits(0.1)) refuse('PRINT_FIT_TOO_SMALL');
    let low = 0.1,
      high = 1;
    if (fits(high)) scale = high;
    else {
      for (let step = 0; step < 52; step++) {
        const middle = (low + high) / 2;
        if (fits(middle)) low = middle;
        else high = middle;
      }
      scale = low;
    }
  }
  if (settings.headerFooter) {
    const hf = settings.headerFooter,
      height = 12 * (hf.scaleWithDoc ? scale : 1);
    for (const side of ['header', 'footer'])
      if (Object.values(hf[side]).some(Boolean)) {
        const top = side === 'header' ? margin.header : size[1] - margin.footer - height;
        if (
          top < 0 ||
          (side === 'header' && top + height > margin.top) ||
          (side === 'footer' && top < size[1] - margin.bottom)
        )
          refuse('PRINT_HEADER_FOOTER_MARGIN');
      }
  }
  const rowPages = chunks(rows, 'heightPt', usable[1] / scale, repeatRows),
    columnPages = chunks(columns, 'widthPt', usable[0] / scale, repeatColumns);
  if (rowPages.length * columnPages.length > PRINT_LIMITS.pages) refuse('PRINT_PAGE_LIMIT');
  const merges = (snapshot.merges ?? []).map((range) => ({ range, bounds: rangeBounds(range) }));
  for (const merge of merges) {
    const b = merge.bounds;
    if (
      b.first.row < bounds.first.row ||
      b.last.row > bounds.last.row ||
      b.first.col < bounds.first.col ||
      b.last.col > bounds.last.col
    )
      refuse('PRINT_PARTIAL_MERGE', { range: merge.range });
  }
  const cells = new Map(snapshot.cells.map((cell) => [cell.ref, cell]));
  if (cells.size !== snapshot.cells.length) refuse('PRINT_DUPLICATE_CELL');
  const sourceRows = new Map(),
    sourceColumns = new Map();
  let at = 0;
  for (const row of rows) {
    sourceRows.set(row.index, at);
    at += row.heightPt;
  }
  at = 0;
  for (const column of columns) {
    sourceColumns.set(column.index, at);
    at += column.widthPt;
  }
  const pages = [],
    boundaries = [];
  let rendered = 0,
    layoutBytes = 0;
  for (const columnGroup of columnPages)
    for (const rowGroup of rowPages) {
      const rowSet = new Set(rowGroup.map((item) => item.index)),
        columnSet = new Set(columnGroup.map((item) => item.index)),
        y = new Map(),
        x = new Map();
      let offset = margin.top;
      for (const row of rowGroup) {
        y.set(row.index, offset);
        offset += row.heightPt * scale;
      }
      offset = margin.left;
      for (const column of columnGroup) {
        x.set(column.index, offset);
        offset += column.widthPt * scale;
      }
      const page = {
        number: pages.length + 1,
        rowIndices: [...rowSet],
        columnIndices: [...columnSet],
        cells: [],
        objects: [],
      };
      const span = (items) => (items.length ? [items[0].index, items.at(-1).index] : null);
      const bodyRows = rowGroup.filter((row) => !repeatRows.some((repeat) => repeat.index === row.index)),
        bodyColumns = columnGroup.filter((column) => !repeatColumns.some((repeat) => repeat.index === column.index));
      boundaries.push({
        number: page.number,
        rows: span(rowGroup),
        columns: span(columnGroup),
        bodyRows: span(bodyRows),
        bodyColumns: span(bodyColumns),
      });

      for (const merge of merges) {
        const visibleRows = rows.filter(
            (item) => item.index >= merge.bounds.first.row && item.index <= merge.bounds.last.row
          ),
          visibleCols = columns.filter(
            (item) => item.index >= merge.bounds.first.col && item.index <= merge.bounds.last.col
          ),
          some =
            visibleRows.some((item) => rowSet.has(item.index)) && visibleCols.some((item) => columnSet.has(item.index));
        if (
          some &&
          (!visibleRows.every((item) => rowSet.has(item.index)) ||
            !visibleCols.every((item) => columnSet.has(item.index)))
        )
          refuse('PRINT_MERGE_PAGE_CUT', { range: merge.range });
        if (
          some &&
          (visibleRows[0]?.index !== merge.bounds.first.row || visibleCols[0]?.index !== merge.bounds.first.col)
        )
          refuse('PRINT_HIDDEN_MERGE_MASTER', { range: merge.range });
      }
      for (const row of rowGroup)
        for (const column of columnGroup) {
          const merge = merges.find(
            (item) =>
              row.index >= item.bounds.first.row &&
              row.index <= item.bounds.last.row &&
              column.index >= item.bounds.first.col &&
              column.index <= item.bounds.last.col
          );
          if (merge && (row.index !== merge.bounds.first.row || column.index !== merge.bounds.first.col)) continue;
          const ref = columnName(column.index) + (row.index + 1),
            source = cells.get(ref);
          if (!source) refuse('PRINT_INCOMPLETE_SNAPSHOT', { cell: ref });
          const width = merge
              ? sum(
                  columnGroup.filter(
                    (item) => item.index >= merge.bounds.first.col && item.index <= merge.bounds.last.col
                  ),
                  'widthPt'
                )
              : column.widthPt,
            height = merge
              ? sum(
                  rowGroup.filter(
                    (item) => item.index >= merge.bounds.first.row && item.index <= merge.bounds.last.row
                  ),
                  'heightPt'
                )
              : row.heightPt,
            cell = {
              ref,
              text: source.text,
              style: source.style,
              bbox: [x.get(column.index), y.get(row.index), width * scale, height * scale],
              ...(merge ? { merge: merge.range } : {}),
            };
          if (++rendered > PRINT_LIMITS.renderedPositions) refuse('PRINT_RENDERED_POSITION_LIMIT');
          layoutBytes += bytes(JSON.stringify(cell));
          if (layoutBytes > PRINT_LIMITS.layoutBytes) refuse('PRINT_LAYOUT_BOUND');
          page.cells.push(cell);
        }
      // A page is a union of contiguous source row/column runs. Repeated prefixes
      // are distinct runs on later pages; each object appears only where its frozen
      // source box intersects the corresponding run's physical clipping area.
      function runs(items, key, positions) {
        const result = [];
        let run = [];
        for (const item of items) {
          if (
            run.length &&
            Math.abs(positions.get(item.index) - positions.get(run.at(-1).index) - run.at(-1)[key]) > 0.000001
          ) {
            result.push(run);
            run = [];
          }
          run.push(item);
        }
        if (run.length) result.push(run);
        return result;
      }
      for (const object of snapshot.objects ?? [])
        for (const rowRun of runs(rowGroup, 'heightPt', sourceRows))
          for (const colRun of runs(columnGroup, 'widthPt', sourceColumns)) {
            const sourceBox = [
              sourceColumns.get(colRun[0].index),
              sourceRows.get(rowRun[0].index),
              sum(colRun, 'widthPt'),
              sum(rowRun, 'heightPt'),
            ];
            if (!intersects(object.boxPt, sourceBox)) continue;
            if (object.protectedReason || !object.anchor || (!object.data && !object.image))
              refuse('PRINT_UNSUPPORTED_OBJECT', { objectId: object.id });
            if (page.objects.some((item) => item.id === object.id))
              refuse('PRINT_OBJECT_REPEAT_CUT', { objectId: object.id });
            const clip = [x.get(colRun[0].index), y.get(rowRun[0].index), sourceBox[2] * scale, sourceBox[3] * scale],
              box = [
                clip[0] + (object.boxPt[0] - sourceBox[0]) * scale,
                clip[1] + (object.boxPt[1] - sourceBox[1]) * scale,
                object.boxPt[2] * scale,
                object.boxPt[3] * scale,
              ],
              clipped =
                box[0] < clip[0] ||
                box[1] < clip[1] ||
                box[0] + box[2] > clip[0] + clip[2] ||
                box[1] + box[3] > clip[1] + clip[3];
            page.objects.push({ id: object.id, kind: object.kind, bbox: box, clipBox: clip, clipped, object });
            layoutBytes += bytes(JSON.stringify(page.objects.at(-1)));
            if (layoutBytes > PRINT_LIMITS.layoutBytes) refuse('PRINT_LAYOUT_BOUND');
          }
      pages.push(page);
    }
  return {
    units: 'pt',
    pageSizePt: size,
    scale,
    sourceFingerprint: snapshot.fingerprint,
    revision: snapshot.revision,
    sheet: snapshot.sheet,
    settings,
    ...(snapshot.geometryProfile ? { geometryProfile: snapshot.geometryProfile } : {}),
    pages,
    boundaries,
  };
}
const escape = (value) =>
  String(value).replace(
    /[&<>"']/g,
    (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]
  );
const cssNumber = (value) => {
  if (!Number.isFinite(value)) refuse('PRINT_INVALID_GEOMETRY');
  return String(value) + 'pt';
};
const color = (value) =>
  typeof value === 'string' && /^#?[0-9a-fA-F]{6}$/.test(value) ? '#' + value.replace(/^#/, '') : null;
function chartSvg(object) {
  const geometry = chartPreviewGeometry(object),
    appearance = (index) => chartSeriesAppearance(object.chart.style, index),
    stroke = (index) => {
      const a = appearance(index);
      return a.stroke
        ? ` stroke="${a.stroke}" stroke-width="${a.width}"${a.dash ? ` stroke-dasharray="${a.dash}"` : ''}`
        : '';
    };
  let body =
    '<title>' +
    escape(object.chart.title || object.name) +
    '</title><text x="300" y="20" text-anchor="middle" font-size="15">' +
    escape(object.chart.title || object.name) +
    '</text>';
  for (const [x1, y1, x2, y2] of geometry.lines)
    body += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#555"/>`;
  for (const bar of geometry.bars)
    body += `<rect x="${bar.x}" y="${bar.y}" width="${bar.width}" height="${bar.height}" fill="${appearance(bar.series).fill}"${stroke(bar.series)}><title>${escape(bar.title)}</title></rect>`;
  for (const segment of geometry.segments) {
    const fill = appearance(segment.series).color;
    body +=
      segment.points.length === 1
        ? `<circle cx="${segment.points[0][0]}" cy="${segment.points[0][1]}" r="2.5" fill="${fill}"/>`
        : `<polyline points="${segment.points.map((point) => point.join(',')).join(' ')}" fill="none" stroke="${fill}" stroke-width="2"${appearance(segment.series).dash ? ` stroke-dasharray="${appearance(segment.series).dash}"` : ''}/>`;
  }
  for (const label of geometry.labels)
    body += `<text x="${label.x}" y="${label.y}" text-anchor="${label.anchor}" font-size="11">${escape(label.text)}</text>`;
  if (object.chart.legend !== 'none')
    object.data.series.forEach((item, index) => {
      body +=
        '<g transform="translate(' +
        (geometry.left + (index * geometry.plotWidth) / object.data.series.length) +
        ',338)"><rect width="9" height="9" y="-8" fill="' +
        appearance(index).fill +
        '"' +
        stroke(index) +
        '/>' +
        '<text x="13" font-size="10">' +
        escape(item.name.slice(0, 30)) +
        '</text></g>';
    });
  return (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 350" width="100%" height="100%" role="img" aria-label="' +
    escape(object.name) +
    '">' +
    body +
    '</svg>'
  );
}
export function printPageHtml(layout, pageNumber) {
  const page = layout.pages.find((item) => item.number === pageNumber);
  if (!page) throw refusal('internal-unknown-reference', 'Unknown print page');
  const scale = layout.scale;
  let html =
    '<section class="print-page" aria-label="' +
    escape(layout.sheet + ' page ' + page.number) +
    '" data-print-page="' +
    page.number +
    '" data-xlsx-print-page="' +
    page.number +
    '" style="width:' +
    cssNumber(layout.pageSizePt[0]) +
    ';height:' +
    cssNumber(layout.pageSizePt[1]) +
    '">';
  const hf = layout.settings.headerFooter;
  if (hf) {
    const m = layout.settings.marginsPt,
      fontSize = 10 * (hf.scaleWithDoc ? scale : 1),
      height = fontSize * 1.2,
      width = (layout.pageSizePt[0] - m.left - m.right) / 3;
    for (const side of ['header', 'footer'])
      for (const [index, align] of ['left', 'center', 'right'].entries()) {
        const text = hf[side][align];
        if (!text) continue;
        const top = side === 'header' ? m.header : layout.pageSizePt[1] - m.footer - height;
        if (
          top < 0 ||
          (side === 'header' && top + height > m.top) ||
          (side === 'footer' && top < layout.pageSizePt[1] - m.bottom)
        )
          refuse('PRINT_HEADER_FOOTER_MARGIN');
        html +=
          '<div class="print-furniture" data-print-furniture="' +
          side +
          '-' +
          align +
          '" style="position:absolute;left:' +
          cssNumber(m.left + index * width) +
          ';top:' +
          cssNumber(top) +
          ';width:' +
          cssNumber(width) +
          ';height:' +
          cssNumber(height) +
          ';font-family:Arial,sans-serif;font-size:' +
          cssNumber(fontSize) +
          ';line-height:1.2;text-align:' +
          align +
          ';white-space:pre;overflow:hidden">' +
          escape(text) +
          '</div>';
      }
  }
  for (const cell of page.cells) {
    const s = cell.style,
      font = s.fontFamily ?? 'Arial';
    if (!/^[\p{L}\p{N} ._-]{1,80}$/u.test(font)) refuse('PRINT_UNSUPPORTED_FONT', { cell: cell.ref });
    const borderWidth = Math.min(
        s.border === 'all' ? 0.5 * scale : layout.settings.gridlines ? 0.25 * scale : 0,
        cell.bbox[2] / 2,
        cell.bbox[3] / 2
      ),
      padding = Math.min(1.5 * scale, Math.max(0, cell.bbox[2] / 2 - borderWidth));
    const align = s.verticalAlign === 'top' ? 'flex-start' : s.verticalAlign === 'center' ? 'center' : 'flex-end',
      style =
        'left:' +
        cssNumber(cell.bbox[0]) +
        ';top:' +
        cssNumber(cell.bbox[1]) +
        ';width:' +
        cssNumber(cell.bbox[2]) +
        ';height:' +
        cssNumber(cell.bbox[3]) +
        ';padding:0 ' +
        cssNumber(padding) +
        ';font-family:' +
        JSON.stringify(font) +
        ', sans-serif;font-size:' +
        cssNumber((s.fontSize ?? 11) * scale) +
        ';font-weight:' +
        (s.bold ? 'bold' : 'normal') +
        ';font-style:' +
        (s.italic ? 'italic' : 'normal') +
        ';text-decoration:' +
        (s.underline ? 'underline' : 'none') +
        ';color:' +
        (color(s.fontColor) ?? '#000000') +
        ';background:' +
        (color(s.fillColor) ?? 'transparent') +
        ';justify-content:' +
        align +
        ';text-align:' +
        (s.horizontalAlign ?? 'left') +
        ';border:0;box-shadow:' +
        (borderWidth
          ? 'inset 0 0 0 ' + cssNumber(borderWidth) + ' ' + (s.border === 'all' ? '#000' : '#b8b8b8')
          : 'none');
    html +=
      '<div class="print-cell" data-print-ref="' +
      escape(cell.ref) +
      '" style="' +
      escape(style) +
      '"><span class="print-text" style="white-space:' +
      (s.wrapText ? 'pre-wrap' : 'pre') +
      ';overflow-wrap:' +
      (s.wrapText ? 'anywhere' : 'normal') +
      '">' +
      escape(cell.text) +
      '</span></div>';
  }
  for (const item of page.objects) {
    const box = item.bbox,
      clip = item.clipBox,
      content =
        item.kind === 'chart'
          ? chartSvg(item.object)
          : '<img alt="' +
            escape(item.object.altText || item.object.name) +
            '" src="data:image/png;base64,' +
            escape(item.object.image.base64) +
            '" style="width:100%;height:100%;object-fit:fill"/>';
    html +=
      '<div class="print-object-clip" style="left:' +
      cssNumber(clip[0]) +
      ';top:' +
      cssNumber(clip[1]) +
      ';width:' +
      cssNumber(clip[2]) +
      ';height:' +
      cssNumber(clip[3]) +
      '"><div data-print-object="' +
      escape(item.id) +
      '" style="position:absolute;left:' +
      cssNumber(box[0] - clip[0]) +
      ';top:' +
      cssNumber(box[1] - clip[1]) +
      ';width:' +
      cssNumber(box[2]) +
      ';height:' +
      cssNumber(box[3]) +
      '">' +
      content +
      '</div></div>';
  }
  return html + '</section>';
}
export function printDocumentHtml(layout, { pages = layout.pages.map((page) => page.number) } = {}) {
  const html =
    '<!doctype html><html><head><meta charset="utf-8"><meta name="xlsx-print-source" content="' +
    escape(layout.sourceFingerprint) +
    '"><meta name="xlsx-print-revision" content="' +
    escape(layout.revision) +
    '"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src data:; style-src \'unsafe-inline\'"><title>' +
    escape(layout.sheet) +
    '</title><style>@page{size:' +
    cssNumber(layout.pageSizePt[0]) +
    ' ' +
    cssNumber(layout.pageSizePt[1]) +
    ';margin:0}*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;color:#000}.print-page{position:relative;overflow:hidden;break-after:page;page-break-after:always}.print-page:last-child{break-after:auto;page-break-after:auto}.print-cell{position:absolute;overflow:hidden;display:flex;flex-direction:column;line-height:1.2}.print-text{display:block;min-height:0;flex-shrink:0}.print-object-clip{position:absolute;overflow:hidden}svg{display:block}body{-webkit-print-color-adjust:exact;print-color-adjust:exact}</style></head><body>' +
    pages.map((number) => printPageHtml(layout, number)).join('') +
    '</body></html>';
  if (bytes(html) > PRINT_LIMITS.htmlBytes) refuse('PRINT_HTML_BOUND');
  return html;
}
/** Fixed readonly DOM inspection for the host after fonts/images are ready.
 * The generated document contains no script; this must run in its isolated
 * renderer before PDF/printer publication, not against the live editor. */
export function printOverflowInspectionSource() {
  return `(async()=>{await document.fonts.ready;await Promise.all([...document.images].map(image=>image.decode()));const failures=[];for(const cell of document.querySelectorAll('[data-print-ref]')){const text=cell.querySelector('.print-text'),style=getComputedStyle(cell),width=cell.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight),height=cell.clientHeight-parseFloat(style.paddingTop)-parseFloat(style.paddingBottom),box=text.getBoundingClientRect();if(text.scrollWidth>width+1||box.height>height+1)failures.push({cell:cell.dataset.printRef,page:Number(cell.closest('[data-print-page]').dataset.printPage),code:'PRINT_TEXT_OVERFLOW'});}for(const text of document.querySelectorAll('[data-print-furniture]')){if(text.scrollWidth>text.clientWidth+1||text.scrollHeight>text.clientHeight+1)failures.push({section:text.dataset.printFurniture,page:Number(text.closest('[data-print-page]').dataset.printPage),code:'PRINT_HEADER_FOOTER_OVERFLOW'});}return{pages:document.querySelectorAll('[data-print-page]').length,failures};})()`;
}
