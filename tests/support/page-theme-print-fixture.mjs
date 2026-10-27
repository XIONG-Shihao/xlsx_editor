import { createWorkbook } from '../../core/workbook-package.mjs';
import { pageThemeXml } from '../../core/page-style-samples.mjs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function pageThemePrintFixture({
  theme = null,
  foreign = false,
  tint = false,
  modified = false,
  formula = true,
} = {}) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetFormatPr defaultRowHeight="30" defaultColWidth="12"/><sheetData><row r="1" ht="30" customHeight="1"><c r="A1" t="inlineStr"><is><t>Body</t></is></c><c r="B1">${formula ? '<f>SUM(D1,1)</f><v>777</v>' : '<v>6</v>'}</c><c r="C1" s="1" t="inlineStr"><is><t>Literal</t></is></c><c r="D1"><v>5</v></c><c r="E1" s="2" t="inlineStr"><is><t>Title</t></is></c><c r="F1" s="3" t="inlineStr"><is><t>Fill</t></is></c></row></sheetData></worksheet>`
  );
  parts['xl/styles.xml'] = strToU8(
    `<styleSheet xmlns="${ns}"><fonts count="3"><font><name val="Calibri"/><sz val="11"/>${theme ? '<scheme val="minor"/>' : ''}<color theme="4"${tint ? ' tint="0.2"' : ''}/></font><font><name val="Courier New"/><sz val="12"/><color rgb="FF11AA44"/></font><font><name val="Arial"/><sz val="14"/><b/><scheme val="major"/><color theme="5"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor theme="6"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/></cellStyleXfs><cellXfs count="4"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="0" borderId="0" numFmtId="0"/><xf fontId="2" fillId="0" borderId="0" numFmtId="0"/><xf fontId="0" fillId="1" borderId="0" numFmtId="0"/></cellXfs></styleSheet>`
  );
  if (theme) {
    const name = foreign ? 'imported.xml' : 'workbook-page-theme.xml';
    parts['xl/theme/' + name] = strToU8(pageThemeXml(theme) + (modified ? '\n' : ''));
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
        '</Relationships>',
        `<Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/${name}"/></Relationships>`
      )
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        `<Override PartName="/xl/theme/${name}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`
      )
    );
  }
  parts['custom/opaque.bin'] = new Uint8Array([1, 3, 5, 255]);
  return zipSync(parts);
}
