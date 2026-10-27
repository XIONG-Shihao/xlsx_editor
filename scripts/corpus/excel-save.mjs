// Pass the made workbooks through Microsoft Excel so they are Excel's own files:
// calculated, with shared strings, a calculation chain and Excel's defaults.
//
//   node scripts/corpus/excel-save.mjs
//
// Reads corpus/made, writes corpus/files. Without Excel (or when Excel will not
// take a file) the workbook is copied as it is and that is recorded.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { excelDone, excelFolder, excelOpenSave, excelVersion } from '../audit/excel.mjs';

const component = new URL('../../', import.meta.url).pathname,
  made = path.join(component, 'corpus/made'),
  files = path.join(component, 'corpus/files'),
  manifestFile = path.join(component, 'corpus/manifest.json'),
  manifest = JSON.parse(await fs.readFile(manifestFile, 'utf8')),
  version = excelVersion(),
  folder = version ? await excelFolder('corpus') : null;
await fs.mkdir(files, { recursive: true });
manifest.workbooks = manifest.workbooks.filter((entry) => entry.url);
let viaExcel = 0;
for (const name of (await fs.readdir(made)).filter((item) => item.endsWith('.xlsx')).sort()) {
  const source = path.join(made, name),
    target = path.join(files, name);
  let savedBy = 'openpyxl (not opened in Excel)',
    problem = null;
  if (folder) {
    const working = path.join(folder, name);
    await fs.copyFile(source, working);
    problem = excelOpenSave(folder, working);
    if (!problem) {
      await fs.copyFile(working, target);
      savedBy = 'Microsoft Excel ' + version;
      viaExcel += 1;
    }
  }
  if (savedBy.startsWith('openpyxl')) await fs.copyFile(source, target);
  const bytes = await fs.readFile(target);
  manifest.workbooks.push({
    name: name.replace(/\.xlsx$/, ''),
    origin: name.startsWith('merged-') ? 'public table with features added here' : 'written here',
    source: 'scripts/corpus/build.py',
    licence: name.startsWith('merged-') ? 'as the public workbook it was taken from' : 'this project',
    savedBy,
    ...(problem ? { excelRefused: problem } : {}),
    bytes: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  });
  console.log((problem ? 'kept as made  ' : 'via Excel     ') + name + (problem ? '  (' + problem + ')' : ''));
}
if (folder) await fs.rm(folder, { recursive: true, force: true });
excelDone();
await fs.writeFile(manifestFile, JSON.stringify(manifest, null, 1) + '\n');
console.log(
  `${viaExcel} saved by Excel${version ? ' ' + version : ' (Excel not available)'}; ${manifest.workbooks.length} workbooks in the manifest`
);
