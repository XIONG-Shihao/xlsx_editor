// Gather the standing set of real workbooks.
//
//   node scripts/corpus/fetch.mjs            download every workbook in corpus/manifest.json
//   node scripts/corpus/fetch.mjs --discover find new public workbooks and add them to the manifest
//   node scripts/corpus/fetch.mjs --refresh  also take files whose publisher has changed them since
//                                            they were recorded (the manifest keeps the old checksum
//                                            under `replaces`); their scores are not comparable
//
// Workbooks are public files from open-data catalogues and government sites.
// They are kept in corpus/files (not committed); the manifest records where
// each came from, its licence, and its checksum, so the set can be rebuilt.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';

const component = new URL('../../', import.meta.url).pathname,
  folder = path.join(component, 'corpus/files'),
  manifestFile = path.join(component, 'corpus/manifest.json'),
  discover = process.argv.includes('--discover'),
  refresh = process.argv.includes('--refresh');
await fs.mkdir(folder, { recursive: true });
const manifest = JSON.parse(await fs.readFile(manifestFile, 'utf8').catch(() => '{"workbooks":[]}'));

// CKAN catalogues that answer the standard search. `take` is how many to aim for.
const catalogues = [
  { id: 'canada', api: 'https://open.canada.ca/data/api/3/action/package_search', take: 18 },
  { id: 'australia', api: 'https://data.gov.au/data/api/3/action/package_search', take: 16 },
  { id: 'ireland', api: 'https://data.gov.ie/api/3/action/package_search', take: 12 },
  { id: 'humanitarian', api: 'https://data.humdata.org/api/3/action/package_search', take: 18 },
];
// Files published at stable addresses by statistics offices.
const direct = [
  [
    'us-census-state-population',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/state/totals/NST-EST2023-POP.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-county-population',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/counties/totals/co-est2023-pop.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-metro-population',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/metro/totals/cbsa-met-est2023-pop.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-city-population',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/cities/totals/SUB-IP-EST2023-POP.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-age-sex',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/national/asrh/nc-est2023-agesex.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-bls-cpi-detail',
    'https://www.bls.gov/cpi/tables/supplemental-files/news-release-table1-202312.xlsx',
    'US Bureau of Labor Statistics',
    'US government work (public domain)',
  ],
  [
    'us-bea-gdp-release',
    'https://www.bea.gov/sites/default/files/2024-01/gdp4q23-adv.xlsx',
    'US Bureau of Economic Analysis',
    'US government work (public domain)',
  ],
  [
    'us-eia-electricity',
    'https://www.eia.gov/electricity/monthly/xls/table_1_01.xlsx',
    'US Energy Information Administration',
    'US government work (public domain)',
  ],
  [
    'us-census-housing-units',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/housing/totals/NST-EST2023-HU.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-components-of-change',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/state/totals/NST-EST2023-COMP.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
  [
    'us-census-puerto-rico',
    'https://www2.census.gov/programs-surveys/popest/tables/2020-2023/municipios/totals/prm-est2023-pop.xlsx',
    'US Census Bureau',
    'US government work (public domain)',
  ],
];
const OPEN_LICENCE =
  /creative commons|\bcc[- ]?(by|0|zero)|open government|\bogl\b|public domain|open data|odc|odbl|pddl/i;
const MIN = 6 * 1024,
  MAX = 6 * 1024 * 1024;

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const slug = (text) =>
  text
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 48) || 'workbook';

/** A workbook the editor is meant to open: a ZIP with the workbook part, not encrypted. */
function inspect(bytes) {
  if (bytes.length < MIN || bytes.length > MAX) return null;
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return null;
  let parts;
  try {
    parts = unzipSync(bytes, {
      filter: (file) => ['docProps/app.xml', 'xl/workbook.xml'].includes(file.name),
    });
  } catch {
    return null;
  }
  if (!parts['xl/workbook.xml']) return null;
  const app = parts['docProps/app.xml'] ? strFromU8(parts['docProps/app.xml']) : '';
  return {
    producer: /<Application>([^<]*)</.exec(app)?.[1] ?? 'unknown',
    version: /<AppVersion>([^<]*)</.exec(app)?.[1] ?? '',
    sheets: (strFromU8(parts['xl/workbook.xml']).match(/<(?:\w+:)?sheet\b/g) ?? []).length,
  };
}
async function download(url) {
  const reply = await fetch(url, {
    redirect: 'follow',
    signal: AbortSignal.timeout(60000),
    headers: { 'user-agent': 'workbook-corpus/1 (compatibility testing of an open workbook editor)' },
  });
  if (!reply.ok) throw Error('HTTP ' + reply.status);
  const length = Number(reply.headers.get('content-length') ?? 0);
  if (length > MAX) throw Error('too large');
  const bytes = new Uint8Array(await reply.arrayBuffer());
  if (bytes.length > MAX) throw Error('too large');
  return bytes;
}
async function keep(entry, bytes) {
  const found = inspect(bytes);
  if (!found) return false;
  const digest = sha256(bytes);
  if (manifest.workbooks.some((item) => item.sha256 === digest)) return false;
  let name = entry.name;
  for (let n = 2; manifest.workbooks.some((item) => item.name === name); n++) name = `${entry.name}-${n}`;
  await fs.writeFile(path.join(folder, name + '.xlsx'), bytes);
  manifest.workbooks.push({ ...entry, name, sha256: digest, bytes: bytes.length, ...found });
  console.log('kept', name, found.producer, found.sheets + ' sheets', Math.round(bytes.length / 1024) + ' KB');
  return true;
}

if (!discover) {
  let present = 0,
    fetched = 0,
    failed = 0;
  const refreshed = [];
  for (const entry of manifest.workbooks) {
    const file = path.join(folder, entry.name + '.xlsx'),
      existing = await fs.readFile(file).catch(() => null);
    if (existing && sha256(existing) === entry.sha256) {
      present += 1;
      continue;
    }
    if (!entry.url) {
      console.log('missing', entry.name, '(made locally: run scripts/corpus/build.mjs)');
      failed += 1;
      continue;
    }
    try {
      const bytes = await download(entry.url);
      if (sha256(bytes) !== entry.sha256) {
        if (!refresh) throw Error('the published file has changed since it was recorded (--refresh takes it)');
        entry.replaces = [...(entry.replaces ?? []), entry.sha256];
        Object.assign(entry, {
          sha256: sha256(bytes),
          bytes: bytes.length,
          refreshed: new Date().toISOString().slice(0, 10),
        });
        refreshed.push(entry.name);
      }
      await fs.writeFile(file, bytes);
      fetched += 1;
    } catch (error) {
      console.log('could not fetch', entry.name, String(error.message ?? error));
      failed += 1;
    }
  }
  if (refreshed.length) {
    await fs.writeFile(manifestFile, JSON.stringify(manifest, null, 1) + '\n');
    console.log(`Changed by their publishers, taken as they are now: ${refreshed.join(', ')}`);
  }
  console.log(`${present} already here, ${fetched} fetched, ${failed} missing, of ${manifest.workbooks.length}`);
  process.exit(failed ? 1 : 0);
}

for (const [name, url, source, licence] of direct) {
  if (manifest.workbooks.some((item) => item.url === url)) continue;
  try {
    await keep({ name, url, source, licence, origin: 'statistics office' }, await download(url));
  } catch (error) {
    console.log('skip', name, String(error.message ?? error));
  }
}
for (const catalogue of catalogues) {
  let kept = manifest.workbooks.filter((item) => item.catalogue === catalogue.id).length;
  const publishers = new Map();
  for (const item of manifest.workbooks.filter((entry) => entry.catalogue === catalogue.id))
    publishers.set(item.source, (publishers.get(item.source) ?? 0) + 1);
  // Walk the catalogue in strides so the set is spread over publishers and years.
  for (let start = 0; kept < catalogue.take && start < 4000; start += 137) {
    let reply;
    try {
      reply = await (
        await fetch(`${catalogue.api}?fq=res_format:XLSX&rows=25&start=${start}`, {
          signal: AbortSignal.timeout(45000),
        })
      ).json();
    } catch (error) {
      console.log('catalogue', catalogue.id, 'did not answer at', start, String(error.message ?? error));
      continue;
    }
    for (const dataset of reply?.result?.results ?? []) {
      if (kept >= catalogue.take) break;
      const publisher =
        (typeof dataset.organization?.title === 'string' ? dataset.organization.title : null) ??
        dataset.organization?.name ??
        catalogue.id;
      if ((publishers.get(publisher) ?? 0) >= 2) continue;
      // Some catalogues mark every licence "not open", so the licence is judged by its name.
      const licence = dataset.license_title ?? dataset.license_id;
      if (!licence || !OPEN_LICENCE.test(licence)) continue;
      const resource = (dataset.resources ?? []).find(
        (item) => String(item.format ?? '').toUpperCase() === 'XLSX' && /^https?:\/\//.test(item.url ?? '')
      );
      if (!resource || manifest.workbooks.some((item) => item.url === resource.url)) continue;
      const title = typeof dataset.title === 'string' ? dataset.title : (dataset.title?.en ?? dataset.name);
      try {
        const saved = await keep(
          {
            name: `${catalogue.id}-${slug(String(title))}`,
            url: resource.url,
            source: publisher,
            licence,
            origin: 'open-data catalogue',
            catalogue: catalogue.id,
            dataset: dataset.name,
          },
          await download(resource.url)
        );
        if (saved) {
          kept += 1;
          publishers.set(publisher, (publishers.get(publisher) ?? 0) + 1);
        }
      } catch (error) {
        console.log('skip', String(title).slice(0, 50), String(error.message ?? error));
      }
    }
    await fs.writeFile(manifestFile, JSON.stringify(manifest, null, 1) + '\n');
  }
}
await fs.writeFile(manifestFile, JSON.stringify(manifest, null, 1) + '\n');
console.log(manifest.workbooks.length + ' workbooks in the manifest');
