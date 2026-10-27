// Collects every refusal site in core/, adapters/ and native/src/ and decides what each one needs.
import fs from 'node:fs';
import path from 'node:path';
import { scanJs, scanJsCoded } from './scan-js.mjs';
import { scanRs, scanRsCoded } from './scan-rs.mjs';
import { codeFor, domainOf, isSignal } from './classify.mjs';
import { DYNAMIC } from './mapping.mjs';

/** The product folder. */
export const ROOT = new URL('../../', import.meta.url).pathname;

const walk = (dir) => {
  const absolute = path.join(ROOT, dir);
  if (!fs.existsSync(absolute)) return [];
  return fs.readdirSync(absolute, { withFileTypes: true }).flatMap((entry) => {
    const relative = path.join(dir, entry.name);
    return entry.isDirectory() ? walk(relative) : [relative];
  });
};

/** Source files that can hold refusal sites, relative to the product folder. */
export function sourceFiles() {
  const js = [...walk('core'), ...walk('adapters')].filter(
    (f) => /\.(mjs|ts|tsx)$/.test(f) && !/\.d\.(mts|ts)$/.test(f)
  );
  // main.rs is the helper's wire framing (frame sizes), a separate crate that cannot reach crate::coded
  const rs = walk('native/src').filter((f) => f.endsWith('.rs') && f !== 'native/src/main.rs');
  return { js: js.filter((f) => f !== REGISTRY_FILE).sort(), rs: rs.sort() };
}

export const REGISTRY_FILE = 'core/refusal-codes.mjs';

/** What to do with a dynamic site (`Error(reason)`): a code, or leave it alone. */
function dynamicDecision(file, argText) {
  for (const rule of DYNAMIC) {
    if (rule.file && !rule.file.test(file)) continue;
    if (!rule.arg.test(argText)) continue;
    if (rule.how === 'code') return { status: 'code', code: rule.code };
    return { status: rule.how };
  }
  // A stored reason string: the file's domain says what it is about.
  return { status: 'code', code: `${domainOf(file, '')}-refused` };
}

/**
 * Every site, each with a `status`:
 *   code       it should carry `code`
 *   forwarded  it re-throws an error made elsewhere
 *   legacy     it already carries an older code of its own
 *   signal     not a refusal: a value the formula engine passes on as an error (see SIGNALS)
 */
export function collect() {
  const files = sourceFiles();
  const sites = [];
  for (const file of files.js) {
    const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
    for (const site of scanJs(file, text)) {
      let decision;
      if (site.legacyCode) decision = { status: 'legacy' };
      else if (site.dynamic) decision = dynamicDecision(file, site.argText);
      else decision = { status: 'code', code: codeFor(file, site.message) };
      sites.push({ lang: 'js', ...site, ...decision });
    }
  }
  for (const file of files.rs) {
    const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
    for (const site of scanRs(file, text)) {
      let decision;
      if (site.dynamic) decision = dynamicDecision(file, site.argText);
      else if (isSignal(file, site.message)) decision = { status: 'signal' };
      else decision = { status: 'code', code: codeFor(file, site.message) };
      sites.push({ lang: 'rs', ...site, ...decision });
    }
  }
  return sites;
}

/** How many sites already carry a code (a `refusal('code', …)` or `crate::coded("code", …)` call). */
export function countCoded() {
  const files = sourceFiles();
  const counts = { js: 0, rs: 0, byFile: {} };
  for (const file of files.js) {
    if (file === REGISTRY_FILE) continue;
    const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
    const n = (text.match(/\brefusal\(\s*['"][a-z-]+['"]/g) ?? []).length;
    counts.js += n;
    if (n) counts.byFile[file] = n;
  }
  for (const file of files.rs) {
    const text = fs.readFileSync(path.join(ROOT, file), 'utf8');
    const n = (text.match(/\bcoded\(\s*"[a-z-]+"/g) ?? []).length;
    counts.rs += n;
    if (n) counts.byFile[file] = n;
  }
  return counts;
}

/** Every refusal that already has a code, with its sentence: { file, code, message, dynamic, line }. */
export function codedSites() {
  const files = sourceFiles();
  const found = [];
  for (const file of files.js) found.push(...scanJsCoded(file, fs.readFileSync(path.join(ROOT, file), 'utf8')));
  for (const file of files.rs) found.push(...scanRsCoded(file, fs.readFileSync(path.join(ROOT, file), 'utf8')));
  return found;
}
