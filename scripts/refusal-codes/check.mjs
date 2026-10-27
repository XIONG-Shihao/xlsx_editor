// Reports refusal sites that still have no code, and keeps the allow-list honest.
//
//   node scripts/refusal-codes/check.mjs             check (exit 1 on a problem)
//   node scripts/refusal-codes/check.mjs --list      also list every uncoded site
//   node scripts/refusal-codes/check.mjs --shrink    lower the allow-list to what is still uncoded (never raises it)
//
// The allow-list (allow-list.json) says how many uncoded sites each file may still have. It may only
// shrink: a new refusal needs a code (add a rule to mapping.mjs, run apply.mjs), not an allowance.
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { collect, countCoded } from './sites.mjs';
import { REFUSAL_CODES } from '../../core/refusal-codes.mjs';

export const ALLOW_LIST = new URL('./allow-list.json', import.meta.url);

/** Sites that should carry a code and do not yet. */
export function uncoded() {
  return collect().filter((s) => s.status === 'code');
}

export function readAllowList() {
  return JSON.parse(fs.readFileSync(ALLOW_LIST, 'utf8'));
}

/** Codes that the sites still waiting for a code will receive; they count as in use. */
export function pendingCodes() {
  return new Set(uncoded().map((s) => s.code));
}

/** Problems as sentences; empty when everything is in step. */
export function problems() {
  const found = [];
  const waiting = uncoded();
  const perFile = Map.groupBy(waiting, (s) => s.file);
  const allowed = readAllowList().files;
  for (const [file, list] of perFile) {
    const limit = allowed[file] ?? 0;
    if (list.length > limit)
      found.push(
        `${file}: ${list.length} refusal sites have no code (allowed ${limit}). First: line ${list[0].line}, "${(list[0].message || list[0].argText).slice(0, 80)}". ` +
          'Run `node scripts/refusal-codes/apply.mjs`, and add a rule to mapping.mjs if it reports a site no rule covers.'
      );
  }
  for (const [file, limit] of Object.entries(allowed)) {
    const now = perFile.get(file)?.length ?? 0;
    if (now < limit)
      found.push(
        `${file}: the allow-list says ${limit} uncoded sites but only ${now} remain. Lower it: node scripts/refusal-codes/check.mjs --shrink`
      );
  }
  const wrong = [...new Set(waiting.map((s) => s.code))].filter((code) => !Object.hasOwn(REFUSAL_CODES, code));
  if (wrong.length)
    found.push(`mapping.mjs gives codes that core/refusal-codes.mjs does not list: ${wrong.join(', ')}`);
  return found;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const waiting = uncoded();
  const perFile = Map.groupBy(waiting, (s) => s.file);
  if (process.argv.includes('--list'))
    for (const s of waiting) console.log(`${s.file}:${s.line}  [${s.code}]  ${s.message || s.argText}`);
  if (process.argv.includes('--shrink')) {
    const allowed = readAllowList();
    const next = {};
    for (const [file, list] of perFile) {
      if (list.length > (allowed.files[file] ?? 0)) {
        console.error(`${file} has more uncoded sites than allowed; the allow-list never grows.`);
        process.exit(1);
      }
      next[file] = list.length;
    }
    fs.writeFileSync(
      ALLOW_LIST,
      JSON.stringify({ ...allowed, files: Object.fromEntries(Object.entries(next).sort()) }, null, 2) + '\n'
    );
  }
  const coded = countCoded();
  const forwarded = collect().filter((s) => s.status !== 'code');
  console.log(
    `Sites with a code: ${coded.js} in JavaScript, ${coded.rs} in Rust. Without: ${waiting.length}. ` +
      `Not refusals of their own (forwarded, older code, formula signal): ${forwarded.length}. Codes listed: ${Object.keys(REFUSAL_CODES).length}.`
  );
  const issues = problems();
  for (const issue of issues) console.error(issue);
  process.exit(issues.length ? 1 : 0);
}
