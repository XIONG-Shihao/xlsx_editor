// Keeps source lines readable. Code is formatted mechanically (Prettier for
// JavaScript/TypeScript/CSS, rustfmt for Rust); what stays long is text the
// formatters will not split: SQL, XML and messages. This check lets those
// stay but stops the count from growing. Lower the numbers when you shorten
// lines; never raise them to make a change pass.
//
//   node scripts/test/line-length.mjs            check
//   node scripts/test/line-length.mjs --update   record the current counts
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';

const LIMIT = 160;
const areas = {
  core: /^core\/.*\.(mjs|ts|mts)$/,
  ui: /^ui\/.*\.(ts|tsx)$/,
  adapters: /^(adapters|standalone)\/.*\.(mjs|cjs|ts|tsx)$/,
  native: /^native\/src\/.*\.rs$/,
  tests: /^(tests|scripts)\/.*\.(mjs|cjs|ts|tsx)$/,
};
const baselineFile = new URL('./line-length-baseline.json', import.meta.url);
const files = execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard'], { cwd: new URL('../..', import.meta.url), encoding: 'utf8' })
  .split('\n')
  .filter((file) => file && existsSync(new URL('../../' + file, import.meta.url)));
const current = {};
for (const [area, pattern] of Object.entries(areas)) {
  let long = 0;
  let longest = 0;
  for (const file of files.filter((name) => pattern.test(name))) {
    for (const line of readFileSync(new URL('../../' + file, import.meta.url), 'utf8').split('\n')) {
      if (line.length > LIMIT) long += 1;
      if (line.length > longest) longest = line.length;
    }
  }
  current[area] = { long, longest };
}
if (process.argv.includes('--update')) {
  writeFileSync(baselineFile, JSON.stringify({ limit: LIMIT, areas: current }, null, 2) + '\n');
  console.log('Recorded', JSON.stringify(current));
} else {
  const baseline = JSON.parse(readFileSync(baselineFile, 'utf8')).areas;
  const problems = Object.entries(current)
    .filter(([area, now]) => now.long > baseline[area].long || now.longest > baseline[area].longest)
    .map(
      ([area, now]) =>
        `${area}: ${now.long} lines over ${LIMIT} characters (allowed ${baseline[area].long}), longest ${now.longest} (allowed ${baseline[area].longest})`
    );
  if (problems.length) {
    console.error('Lines grew longer. Format the code, or split the long text:\n' + problems.join('\n'));
    process.exit(1);
  }
  console.log('PASS line length:', JSON.stringify(current));
}
