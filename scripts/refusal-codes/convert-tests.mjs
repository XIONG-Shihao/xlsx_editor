// Turns test assertions that match a refusal's wording into assertions on its code.
//
//   assert.rejects(p, /Invalid range/)                  ->  assert.rejects(p, { code: 'internal-invalid-field' })
//   assert.throws(f, { message: /Unknown sheet/ })      ->  assert.throws(f, { code: '…' })
//   assert.match(error.message, /Unknown sheet/)        ->  assert.equal(error.code, '…')
//
// A pattern is converted only when every refusal site whose sentence it matches already has a code and the
// matches share at most three codes. Anything else stays on wording. Safe to run again; a converted test is
// not found again.
//
//   node scripts/refusal-codes/convert-tests.mjs --dry-run         what would change
//   node scripts/refusal-codes/convert-tests.mjs                   rewrite the tests
//   node scripts/refusal-codes/convert-tests.mjs --verify          rewrite, run each changed test file, and put
//                                                                  back the assertions that then fail
//   options: --only=name.mjs (one file)   --jobs=N (verify, default 3)
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import ts from 'typescript';
import { codedSites, collect, ROOT } from './sites.mjs';

const args = process.argv.slice(2);
const flag = (name) => args.includes('--' + name);
const option = (name) => args.find((a) => a.startsWith(`--${name}=`))?.slice(name.length + 3);
const only = option('only');
const jobs = Number(option('jobs') ?? 3);

// Every refusal sentence we know, with its code, or with null for a site that has no code yet.
const catalog = [
  ...codedSites()
    .filter((s) => !s.dynamic)
    .map((s) => ({ message: s.message, code: s.code })),
  ...collect()
    .filter((s) => s.status === 'code' && !s.dynamic)
    .map((s) => ({ message: s.message, code: null })),
];

/** The codes of every refusal sentence a pattern matches; null when none match or one of them has no code yet. */
function codesFor(pattern) {
  const found = new Set();
  for (const site of catalog) {
    // `{}` stands for a computed part of the sentence
    if (!site.message || !pattern.test(site.message.replace(/\{\}/g, 'X'))) continue;
    if (!site.code) return (skipped.uncoded++, null);
    found.add(site.code);
  }
  if (!found.size) skipped.noMatch++;
  else if (found.size > 3) skipped.tooMany++;
  return found.size ? found : null;
}
const skipped = { noMatch: 0, uncoded: 0, tooMany: 0 };

function replacement(codes) {
  const list = [...codes].sort();
  return list.length === 1 ? `'${list[0]}'` : `/^(${list.join('|')})$/`;
}

function parseRegex(text) {
  const m = /^\/((?:\\.|[^/\\])+)\/([a-z]*)$/s.exec(text);
  if (!m) return null;
  try {
    return new RegExp(m[1], m[2].replace(/[gy]/g, ''));
  } catch {
    return null;
  }
}

const testsDir = path.join(ROOT, 'tests');
const files = fs
  .readdirSync(testsDir)
  .filter((f) => f.endsWith('.mjs') && (!only || f === only))
  .sort();

/** { file, edits: [{ start, end, text, line, endLine }] } for one test file. */
function convertFile(name) {
  const text = fs.readFileSync(path.join(testsDir, name), 'utf8');
  const source = ts.createSourceFile(name, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  const edits = [];
  const lineOf = (pos) => source.getLineAndCharacterOfPosition(pos).line + 1;
  const visit = (node) => {
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
      const method = node.expression.name.text;
      const receiver = node.expression.expression.getText(source);
      const isAssert = /^(assert|assert\.strict|t\.assert|\w+\.assert)$/.test(receiver);
      if (isAssert && (method === 'rejects' || method === 'throws') && node.arguments.length >= 2) {
        const validator = node.arguments[1];
        if (ts.isRegularExpressionLiteral(validator)) {
          const pattern = parseRegex(validator.getText(source));
          const codes = pattern && codesFor(pattern);
          if (codes && codes.size <= 3)
            edits.push({
              start: validator.getStart(source),
              end: validator.end,
              text: `{ code: ${replacement(codes)} }`,
              line: lineOf(node.getStart(source)),
              endLine: lineOf(node.end),
            });
        } else if (ts.isObjectLiteralExpression(validator)) {
          for (const property of validator.properties) {
            if (
              !ts.isPropertyAssignment(property) ||
              property.name.getText(source) !== 'message' ||
              !ts.isRegularExpressionLiteral(property.initializer)
            )
              continue;
            const pattern = parseRegex(property.initializer.getText(source));
            const codes = pattern && codesFor(pattern);
            if (codes && codes.size <= 3)
              edits.push({
                start: property.getStart(source),
                end: property.end,
                text: `code: ${replacement(codes)}`,
                line: lineOf(node.getStart(source)),
                endLine: lineOf(node.end),
              });
          }
        }
      } else if (
        isAssert &&
        method === 'match' &&
        node.arguments.length >= 2 &&
        ts.isRegularExpressionLiteral(node.arguments[1])
      ) {
        const subject = node.arguments[0];
        if (
          ts.isPropertyAccessExpression(subject) &&
          subject.name.text === 'message' &&
          /^[\w$.?]+$/.test(subject.expression.getText(source))
        ) {
          const pattern = parseRegex(node.arguments[1].getText(source));
          const codes = pattern && codesFor(pattern);
          if (codes && codes.size <= 3) {
            const who = subject.expression.getText(source);
            const call =
              codes.size === 1
                ? `${receiver}.equal(${who}.code, ${replacement(codes)}`
                : `${receiver}.match(${who}.code, ${replacement(codes)}`;
            const rest =
              node.arguments.length > 2
                ? ', ' + source.text.slice(node.arguments[2].getStart(source), node.end - 1)
                : '';
            edits.push({
              start: node.getStart(source),
              end: node.end,
              text: `${call}${rest})`,
              line: lineOf(node.getStart(source)),
              endLine: lineOf(node.end),
            });
          }
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return { text, edits };
}

function apply(text, edits) {
  let out = text;
  for (const e of [...edits].sort((a, b) => b.start - a.start)) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return out;
}

function runTest(name) {
  return new Promise((resolve) => {
    const child = spawn(
      process.execPath,
      ['--import', './scripts/ts-resolve-register.mjs', '--test', '--test-concurrency=1', `tests/${name}`],
      { cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe'] }
    );
    let output = '';
    child.stdout.on('data', (c) => (output += c));
    child.stderr.on('data', (c) => (output += c));
    const timer = setTimeout(() => child.kill('SIGKILL'), 600000);
    child.on('close', (code) => {
      clearTimeout(timer);
      resolve({ ok: code === 0, output });
    });
  });
}

const summary = { files: 0, converted: 0, reverted: 0, filesReverted: [] };
const queue = files.map((f) => ({ name: f, ...convertFile(f) })).filter((f) => f.edits.length);
if (flag('dry-run')) {
  for (const f of queue) console.log(`${f.name}: ${f.edits.length}`);
  console.log(
    `${queue.reduce((n, f) => n + f.edits.length, 0)} assertions in ${queue.length} files could be converted; left on wording: ${JSON.stringify(skipped)}`
  );
  process.exit(0);
}
const work = async (file) => {
  const full = path.join(testsDir, file.name);
  let edits = file.edits;
  fs.writeFileSync(full, apply(file.text, edits));
  if (flag('verify')) {
    // Run the file. Where it fails, put back the assertions on the lines the failure points at and run again;
    // if nothing points at an assertion, or it still fails after a few rounds, put the whole file back.
    let passed = false;
    for (let round = 0; round < 6 && edits.length; round++) {
      const result = await runTest(file.name);
      if (result.ok) {
        passed = true;
        break;
      }
      const lines = [...result.output.matchAll(new RegExp(`${file.name.replace('.', '\\.')}:(\\d+):\\d+`, 'g'))].map(
        (m) => Number(m[1])
      );
      // converted calls keep their line count, so the failing line is the original's line
      const keep = edits.filter((e) => !lines.some((l) => l >= e.line && l <= e.endLine));
      edits = keep.length === edits.length ? [] : keep;
      fs.writeFileSync(full, apply(file.text, edits));
    }
    if (!passed && edits.length) {
      edits = [];
      fs.writeFileSync(full, file.text);
      summary.filesReverted.push(file.name);
    }
  }
  summary.files += edits.length ? 1 : 0;
  summary.converted += edits.length;
  summary.reverted += file.edits.length - edits.length;
  console.log(`${file.name}: ${edits.length}/${file.edits.length} converted`);
};
const pool = [...queue];
await Promise.all(
  Array.from({ length: flag('verify') ? jobs : 1 }, async () => {
    while (pool.length) await work(pool.shift());
  })
);
console.log(JSON.stringify(summary));
