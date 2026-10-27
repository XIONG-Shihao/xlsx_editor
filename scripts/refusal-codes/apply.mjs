// Gives every refusal site its stable code, as decided by mapping.mjs. Safe to run again and again: a site that
// already carries a code is not touched, and the sentence of a refusal never changes, only the code is added.
//
//   node scripts/refusal-codes/apply.mjs                        all files
//   node scripts/refusal-codes/apply.mjs --skip=a.mjs,b.rs      leave these files (paths relative to the product folder) alone
//   node scripts/refusal-codes/apply.mjs --dry-run              report only
//
// After it, format the changed files: `npm run format` from the repository root (it only moves whitespace).
import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';
import { collect, ROOT } from './sites.mjs';
import { REFUSAL_CODES } from '../../core/refusal-codes.mjs';

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const skip = new Set(
  args
    .filter((a) => a.startsWith('--skip='))
    .flatMap((a) => a.slice(7).split(','))
    .filter(Boolean)
);

const sites = collect().filter((s) => s.status === 'code' && !skip.has(s.file));
const unknown = [...new Set(sites.map((s) => s.code).filter((c) => !Object.hasOwn(REFUSAL_CODES, c)))];
if (unknown.length) {
  console.error(
    'These codes are used by mapping.mjs but missing from core/refusal-codes.mjs:\n  ' + unknown.join('\n  ')
  );
  process.exit(1);
}

const byFile = Map.groupBy(sites, (s) => s.file);
let changedFiles = 0;
let changedSites = 0;
for (const [file, list] of byFile) {
  const absolute = path.join(ROOT, file);
  let text = fs.readFileSync(absolute, 'utf8');
  const edits = [];
  if (file.endsWith('.rs')) {
    for (const s of list) {
      if (s.dynamic) {
        // `Err(reason)`: wrap the whole argument
        const arg = text.slice(s.argStart, s.argEnd).replace(/\.(into|clone)\(\)$/, '');
        edits.push([s.argStart, s.argEnd, `crate::coded("${s.code}", ${arg})`]);
        continue;
      }
      const e = s.edit;
      const inner = e.wrapFormat ? text.slice(e.start, e.fmtEnd) : text.slice(e.strStart, e.strEnd);
      if (s.form === 'ok_or') {
        edits.push([s.nameStart, s.nameEnd, 'ok_or_else']);
        edits.push([e.start, e.end, `|| crate::coded("${s.code}", ${inner})`]);
      } else if (s.rewrap) {
        // keeps the code the inner error already has
        const wrapped = inner.replaceAll('{' + s.rewrap + '}', '{said}');
        edits.push([e.start, e.end, `crate::rephrase("${s.code}", ${s.rewrap}, |said| ${wrapped})`]);
      } else edits.push([e.start, e.end, `crate::coded("${s.code}", ${inner})`]);
    }
  } else {
    for (const s of list) edits.push([s.start, s.argStart, `refusal('${s.code}', `]);
  }
  edits.sort((a, b) => b[0] - a[0]);
  for (let i = 1; i < edits.length; i++) if (edits[i][1] > edits[i - 1][0]) throw Error(`Overlapping edits in ${file}`);
  for (const [start, end, replacement] of edits) text = text.slice(0, start) + replacement + text.slice(end);
  if (!file.endsWith('.rs')) text = withImport(file, text);
  changedSites += list.length;
  changedFiles++;
  if (!dryRun) fs.writeFileSync(absolute, text);
}
// Code that reads what an error says must not see its code in front. Each change is made once, where the old text stands.
const FOLLOW_UPS = [
  [
    'native/src/insertion.rs',
    'reason.starts_with("Unsupported structural formula: ")',
    'crate::sentence(&reason).starts_with("Unsupported structural formula: ")',
  ],
  [
    'native/src/package_edit/rewrite_admission.rs',
    'Err(error) if droppable && error == DELETED_GEOMETRY',
    'Err(error) if droppable && crate::sentence(&error) == DELETED_GEOMETRY',
  ],
];
for (const [file, from, to] of FOLLOW_UPS) {
  if (skip.has(file) || dryRun) continue;
  const absolute = path.join(ROOT, file);
  const text = fs.readFileSync(absolute, 'utf8');
  if (text.includes(from)) fs.writeFileSync(absolute, text.replace(from, to));
}
console.log(
  `${dryRun ? 'Would give' : 'Gave'} ${changedSites} sites a code in ${changedFiles} files${skip.size ? ` (skipped ${skip.size} files)` : ''}.`
);

/** Makes sure the file imports `refusal`. */
function withImport(file, text) {
  const kind = file.endsWith('.tsx') ? ts.ScriptKind.TSX : file.endsWith('.ts') ? ts.ScriptKind.TS : ts.ScriptKind.JS;
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, kind);
  // A file in a folder under core/ is one step further from the registry for each folder.
  const depth = file.split('/').length - 2;
  const specifier = file.startsWith('core/')
    ? (depth > 0 ? '../'.repeat(depth) : './') + 'refusal-codes.mjs'
    : '../core/refusal-codes.mjs';
  const imports = source.statements.filter(ts.isImportDeclaration);
  const existing = imports.find((i) => i.moduleSpecifier.text.endsWith('refusal-codes.mjs'));
  if (existing) {
    const bindings = existing.importClause?.namedBindings;
    if (bindings && ts.isNamedImports(bindings)) {
      if (bindings.elements.some((el) => el.name.text === 'refusal')) return text;
      const first = bindings.elements[0];
      return text.slice(0, first.getStart(source)) + 'refusal, ' + text.slice(first.getStart(source));
    }
    throw Error(`Unexpected import of refusal-codes in ${file}`);
  }
  const line = `import { refusal } from '${specifier}';\n`;
  if (imports.length) {
    const last = imports[imports.length - 1];
    return text.slice(0, last.end) + '\n' + line.trimEnd() + text.slice(last.end);
  }
  // No imports: after the opening comment block, if there is one
  const comments = ts.getLeadingCommentRanges(text, 0) ?? [];
  const at = comments.length ? comments[comments.length - 1].end : 0;
  return text.slice(0, at) + (at ? '\n' : '') + line + text.slice(at);
}
