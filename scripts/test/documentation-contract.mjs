// Every link between the editor's documents must lead somewhere.
//
//   node scripts/test/documentation-contract.mjs
//
// Checks relative links in docs/**/*.md and README.md: the file exists and,
// when the link names a heading, that heading exists. Web links are not fetched.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const component = fileURLToPath(new URL('../../', import.meta.url)),
  files = [path.join(component, 'README.md')];
async function walk(directory) {
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) await walk(full);
    else if (entry.name.endsWith('.md')) files.push(full);
  }
}
await walk(path.join(component, 'docs'));

const headings = new Map();
async function anchorsOf(file) {
  if (!headings.has(file)) {
    const text = await fs.readFile(file, 'utf8'),
      found = new Set([...text.matchAll(/<a id="([^"]+)"/g)].map((match) => match[1])),
      seen = new Map();
    for (const match of text.matchAll(/^#{1,6} (.+)$/gm)) {
      const slug = match[1]
          .toLowerCase()
          .replace(/[^\p{L}\p{N}_\- ]/gu, '')
          .replaceAll(' ', '-'),
        count = seen.get(slug) ?? 0;
      seen.set(slug, count + 1);
      found.add(slug + (count ? '-' + count : ''));
    }
    headings.set(file, found);
  }
  return headings.get(file);
}

const problems = [];
let links = 0;
for (const file of files) {
  const text = (await fs.readFile(file, 'utf8')).replace(/```[\s\S]*?```/g, '');
  for (const match of text.matchAll(/\]\(([^)\s]+)\)/g)) {
    const target = match[1];
    if (/^[a-z]+:/i.test(target)) continue;
    links += 1;
    const [relative, anchor] = target.split('#'),
      resolved = relative ? path.resolve(path.dirname(file), decodeURIComponent(relative)) : file,
      shown = path.relative(component, file);
    const info = await fs.stat(resolved).catch(() => null);
    if (!info) problems.push(`${shown}: ${target} does not exist`);
    else if (anchor && info.isFile() && resolved.endsWith('.md') && !(await anchorsOf(resolved)).has(anchor))
      problems.push(`${shown}: ${target} names a heading that does not exist`);
  }
}
for (const problem of problems) console.error(problem);
console.log(`${files.length} documents, ${links} links, ${problems.length} broken`);
process.exitCode = problems.length ? 1 : 0;
