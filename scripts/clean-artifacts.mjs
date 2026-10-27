/** Preview-first component cleanup. Qualified deliveries and reports are retained. */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const metadata = (file) => {
  const s = fs.lstatSync(file);
  return { dev: s.dev, ino: s.ino, size: s.size, mtimeMs: s.mtimeMs };
};
function safe(root, relative) {
  const file = path.resolve(root, relative);
  if (!file.startsWith(root + path.sep)) throw Error('Path escapes component');
  let current = root;
  for (const part of path.relative(root, file).split(path.sep)) {
    current = path.join(current, part);
    if (fs.existsSync(current) && fs.lstatSync(current).isSymbolicLink())
      throw Error('Symlinked cleanup path: ' + relative);
  }
  return file;
}
export function planCleanup(root) {
  root = fs.realpathSync(root);
  const artifactRoot = path.join(root, 'artifacts'),
    marker = path.join(artifactRoot, 'current.json');
  if (!fs.existsSync(marker)) throw Error('current.json is required before retiring any checkpoint');
  const current = JSON.parse(fs.readFileSync(marker)),
    app = path.resolve(artifactRoot, current.app);
  safe(root, path.relative(root, app));
  if (!fs.existsSync(app)) throw Error('The qualified app must exist before cleanup');
  const moves = [],
    removals = [],
    checks = [{ path: 'artifacts/current.json', identity: metadata(marker) }];
  const move = (from, to) => moves.push({ from, to, identity: metadata(safe(root, from)) });
  const remove = (file, reason) => removals.push({ path: file, reason, identity: metadata(safe(root, file)) });
  const entries = fs.readdirSync(artifactRoot, { withFileTypes: true });
  for (const entry of entries) {
    const from = 'artifacts/' + entry.name;
    if (entry.isSymbolicLink()) throw Error('Unclassified artifact symlink: ' + from);
    if (/^checkpoint-\d+$/.test(entry.name) && entry.isDirectory()) {
      const directory = path.join(root, from);
      if (app.startsWith(directory + path.sep)) move(from, 'artifacts/releases/' + entry.name);
      else remove(from, 'Superseded standalone checkpoint; qualified app is protected');
    } else if (entry.name === 'final') move(from, 'artifacts/packages');
    else if (entry.name === 'npm-cache') move(from, '.build/cache/npm');
    else if (entry.name === 'engine-package') move(from, '.build/staging/engine-package');
    else if (
      ['bench', 'browser', 'roundtrip', 'preservation-matrix', 'corpus'].includes(entry.name) ||
      /^(?:large-desktop|desktop-test)-\d+$/.test(entry.name)
    )
      move(from, '.build/runs/' + entry.name);
    else if (!['current.json', 'releases', 'packages'].includes(entry.name))
      throw Error('Unclassified artifact: ' + from);
  }
  const large = entries
    .filter((e) => /^large-desktop-\d+$/.test(e.name) && e.isDirectory())
    .map((e) => {
      const reportPath = 'artifacts/' + e.name + '/report.json',
        file = safe(root, reportPath);
      if (!fs.existsSync(file)) throw Error('Retain incomplete run without report: ' + e.name);
      checks.push({ path: reportPath, identity: metadata(file) });
      return { name: e.name, report: JSON.parse(fs.readFileSync(file)) };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
  const latest = large.filter((r) => r.report.passed === true).at(-1)?.name;
  for (const { name, report } of large) {
    const directory = 'artifacts/' + name;
    if (name !== latest)
      for (const item of report.cases ?? []) {
        const file = path.resolve(item.file);
        // Only direct disposable copies recorded by this run; never source paths.
        if (path.dirname(file) !== path.join(root, directory) || file === path.resolve(item.source))
          throw Error('Unproved disposable fixture: ' + file);
        if (!['.xlsx', '.csv', '.tsv'].includes(path.extname(file))) throw Error('Unexpected fixture type');
        if (fs.existsSync(file))
          remove(
            path.relative(root, file),
            'Older disposable saved workbook; report retained and latest successful outputs kept'
          );
      }
  }
  for (const e of entries.filter((e) => /^(?:large-desktop|desktop-test)-\d+$/.test(e.name))) {
    const report = 'artifacts/' + e.name + '/report.json',
      profile = 'artifacts/' + e.name + '/profile';
    if (fs.existsSync(safe(root, profile)) && fs.existsSync(safe(root, report)))
      remove(profile, 'Closed standalone test browser profile; documents and reports are separate');
  }
  for (const m of moves) if (fs.existsSync(safe(root, m.to))) throw Error('Destination already exists: ' + m.to);
  return {
    schema: 1,
    root,
    createdAt: new Date().toISOString(),
    moves,
    removals,
    checks,
    current,
    latestSuccessfulLargeRun: latest,
    policy:
      'Keep qualified deliveries, all reports/screenshots, benchmark baselines and latest successful large-document outputs. Never follow symlinks or touch original supplied files.',
  };
}
export function applyCleanup(plan, processLines) {
  const root = fs.realpathSync(plan.root);
  if (!Array.isArray(processLines)) throw Error('Process inspection is required');
  if (processLines.some((line) => line.includes(root + '/artifacts/') || line.includes(root + '/.build/')))
    throw Error('An active process refers to component outputs; stop it first');
  if (!plan.moves.length && !plan.removals.length) return { ...plan, completed: { moves: [], removals: [] } };
  for (const item of [...plan.checks, ...plan.removals, ...plan.moves.map((m) => ({ ...m, path: m.from }))])
    if (JSON.stringify(metadata(safe(root, item.path))) !== JSON.stringify(item.identity))
      throw Error('Changed after preview: ' + item.path);
  for (const m of plan.moves) if (fs.existsSync(safe(root, m.to))) throw Error('Destination already exists: ' + m.to);
  fs.mkdirSync(path.join(root, '.build'), { recursive: true });
  const completed = { moves: [], removals: [] },
    ledger = path.join(root, '.build/artifact-cleanup.json');
  const record = () => fs.writeFileSync(ledger, JSON.stringify({ ...plan, completed }, null, 2) + '\n');
  record();
  for (const item of plan.removals) {
    fs.rmSync(safe(root, item.path), { recursive: true });
    completed.removals.push(item);
    record();
  }
  let app = path.resolve(root, 'artifacts', plan.current.app);
  for (const move of plan.moves) {
    const from = safe(root, move.from),
      to = safe(root, move.to);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.renameSync(from, to);
    if (app.startsWith(from + path.sep)) app = to + app.slice(from.length);
    completed.moves.push(move);
    record();
  }
  fs.writeFileSync(path.join(root, 'artifacts/current.json'), JSON.stringify({ ...plan.current, app }, null, 2) + '\n');
  for (const name of ['.build/cache', '.build/staging', '.build/runs', 'artifacts/releases', 'artifacts/packages'])
    fs.mkdirSync(path.join(root, name), { recursive: true });
  return { ...plan, completed, currentApp: app };
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  if (process.argv.slice(2).some((a) => a !== '--apply'))
    throw Error('Usage: node scripts/clean-artifacts.mjs [--apply]');
  const plan = planCleanup(fileURLToPath(new URL('../', import.meta.url)));
  if (process.argv.includes('--apply')) {
    const lines = execFileSync('ps', ['-axo', 'args='], { encoding: 'utf8' }).split('\n');
    const result = applyCleanup(plan, lines);
    console.log(
      JSON.stringify({
        moved: result.completed.moves.length,
        removed: result.completed.removals.length,
        currentApp: result.currentApp,
      })
    );
  } else console.log(JSON.stringify(plan, null, 2));
}
