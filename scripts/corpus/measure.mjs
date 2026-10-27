// The scoreboard: how the editor does on the standing set of real workbooks.
//
//   node --import ./scripts/ts-resolve-register.mjs scripts/corpus/measure.mjs [--excel] [--only text] [--fresh]
//
// For every workbook in corpus/files:
//   1. How does it open: editable, kept as stored, read only, not at all?
//   2. The thirty everyday actions (scripts/audit/refusals.mjs).
//   3. A round trip: type one number into an empty cell, save, and compare
//      everything else with the original using an independent reader.
//   4. With --excel: Microsoft Excel opens that saved file, recalculates and
//      saves it; the result is compared with what the editor saved.
//
// Results are kept per workbook in corpus/reports/scoreboard.json, so a run can
// be stopped and continued; --fresh starts again. The summary is printed and
// written to corpus/reports/scoreboard.md.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { actions, attempt, checks, opening, roundTrip, closeNativeWorkbookWorker } from '../audit/refusals.mjs';
import { excelAct, excelDone, excelFolder, excelOpenSave, excelVersion } from '../audit/excel.mjs';
import { pythonOnPath } from '../audit/python.mjs';
import { unzipSync } from 'fflate';

const python = pythonOnPath();
const component = new URL('../../', import.meta.url).pathname,
  folder = path.join(component, 'corpus/files'),
  reports = path.join(component, 'corpus/reports'),
  boardFile = path.join(reports, 'scoreboard.json'),
  args = process.argv.slice(2),
  only = args.includes('--only') ? args[args.indexOf('--only') + 1] : null,
  withExcel = args.includes('--excel') || args.includes('--excel-actions'),
  // --detail: check every action that worked, not only that it completed.
  // --excel-actions: also compare moving actions with Excel doing the same.
  detail = args.includes('--detail') || args.includes('--excel-actions'),
  excelActions = args.includes('--excel-actions'),
  CHECKING = 420000,
  LIMIT = 150000,
  BUDGET = 240000;
await fs.mkdir(reports, { recursive: true });
const board = args.includes('--fresh')
  ? { workbooks: {} }
  : JSON.parse(await fs.readFile(boardFile, 'utf8').catch(() => '{"workbooks":{}}'));
const scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-corpus-'));
const version = withExcel ? excelVersion() : null,
  excel = version ? await excelFolder('corpus-check') : null;
if (withExcel && !version) console.log('Microsoft Excel cannot be controlled from here; the Excel step is skipped.');

const within = (promise, label) =>
  Promise.race([
    promise,
    new Promise((resolve) =>
      setTimeout(
        () => resolve({ outcome: 'refused', stage: 'slow', message: `${label} took over ${LIMIT / 1000} s` }),
        LIMIT
      )
    ),
  ]);
const compare = (before, after, extra = []) => {
  try {
    return JSON.parse(
      execFileSync(python, [path.join(component, 'scripts/audit/compare-workbooks.py'), before, after, ...extra], {
        encoding: 'utf8',
        timeout: 600000,
        maxBuffer: 64 * 1024 * 1024,
      })
    );
  } catch (error) {
    return { same: null, unreadable: String(error?.message ?? error).slice(0, 200), differences: [], counts: {} };
  }
};

/**
 * Check one completed action in detail. A contained action must leave
 * everything but its own cells and its own kind of change as it was. A moving
 * action is compared with Excel performing the same action on the original;
 * the editor's file is first re-saved by Excel so that Excel's own tidying is
 * on both sides of the comparison.
 */
async function verify(spec, sheet, source, saved, at, linked) {
  if (spec.skip) return { done: false, reason: spec.skip };
  if (!spec.excel) {
    const result = compare(source, saved, [
      ...(spec.ignore ?? []).flatMap((area) => ['--ignore-range', `${sheet}!${area}`]),
      ...(spec.allow ?? []).flatMap((kind) => ['--allow', kind]),
    ]);
    return {
      done: result.same !== null,
      how: 'original',
      same: result.same,
      counts: result.counts,
      differences: result.differences,
    };
  }
  if (!excelActions || !excel) return { done: false, reason: 'not compared with Excel in this run' };
  if (linked) return { done: false, reason: 'links to other workbooks: Excel asks a person before opening' };
  const byExcel = path.join(excel, `w${at}-excel.xlsx`),
    byEditor = path.join(excel, `w${at}-editor.xlsx`);
  await fs.copyFile(source, byExcel);
  await fs.copyFile(saved, byEditor);
  try {
    const refused = excelAct(excel, byExcel, sheet, ...spec.excel);
    if (refused)
      return {
        done: true,
        how: 'excel',
        same: false,
        excelRefused: refused,
        counts: { 'Excel refuses this action': 1 },
      };
    const rejected = excelOpenSave(excel, byEditor);
    if (rejected)
      return {
        done: true,
        how: 'excel',
        same: false,
        counts: { 'Excel would not open the result': 1 },
        reason: rejected,
      };
    const result = compare(byExcel, byEditor, ['--layout-may-change']);
    return {
      done: result.same !== null,
      how: 'excel',
      same: result.same,
      counts: result.counts,
      differences: result.differences,
    };
  } finally {
    await fs.rm(byExcel, { force: true });
    await fs.rm(byEditor, { force: true });
  }
}

const names = (await fs.readdir(folder)).filter((name) => /\.xlsx$/i.test(name) && !name.startsWith('~$')).sort();
for (const [at, name] of names.entries()) {
  if (only && !name.includes(only)) continue;
  const done = board.workbooks[name];
  // A finished workbook is kept, unless it was named with --only: then it is measured again.
  if (!only && done?.actions && done.roundTrip && (!excel || done.excel) && (!detail || done.detailed)) continue;
  const source = path.join(scratch, `w${at}.xlsx`),
    index = path.join(scratch, 'index'),
    row = only ? {} : (done ?? {});
  await fs.copyFile(path.join(folder, name), source);
  const started = Date.now();
  row.opening ??= await opening(source, index);
  if (detail && !row.detailed) delete row.actions;
  if (!row.actions) {
    row.detailed = detail;
    row.actions = [];
    let slow = false,
      checked = 0;
    const kept = path.join(scratch, `w${at}-action.xlsx`),
      linked =
        Object.keys(
          unzipSync(await fs.readFile(source), { filter: (file) => file.name.startsWith('xl/externalLinks/') })
        ).length > 0;
    for (const [label, perform] of actions) {
      // A large workbook is not allowed to hold up the whole run: after
      // BUDGET seconds the remaining actions are recorded as not tried and
      // are left out of the percentages.
      const result = slow
        ? { outcome: 'refused', stage: 'slow', message: 'not tried: an earlier action on this workbook was too slow' }
        : Date.now() - started > BUDGET
          ? { outcome: 'skipped', stage: 'time', message: `not tried: the workbook used its ${BUDGET / 1000} s budget` }
          : await within(attempt(source, perform, index, detail ? kept : undefined), label);
      slow ||= result.stage === 'slow';
      const on = result.on;
      delete result.on;
      if (detail && result.outcome === 'ok' && on && checks[label]) {
        if (checked > CHECKING) result.check = { done: false, reason: 'the workbook used its checking time' };
        else {
          const began = Date.now();
          result.check = await verify(checks[label](on), on.name, source, kept, at, linked);
          checked += Date.now() - began;
        }
      }
      row.actions.push({ action: label, ...result });
    }
  }
  if (!row.roundTrip) {
    const saved = path.join(scratch, `w${at}-saved.xlsx`),
      edit = await roundTrip(source, index, saved);
    row.roundTrip = edit.refused
      ? { saved: false, reason: edit.refused }
      : {
          saved: true,
          edited: `${edit.sheet}!${edit.cell}`,
          ...compare(source, saved, ['--ignore', `${edit.sheet}!${edit.cell}`]),
        };
    // Excel on a Mac asks a person for access to a workbook that another
    // workbook links to, and a script cannot answer. Those are not sent.
    const linksOut = unzipSync(await fs.readFile(source), {
      filter: (file) => file.name.startsWith('xl/externalLinks/'),
    });
    if (excel && row.roundTrip.saved && Object.keys(linksOut).length)
      row.excel = { accepted: null, reason: 'links to other workbooks: Excel asks a person before opening' };
    else if (excel && row.roundTrip.saved) {
      const working = path.join(excel, `w${at}.xlsx`);
      await fs.copyFile(saved, working);
      const problem = excelOpenSave(excel, working);
      if (problem) {
        // Is it this file, or does Excel also stop on the original (a question
        // it asks a person, a file it cannot read)? Only the first is ours.
        const control = path.join(excel, `w${at}-original.xlsx`);
        await fs.copyFile(source, control);
        const originalProblem = excelOpenSave(excel, control);
        await fs.rm(control, { force: true });
        row.excel = { accepted: originalProblem ? null : false, reason: problem, originalAlso: !!originalProblem };
      } else {
        row.excel = { accepted: true, version, ...compare(saved, working, ['--layout-may-change']) };
        if (row.excel.same === false) {
          // Excel tidies files as it saves (fills in default fonts, re-encodes
          // pictures). Is this difference Excel's own, or the editor's doing?
          // Let Excel re-save the original as well and compare the two Excel
          // copies: what Excel does to both is not a loss caused here.
          const control = path.join(excel, `w${at}-original.xlsx`);
          await fs.copyFile(source, control);
          if (!excelOpenSave(excel, control)) {
            const against = compare(control, working, ['--layout-may-change', '--ignore', row.roundTrip.edited]);
            row.excel.sameAsExcelSavedOriginal = against.same;
            row.excel.againstOriginal = { counts: against.counts, differences: against.differences };
          }
          await fs.rm(control, { force: true });
        }
      }
      await fs.rm(working, { force: true });
    } else if (excel) row.excel = { accepted: null, reason: 'nothing was saved to open' };
    await fs.rm(saved, { force: true });
  }
  row.seconds = Math.round((Date.now() - started) / 1000);
  board.workbooks[name] = row;
  await fs.writeFile(boardFile, JSON.stringify(board, null, 1));
  const refused = row.actions.filter((item) => item.outcome === 'refused').length;
  console.log(
    `${String(at + 1).padStart(3)} ${name.slice(0, 46).padEnd(46)} ${String(actions.length - refused).padStart(2)}/${actions.length}` +
      `  round trip ${!row.roundTrip.saved ? 'not saved' : row.roundTrip.same === true ? 'same' : row.roundTrip.same === null ? 'unreadable' : 'DIFFERS'}` +
      (row.excel
        ? `  Excel ${row.excel.accepted === true ? (row.excel.same === false ? 'opens, differs' : 'ok') : row.excel.accepted === false ? 'REJECTS' : '-'}`
        : '') +
      `  ${row.seconds}s`
  );
  await fs.rm(source, { force: true });
}
closeNativeWorkbookWorker();
if (excel) await fs.rm(excel, { recursive: true, force: true });
excelDone();
await fs.rm(scratch, { recursive: true, force: true });

// ---- summary
const rows = Object.entries(board.workbooks).filter(([name]) => names.includes(name));
const tally = (items) => {
  const counted = new Map();
  for (const item of items) counted.set(item, (counted.get(item) ?? 0) + 1);
  return [...counted].sort((a, b) => b[1] - a[1]);
};
// Refusals Excel makes as well: the sheet is protected, or the cells are inside a table that cannot be merged.
const agrees = (item) =>
  item.outcome === 'refused' &&
  (item.code === 'sheet-protected' || /Excel does not allow it either/.test(item.message));
const attempts = rows.flatMap(([name, row]) => row.actions.map((item) => ({ name, ...item }))),
  refused = attempts.filter((item) => item.outcome === 'refused'),
  skipped = attempts.filter((item) => item.outcome === 'skipped'),
  worked = attempts.length - refused.length - skipped.length,
  tried = attempts.length - skipped.length,
  agreed = attempts.filter(agrees).length,
  share = (n) => ((n / tried) * 100).toFixed(1),
  trips = rows.map(([, row]) => row.roundTrip).filter(Boolean),
  excelRows = rows.map(([, row]) => row.excel).filter((item) => item && item.accepted !== null);
const lines = [];
const say = (text = '') => lines.push(text);
say(`# Scoreboard: ${rows.length} real workbooks`);
say();
say(`Measured ${new Date().toISOString().slice(0, 10)}.${version ? ` Microsoft Excel ${version}.` : ''}`);
say();
say('| Measure | Result |');
say('|---|---|');
say(
  `| Workbooks that open for editing | ${rows.filter(([, r]) => r.opening.opens && !r.opening.readOnly).length} of ${rows.length} |`
);
say(
  `| ...of which kept as stored (formulas not recalculated here) | ${rows.filter(([, r]) => r.opening.keptAsStored).length} |`
);
const count = (list, test) => list.filter(test).length,
  compared = count(trips, (t) => t.saved && t.same !== null);
say(`| Everyday actions that worked | ${worked} of ${tried} tried (${((worked / tried) * 100).toFixed(1)}%) |`);
// A refusal Excel makes too (a protected sheet, merging inside a table) is the right answer, not a gap.
say(
  `| ...or were refused exactly where Excel refuses too | ${worked + agreed} of ${tried} (${share(worked + agreed)}%) |`
);
say(
  `| Actions with nothing suitable on the sheet (not counted) | ${count(skipped, (item) => item.stage !== 'time')} |`
);
say(
  '| Actions not tried because a large workbook used its time budget (not counted) | ' +
    `${count(skipped, (item) => item.stage === 'time')} |`
);
const allWorked = count(rows, ([, r]) => !r.actions.some((item) => item.outcome === 'refused' && !agrees(item)));
say(`| Workbooks where all tried actions worked | ${allWorked} of ${rows.length} |`);
say(`| Round trip saved | ${count(trips, (t) => t.saved)} of ${trips.length} |`);
say(`| Round trip with nothing lost or changed | ${count(trips, (t) => t.same === true)} of ${compared} compared |`);
say(`| Round trip the independent reader could not compare | ${count(trips, (t) => t.saved && t.same === null)} |`);
if (excelRows.length) {
  const accepted = excelRows.filter((e) => e.accepted),
    matched = count(accepted, (e) => e.same === true || e.sameAsExcelSavedOriginal === true);
  say(
    `| Saved files Excel opened, recalculated and saved without repair | ${accepted.length} of ${excelRows.length} |`
  );
  say(
    "| ...and Excel's copy matched the editor's file, or matched Excel's own re-save of the original | " +
      `${matched} of ${count(accepted, (e) => e.same !== null)} compared |`
  );
}
const checkedRows = attempts.filter((item) => item.check?.done),
  byOriginal = checkedRows.filter((item) => item.check.how === 'original'),
  byExcel = checkedRows.filter((item) => item.check.how === 'excel');
if (checkedRows.length) {
  say();
  say('## Actions checked in detail');
  say();
  say('| Measure | Result |');
  say('|---|---|');
  say(
    `| Contained actions that changed only what they were meant to | ${byOriginal.filter((item) => item.check.same).length} of ${byOriginal.length} checked |`
  );
  if (byExcel.length)
    say(
      `| Moving actions whose result matched Excel doing the same action | ${byExcel.filter((item) => item.check.same).length} of ${byExcel.length} compared |`
    );
  say(
    '| Actions that worked but were not checked (time, or links to other workbooks) | ' +
      `${count(attempts, (item) => item.outcome === 'ok' && !item.check?.done)} |`
  );
  const off = new Map();
  for (const item of checkedRows.filter((entry) => !entry.check.same))
    for (const kind of Object.keys(item.check.counts ?? {})) {
      const key = `${item.action}\u0000${kind}`;
      off.set(key, [...(off.get(key) ?? []), item.name]);
    }
  if (off.size) {
    say();
    say('| Workbooks | Action | What differed |');
    say('|---:|---|---|');
    for (const [key, files] of [...off].sort((a, b) => b[1].length - a[1].length).slice(0, 40))
      say(`| ${new Set(files).size} | ${key.split('\u0000')[0]} | ${key.split('\u0000')[1]} |`);
  }
}
say();
say('## Refusals, most frequent first');
say();
say('| Workbooks | Action | What the editor said (first example) |');
say('|---:|---|---|');
const byAction = new Map();
for (const item of refused) byAction.set(item.action, [...(byAction.get(item.action) ?? []), item]);
for (const [action, items] of [...byAction].sort((a, b) => b[1].length - a[1].length))
  say(
    `| ${items.length} | ${action} | ${tally(items.map((item) => item.message.replace(/\|/g, '/').slice(0, 110)))[0][0]} |`
  );
say();
say('## The same refusals by message');
say();
say('| Times | Message |');
say('|---:|---|');
for (const [message, count] of tally(
  refused.map((item) =>
    item.message
      .replace(/\|/g, '/')
      .replace(/[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?/g, 'RANGE')
      .slice(0, 120)
  )
).slice(0, 25))
  say(`| ${count} | ${message} |`);
const lost = tally(trips.filter((t) => t.same === false).flatMap((t) => Object.keys(t.counts)));
say();
say('## Round trip: what differed');
say();
if (!lost.length) say('Nothing.');
else {
  say('| Workbooks | What differed |');
  say('|---:|---|');
  for (const [kind, count] of lost) say(`| ${count} | ${kind} |`);
}
const notSaved = tally(trips.filter((t) => !t.saved).map((t) => String(t.reason).slice(0, 120)));
if (notSaved.length) {
  say();
  say('## Round trip: not saved');
  say();
  say('| Workbooks | Why |');
  say('|---:|---|');
  for (const [reason, count] of notSaved) say(`| ${count} | ${reason.replace(/\|/g, '/')} |`);
}
if (excelRows.length) {
  const rejected = tally(excelRows.filter((e) => !e.accepted).map((e) => String(e.reason).slice(0, 120)));
  const changed = tally(
    excelRows
      .filter((e) => e.accepted && e.same === false && e.sameAsExcelSavedOriginal !== true)
      .flatMap((e) => Object.keys(e.againstOriginal?.counts ?? e.counts))
  );
  say();
  say('## Excel');
  say();
  if (!rejected.length) say('Excel opened every saved file without repair.');
  else {
    say('| Workbooks | Excel said |');
    say('|---:|---|');
    for (const [reason, count] of rejected) say(`| ${count} | ${reason.replace(/\|/g, '/')} |`);
  }
  if (changed.length) {
    say();
    say('What differed after Excel re-saved the file, beyond what Excel also changes in the original:');
    say();
    say('| Workbooks | What differed |');
    say('|---:|---|');
    for (const [kind, count] of changed) say(`| ${count} | ${kind} |`);
  }
}
await fs.writeFile(path.join(reports, 'scoreboard.md'), lines.join('\n') + '\n');
console.log('\n' + lines.join('\n'));
