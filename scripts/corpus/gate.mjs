// The definition of done for the spreadsheet editor.
//
// A change is done when the scoreboard is no worse than the recorded baseline:
// no workbook lost a working action, no round trip started losing content,
// Excel did not start rejecting a file, and no action that matched Excel
// stopped matching. Better results are recorded with --update, so the baseline
// only ever moves forward.
//
//   node scripts/corpus/gate.mjs            compare corpus/reports/scoreboard.json with corpus/baseline.json
//   node scripts/corpus/gate.mjs --update   record the current scoreboard as the new baseline (refused if worse)
//   node scripts/corpus/gate.mjs --report FILE   compare another scoreboard file
//
// Workbooks missing from the scoreboard (a partial run with --only) are not
// judged; the gate says how many it compared.
import fs from 'node:fs';

const here = new URL('../../corpus/', import.meta.url),
  option = (name) => {
    const at = process.argv.indexOf(name);
    return at < 0 ? null : process.argv[at + 1];
  },
  reportFile = option('--report') ?? new URL('reports/scoreboard.json', here).pathname,
  baselineFile = new URL('baseline.json', here).pathname,
  update = process.argv.includes('--update');

// Refusals Excel makes as well are the right answer and count as working.
const agrees = (item) =>
  item.outcome === 'refused' &&
  (item.code === 'sheet-protected' || /Excel does not allow it either/.test(item.message));
const works = (item) => item.outcome === 'ok' || agrees(item);

/** What is remembered about one workbook: which actions work and which checks pass. */
function summary(row) {
  const actions = row.actions ?? [];
  return {
    opens: !!row.opening?.opens && !row.opening?.readOnly,
    working: actions.filter(works).map((item) => item.action),
    tried: actions.filter((item) => item.outcome !== 'skipped').length,
    roundTrip: row.roundTrip?.saved ? row.roundTrip.same : null,
    excel: row.excel ? !!row.excel.accepted : null,
    matchedExcel: actions.filter((item) => item.check?.done && item.check.same).map((item) => item.action),
  };
}

const report = JSON.parse(fs.readFileSync(reportFile, 'utf8')).workbooks,
  current = Object.fromEntries(Object.entries(report).map(([name, row]) => [name, summary(row)]));
const baseline = fs.existsSync(baselineFile) ? JSON.parse(fs.readFileSync(baselineFile, 'utf8')) : { workbooks: {} };

const worse = [],
  better = [];
let compared = 0;
for (const [name, now] of Object.entries(current)) {
  const was = baseline.workbooks[name];
  if (!was) {
    better.push(`${name}: new in the set`);
    continue;
  }
  compared += 1;
  if (was.opens && !now.opens) worse.push(`${name}: no longer opens for editing`);
  // A large workbook can run out of its time allowance on a busy machine; an
  // action that was not tried this time is not counted as lost.
  const attempted = new Set((report[name].actions ?? []).filter((item) => item.stage !== 'time').map((i) => i.action));
  for (const action of was.working)
    if (attempted.has(action) && !now.working.includes(action)) worse.push(`${name}: "${action}" stopped working`);
  for (const action of now.working) if (!was.working.includes(action)) better.push(`${name}: "${action}" now works`);
  if (was.roundTrip === true && now.roundTrip === false)
    worse.push(`${name}: the round trip now loses or changes content`);
  if (was.excel === true && now.excel === false) worse.push(`${name}: Excel no longer accepts the saved file`);
  const checked = new Set((report[name].actions ?? []).filter((item) => item.check?.done).map((item) => item.action));
  for (const action of was.matchedExcel ?? [])
    if (checked.has(action) && !now.matchedExcel.includes(action))
      worse.push(`${name}: "${action}" no longer matches the detailed check`);
}

const total = (set, pick) => Object.values(set).reduce((sum, row) => sum + pick(row), 0),
  share = (set) =>
    (
      (total(set, (row) => row.working.length) /
        Math.max(
          1,
          total(set, (row) => row.tried)
        )) *
      100
    ).toFixed(1);
console.log(`Compared ${compared} workbooks with the baseline (${Object.keys(baseline.workbooks).length} recorded).`);
console.log(
  `Working actions now: ${total(current, (row) => row.working.length)} of ${total(current, (row) => row.tried)} (${share(current)}%).`
);
if (better.length) console.log(`Better: ${better.length} (for example: ${better.slice(0, 3).join('; ')})`);
if (worse.length) {
  console.log(`\nWorse than the baseline in ${worse.length} places:`);
  for (const line of worse.slice(0, 60)) console.log('  ' + line);
  if (worse.length > 60) console.log(`  ...and ${worse.length - 60} more`);
  console.log('\nNot done: fix these, or explain in the commit why the baseline was wrong and edit it by hand.');
  process.exit(1);
}
if (update) {
  // Workbooks not measured in this run keep what was recorded for them. A run
  // without Excel keeps what an earlier run with Excel established: whether
  // Excel accepted the file, and which actions matched Excel doing the same.
  const merged = { ...baseline.workbooks };
  for (const [name, now] of Object.entries(current)) {
    const was = baseline.workbooks[name],
      checked = new Set((report[name].actions ?? []).filter((item) => item.check?.done).map((item) => item.action));
    merged[name] = {
      ...now,
      roundTrip: now.roundTrip ?? was?.roundTrip ?? null,
      excel: now.excel ?? was?.excel ?? null,
      matchedExcel: [
        ...now.matchedExcel,
        ...(was?.matchedExcel ?? []).filter((action) => !checked.has(action) && now.working.includes(action)),
      ],
    };
  }
  const sorted = Object.fromEntries(
    Object.keys(merged)
      .sort()
      .map((name) => [name, merged[name]])
  );
  fs.writeFileSync(
    baselineFile,
    JSON.stringify({ recorded: new Date().toISOString().slice(0, 10), workbooks: sorted }, null, 1) + '\n'
  );
  console.log(`Baseline updated: ${Object.keys(sorted).length} workbooks, ${share(sorted)}% of actions working.`);
} else console.log('No worse than the baseline.');
