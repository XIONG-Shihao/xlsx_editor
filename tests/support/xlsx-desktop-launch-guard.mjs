/** Refuse a second native spreadsheet owner without acting on unrelated apps. */
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
export function desktopLaunchInventory(rows) {
  const blockers = [],
    foreignFixtures = [];
  for (const row of rows) {
    if (row.executable.endsWith('/XLSX Editor') || row.executable.endsWith('/xlsx-desktop') || row.executable.endsWith('/workbook-engine')) {
      blockers.push(row);
      continue;
    }
  }
  return { blockers, foreignFixtures };
}
export function assertDesktopLaunchAllowed() {
  const read = (format) =>
    new Map(
      execFileSync('/bin/ps', ['-axo', format], { encoding: 'utf8' })
        .split('\n')
        .flatMap((line) => {
          const m = line.trim().match(/^(\d+)\s+(.+)$/);
          return m ? [[Number(m[1]), m[2]]] : [];
        })
    );
  const commands = read('pid=,command='),
    executables = read('pid=,comm=');
  const inventory = desktopLaunchInventory(
    [...executables].map(([pid, executable]) => ({ pid, executable, command: commands.get(pid) }))
  );
  assert.deepEqual(
    inventory.blockers,
    [],
    'Another XLSX editor/native owner is running; preserve it and do not launch another editor.'
  );
  return inventory.foreignFixtures;
}
