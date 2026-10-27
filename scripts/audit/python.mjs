// The Python that the audit tools and independent-reader tests use.
//
// They need Python 3.10 or newer (zip(strict=...)) with openpyxl. In order:
// XLSX_EDITOR_TEST_PYTHON when set (the repository's test runner sets it to its
// uv environment); otherwise a private uv environment under .build/python-audit,
// made with `uv` on PATH or the uv named by XLSX_EDITOR_UV; only without uv, a
// qualifying python3 on PATH, or a venv made from one with pip.
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

const component = new URL('../../', import.meta.url).pathname,
  home = path.join(component, '.build/python-audit'),
  owned = path.join(home, 'bin/python3');
const run = (program, args) => spawnSync(program, args, { encoding: 'utf8' });
const recent = (program) =>
  run(program, ['-c', 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)']).status === 0;
const works = (program) => recent(program) && run(program, ['-c', 'import openpyxl']).status === 0;
const uv = () => [process.env.XLSX_EDITOR_UV, 'uv'].find((program) => program && run(program, ['--version']).status === 0);
let found;

function create(uvProgram) {
  console.error('Preparing a private Python environment in ' + home);
  const base = uvProgram ? null : ['python3', 'python3.13', 'python3.12', 'python3.11', 'python3.10'].find(recent);
  const steps = uvProgram
    ? [
        [uvProgram, ['venv', '--clear', '--python', '3.12', home]],
        [uvProgram, ['pip', 'install', '--quiet', '--python', owned, 'openpyxl']],
      ]
    : base
      ? [
          [base, ['-m', 'venv', '--clear', home]],
          [owned, ['-m', 'pip', 'install', '--quiet', '--disable-pip-version-check', 'openpyxl']],
        ]
      : null;
  if (!steps)
    throw Error(
      'The XLSX tests need uv, or Python 3.10 or newer with openpyxl, and neither was found (macOS python3 is 3.9). ' +
        'Install uv, or point XLSX_EDITOR_TEST_PYTHON at a Python that has openpyxl.'
    );
  for (const [program, args] of steps) {
    const step = run(program, args);
    if (step.status !== 0)
      throw Error(
        'Could not prepare Python with openpyxl. Point XLSX_EDITOR_TEST_PYTHON at a Python 3.10+ that has it.\n' +
          (step.stderr || step.error?.message || '').trim().slice(-600)
      );
  }
  return owned;
}

export function auditPython() {
  if (found) return found;
  const chosen = process.env.XLSX_EDITOR_TEST_PYTHON;
  if (chosen) {
    if (!works(chosen)) throw Error(`XLSX_EDITOR_TEST_PYTHON (${chosen}) is not Python 3.10+ with openpyxl`);
    return (found = chosen);
  }
  if (fs.existsSync(owned) && works(owned)) return (found = owned);
  const uvProgram = uv();
  if (uvProgram) return (found = create(uvProgram));
  if (works('python3')) return (found = 'python3');
  return (found = create());
}

/** Put that Python first on PATH, so every `python3` this process starts has openpyxl. */
export function pythonOnPath() {
  const program = auditPython();
  if (program !== 'python3') process.env.PATH = path.dirname(program) + path.delimiter + process.env.PATH;
  return program;
}
