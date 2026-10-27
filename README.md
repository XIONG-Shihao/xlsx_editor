# Spreadsheet editor

A spreadsheet editor for `.xlsx` workbooks and `.csv`/`.tsv` files that runs on the user's own
computer, and the desktop app around it (`desktop/`). This folder is the whole product: it installs,
builds, tests and packages from here alone, with its own lockfile, pinned Bun and scripts.

**Goal:** everyday spreadsheet work on real Excel files, with nothing lost going to or from Excel.
It is not a copy of Excel. The goal, where it stands and what comes next are in
[docs/README.md](docs/README.md).

## Read

| Document | For |
|---|---|
| [docs/README.md](docs/README.md) | The goal and the scoreboard |
| [docs/ROADMAP.md](docs/ROADMAP.md) | The next big steps |
| [docs/FEATURES.md](docs/FEATURES.md) | What a person can do, tab by tab |
| [docs/BIDIRECTIONAL_COMPATIBILITY.md](docs/BIDIRECTIONAL_COMPATIBILITY.md) | The promise about Excel, and how it is checked |
| [docs/XLSX_EDITOR.md](docs/XLSX_EDITOR.md) | How the editor is built, and its limits |
| [docs/XLSX_EDITOR_INTEGRATION.md](docs/XLSX_EDITOR_INTEGRATION.md) | How the standalone app uses it |
| [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) | Building, testing and measuring |

## Layout

```text
ui/          grid, ribbon and panels (React)
core/        session layer: commands, undo, what to refuse (JavaScript)
adapters/    one worker per document, recovery, saving
native/      the engine: storage, calculation, writing (Rust, program workbook-engine)
standalone/  shell and screen code the desktop app runs
desktop/     the desktop app: Tauri window, native operations, its Bun coordinator and packaging
fixtures/    test workbooks, including ones calculated by Excel
tests/       tests
scripts/     build, measuring and timing tools
docs/        documentation
```

## Build and test

Needs Node.js 24, Rust and Python 3 (see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md#what-you-need)).

```sh
npm run desktop:build   # downloads the pinned Bun once, installs from bun.lock, builds the editor and the app
npm run typecheck
npm run test:all        # every test file, several at a time
npm run gates           # everything above, the desktop shell's Rust tests and the coordinator check
```

## Measure

```sh
# how many of 30 everyday actions work, and what the refusals say
node --import ./scripts/ts-resolve-register.mjs scripts/audit/refusals.mjs [FILE or FOLDER...]

# does real Excel accept the files the editor saves (Mac with Excel; opens Excel on screen)
node --import ./scripts/ts-resolve-register.mjs scripts/audit/excel-check.mjs
```

## The apps

The editor is not an app by itself. Its one app is the desktop app:

```sh
npm run desktop:package   # desktop/release/mac-arm64/internal_release/XLSX Editor.app
```

The coordinator, renderer and frontend bridge in `standalone/` are host-neutral.
`desktop/` supplies the Tauri window, native operations and contained Bun.
See [host integration](docs/XLSX_EDITOR_INTEGRATION.md#the-standalone-spreadsheet-app).
