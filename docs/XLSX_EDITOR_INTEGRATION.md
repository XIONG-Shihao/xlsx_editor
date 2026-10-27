---
status: current
---

# Where the spreadsheet editor is used

The editor has one user:

| User | What it adds | Where |
|---|---|---|
| The desktop app | A window, menus, file dialogs, tabs, printing | `desktop/` |

The editor is in no other app.

The standalone app is built with Tauri 2: a Rust host owns the window and every child process, and
the JavaScript parts run on a Bun runtime shipped inside the app.
Nothing that ships uses Electron or Node.js.

## What the editor offers a host

| Entry point | For |
|---|---|
| `@office-editors/xlsx-editor` and `/ui` | The grid and ribbon as React components |
| `/host` | The pieces a host supplies: icons, clipboard, language, callbacks that fetch cells |
| `/contracts` | The shapes of commands and replies |
| `core/indexed-session.mjs` | Open a workbook, apply commands, read cells, save: no screen needed |
| `adapters/document-session.mjs` | The same with undo, recovery and safe saving, one worker per document |
| `native/` (`workbook-engine`) | The engine program the session talks to |

A host must supply: where files are, who may edit them, where recovery data is kept, and how the
screen talks to the session. The editor must stay buildable and testable with none of the host code
present.

Request rules and command capture stay in the editor (`adapters/owner-request.mjs`); a new command
family is added there once.

## The standalone spreadsheet app

`desktop/` is a small Tauri app around the editor.

| Part | Role |
|---|---|
| `src-tauri/` | Rust host: window, menus, file dialogs, starts the backend; `src/host/` supervises the backend process and holds the profile lease |
| `backend/` | Runs the editor's own shell code (`standalone/main.mjs`) on Bun; `backend/host/` is its own connection to the Rust host |
| `frontend/` | Loads the editor's own screen code (`standalone/renderer.tsx`) and connects it to the backend |
| `scripts/build.mjs`, `scripts/package.mjs` | Build and package; `scripts/tooling/` holds the pinned Bun's download and checks, staging and native-code signing |

The coordinator in `standalone/` asks for native desktop ports through
`standalone/platform.mjs`. A host calls `setStandalonePlatform` before importing
`main.mjs`. `desktop/backend/platform.ts` supplies windows, menus, dialogs,
links, PDF generation and printing on top of Tauri. Both print implementations
are required; there is no Electron fallback.

The frontend creates `window.xlsxDesktop` with the component's
`standalone/bridge.mjs` API factory and the native host's invocation/subscription
ports. No preload, Electron module substitution or editor source rewrite is
needed. `verify-coordinator.mjs` checks that every bridge operation is available.
Nothing in this folder uses anything outside it (`tests/product-folder.mjs`).

It leaves out everything not needed for spreadsheets: agents, Python, the Word editor.

Build: `npm run desktop:package`. Result:
`desktop/release/mac-arm64/internal_release/XLSX Editor.app`. Packaging downloads the pinned Bun
once and builds the editor from its own lockfile first (`desktop/scripts/build-editor.mjs`). The
packaged app's `app/` folder mirrors this folder (`app/desktop/out/main`, `app/core`, ...).

## After a change

Run the editor's tests, then rebuild the standalone app.

A change whose app is not rebuilt is not in any app. The build steps are in
[DEVELOPMENT.md](DEVELOPMENT.md).
