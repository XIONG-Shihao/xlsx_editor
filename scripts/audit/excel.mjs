// Driving real Microsoft Excel for Mac from a script.
//
// Excel is sandboxed and asks a person for access to any file outside its own
// container; inside the container it opens files without a prompt. Everything
// here therefore works in a folder under that container.
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

const container = path.join(os.homedir(), 'Library/Containers/com.microsoft.Excel/Data/Documents');

// Whether Excel was already open when this process first touched it. If it was,
// it may hold someone's work and is never closed from here.
let startedHere = null;
const running = () => spawnSync('pgrep', ['-x', 'Microsoft Excel']).status === 0;

/** Excel's version, or null when Excel cannot be controlled from here. */
export function excelVersion() {
  startedHere ??= !running();
  const reply = spawnSync('osascript', ['-e', 'tell application "Microsoft Excel" to get version'], {
    encoding: 'utf8',
    timeout: 60000,
  });
  return reply.status === 0 ? reply.stdout.trim() : null;
}

/**
 * Leave Excel as it was found. An Excel this process started is asked to quit
 * (not killed) when the work is done; an Excel that was already open is left
 * open. Call once at the end of a run.
 */
export function excelDone() {
  if (startedHere && running())
    spawnSync('osascript', ['-e', 'tell application "Microsoft Excel" to quit saving no'], { timeout: 30000 });
}

/** A new folder Excel can open files in without asking. The caller removes it. */
export async function excelFolder(label) {
  if (!existsSync(container))
    throw Error('Microsoft Excel has not been run on this Mac (its container folder is missing)');
  const folder = path.join(container, `workbook-${label}-${new Date().toISOString().replace(/[:.]/g, '-')}`);
  await fs.mkdir(folder, { recursive: true });
  await fs.writeFile(
    path.join(folder, 'open-save.applescript'),
    `on run argv
 set ownedFile to item 1 of argv
 set ownedBook to missing value
 tell application "Microsoft Excel"
  with timeout of 90 seconds
   try
    set display alerts to false
    set ask to update links to false
    open (POSIX file ownedFile)
    set matches to every workbook whose full name is ownedFile
    if (count of matches) is not 1 then error "Excel did not open the file under its own name (it may have repaired it)"
    set ownedBook to item 1 of matches
    calculate
    save ownedBook
    close ownedBook saving no
    return "saved"
   on error errText number errNumber
    if ownedBook is not missing value then
     try
      close ownedBook saving no
     end try
    end if
    error errText number errNumber
   end try
  end timeout
 end tell
end run
`
  );
  await fs.writeFile(path.join(folder, 'do.applescript'), ACTION_SCRIPT);
  return folder;
}

/**
 * Have Excel open a file (which must be inside an excelFolder), recalculate it
 * and save it in place. Returns null on success or what went wrong. A file
 * Excel has to repair opens under another name and is reported as a failure.
 */
export function excelOpenSave(folder, file) {
  const reply = spawnSync('osascript', [path.join(folder, 'open-save.applescript'), file], {
    encoding: 'utf8',
    timeout: 120000,
  });
  if (reply.status === 0) return null;
  // A timeout usually means Excel is showing a question only a person can
  // answer, and it will block every later file. Close the Excel this process
  // started; one that was already open is left alone.
  if (/timed out/i.test(reply.stderr || 'timed out') && startedHere) {
    spawnSync('pkill', ['-x', 'Microsoft Excel']);
    spawnSync('/bin/sleep', ['3']);
  }
  return (reply.stderr || 'timed out, possibly on a prompt that needs a person')
    .trim()
    .replace(/^.*execution error: /, '')
    .slice(0, 240);
}

// One action performed by Excel itself, so that the editor's result for the
// same action can be compared with what Excel does.
const ACTION_SCRIPT = `-- Perform one action in Microsoft Excel on a workbook and save it.
-- argv: file, sheet name, action, then the action's arguments.
on run argv
 set f to item 1 of argv
 set sheetName to item 2 of argv
 set what to item 3 of argv
 tell application "Microsoft Excel"
  with timeout of 90 seconds
   set display alerts to false
   open (POSIX file f)
   set b to active workbook
   set s to worksheet sheetName of b
   try
    if what is "insertRow" then
     insert into range (row ((item 4 of argv) as integer) of s) shift shift down
    else if what is "deleteRow" then
     delete range (row ((item 4 of argv) as integer) of s) shift shift up
    else if what is "insertColumn" then
     insert into range (column ((item 4 of argv) as integer) of s) shift shift to right
    else if what is "deleteColumn" then
     delete range (column ((item 4 of argv) as integer) of s) shift shift to left
    else if what is "sort" then
     if (item 6 of argv) is "desc" then
      sort (range (item 4 of argv) of s) key1 (range (item 5 of argv) of s) order1 sort descending header header no
     else
      sort (range (item 4 of argv) of s) key1 (range (item 5 of argv) of s) order1 sort ascending header header no
     end if
    else if what is "rename" then
     set name of s to (item 4 of argv)
    else if what is "duplicate" then
     copy worksheet s after (sheet (count of sheets of b) of b)
     set name of active sheet to (item 4 of argv)
    else if what is "merge" then
     merge (range (item 4 of argv) of s)
    else if what is "fillDown" then
     fill down (range (item 4 of argv) of s)
    else
     error "unknown action " & what
    end if
    save b
    close b saving no
    return "done"
   on error errText number errNumber
    try
     close b saving no
    end try
    error errText number errNumber
   end try
  end timeout
 end tell
end run
`;

/**
 * Have Excel perform one action on a workbook (inside an excelFolder) and save
 * it. `action` is insertRow, deleteRow, insertColumn, deleteColumn (1-based
 * number), sort (range, key cell, asc|desc), rename (new name), duplicate (new
 * name), merge (range) or fillDown (range). Returns null on success or what
 * Excel said; Excel refusing an action is a result, not a failure of the tool.
 */
export function excelAct(folder, file, sheet, action, ...args) {
  const reply = spawnSync(
    'osascript',
    [path.join(folder, 'do.applescript'), file, sheet, action, ...args.map(String)],
    {
      encoding: 'utf8',
      timeout: 150000,
    }
  );
  if (reply.status === 0) return null;
  if (/timed out/i.test(reply.stderr || 'timed out') && startedHere) {
    spawnSync('pkill', ['-x', 'Microsoft Excel']);
    spawnSync('/bin/sleep', ['3']);
  }
  return (reply.stderr || 'timed out')
    .trim()
    .replace(/^.*execution error: /, '')
    .slice(0, 240);
}
