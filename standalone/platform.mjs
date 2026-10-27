// Native desktop ports for the standalone editor coordinator. A host sets them
// once before importing main.mjs, including its PDF and printer implementations.
import { refusal } from '../core/refusal-codes.mjs';

const required = [
  'app', 'BrowserWindow', 'clipboard', 'dialog', 'ipcMain', 'Menu', 'session', 'shell',
  'createWorkbookPdfRenderer', 'createWorkbookNativePrinter',
];
let current;

export function setStandalonePlatform(platform) {
  if (current) throw refusal('internal-mismatch', 'The standalone platform is already set.');
  const missing = required.filter((name) => !platform?.[name]);
  if (missing.length) throw refusal('internal-missing-field', `The standalone platform lacks ${missing.join(', ')}.`);
  current = platform;
}

export function standalonePlatform() {
  if (!current) throw refusal('internal-missing-field', 'Set the standalone platform before importing main.mjs.');
  return current;
}
