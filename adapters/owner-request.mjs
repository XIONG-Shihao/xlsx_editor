// One set of request rules for every workbook owner: the editor's own document
// session and an owner run by a host. A new command family that needs capture
// or normalization is added here once, and both owners follow.
import { captureInkStrokeCommands, isInkStrokeCommand } from '../core/ink-stroke.mjs';
import { accessibilityFields } from '../core/accessibility.mjs';
import { notesFields, captureNoteCommands, isNoteCommand } from '../core/notes.mjs';
import { captureModel3DCommands, isModel3DCommand } from '../core/model3d-authoring.mjs';
import { captureSmartArtCommands, isSmartArtCommand } from '../core/smartart-model.mjs';
import { captureShapeCommands, isShapeCommand } from '../core/shape-model.mjs';
import { capturePivotCommands, isIndexedPivotCommand } from '../core/indexed-pivots.mjs';
import { normalizeBackgroundCommands } from '../core/background-command.mjs';
import { normalizePageThemeCommands } from '../core/page-theme-command.mjs';
import { normalizeBooleanCellCommands } from '../core/cell-insertions.mjs';
import { hyperlinkFields, normalizeHyperlinkCommands } from '../core/hyperlinks.mjs';
import { normalizeClearHyperlinksCommands } from '../core/clear-hyperlinks.mjs';
import { ownerCommands, decodedCommands } from '../dist/core/transport.js';

/**
 * Checks and normalizes one owner request before it reaches a workbook.
 *
 * `transport` is true for requests that crossed a host transport: they carry
 * `kind` and `sessionId`, and an apply may carry `encodedCommands`. Such an
 * apply is decoded first and, when a rule changed it, encoded again.
 */
export function normalizeOwnerRequest(request, { transport = false } = {}) {
  const wire = transport && request.operation === 'apply' && typeof request.encodedCommands === 'string';
  const original = wire ? Object.freeze({ ...withoutCommands(request), commands: decodedCommands(request) }) : request;
  const normalized = normalize(original, transport);
  if (!transport || request.operation !== 'apply' || normalized === original) return wire ? request : normalized;
  return Object.freeze({ ...withoutCommands(normalized), ...ownerCommands(normalized.commands) });
}

function withoutCommands(request) {
  const { commands: _commands, encodedCommands: _encoded, ...rest } = request;
  return rest;
}

function normalize(request, transport) {
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => ['createChart', 'updateChart', 'deleteObject'].includes(c?.operation))
  )
    request = structuredClone(request);
  if (request.operation === 'apply' && Array.isArray(request.commands) && request.commands.some(isModel3DCommand))
    request = Object.freeze({ ...structuredClone(request), commands: captureModel3DCommands(request.commands) });
  if (request.operation === 'apply' && Array.isArray(request.commands) && request.commands.some(isSmartArtCommand))
    request = Object.freeze({ ...structuredClone(request), commands: captureSmartArtCommands(request.commands) });
  if (request.operation === 'apply' && Array.isArray(request.commands) && request.commands.some(isInkStrokeCommand))
    request = Object.freeze({ ...structuredClone(request), commands: captureInkStrokeCommands(request.commands) });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => isShapeCommand(c) || c?.operation === 'deleteObject')
  )
    request = Object.freeze({ ...structuredClone(request), commands: captureShapeCommands(request.commands) });
  if (request.operation === 'apply' && Array.isArray(request.commands) && request.commands.some(isIndexedPivotCommand))
    request = Object.freeze({ ...structuredClone(request), commands: capturePivotCommands(request.commands) });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => ['renameTableColumn', 'insertTableColumn', 'deleteTableColumn'].includes(c?.operation))
  )
    request = Object.freeze(structuredClone(request));
  if (request.operation === 'accessibility') {
    accessibilityFields(request, { transport });
    request = Object.freeze(structuredClone(request));
  }
  if (request.operation === 'notes') {
    notesFields(request, { transport });
    request = Object.freeze(structuredClone(request));
  }
  if (request.operation === 'apply' && Array.isArray(request.commands) && request.commands.some(isNoteCommand))
    request = Object.freeze({ ...structuredClone(request), commands: captureNoteCommands(request.commands) });
  if (request.operation === 'hyperlink') {
    hyperlinkFields(request, { transport });
    request = Object.freeze(structuredClone(request));
  }
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => ['setSheetBackground', 'removeSheetBackground'].includes(c?.type))
  )
    request = Object.freeze({ ...structuredClone(request), commands: normalizeBackgroundCommands(request.commands) });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => c?.type === 'applyPageTheme')
  )
    request = Object.freeze({ ...structuredClone(request), commands: normalizePageThemeCommands(request.commands) });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => c?.type === 'booleanCell')
  )
    request = Object.freeze({
      ...structuredClone(request),
      commands: normalizeBooleanCellCommands(request.commands),
    });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((c) => ['setHyperlink', 'removeHyperlink'].includes(c?.type))
  )
    request = Object.freeze({ ...structuredClone(request), commands: normalizeHyperlinkCommands(request.commands) });
  if (
    request.operation === 'apply' &&
    Array.isArray(request.commands) &&
    request.commands.some((command) => command?.type === 'clearHyperlinks')
  )
    request = Object.freeze({
      ...structuredClone(request),
      commands: normalizeClearHyperlinksCommands(request.commands),
    });
  if (request.operation === 'usedRange' || request.operation === 'noteMembership')
    request = Object.freeze({ ...request });
  if (request.operation === 'duplicatePreview')
    request = Object.freeze({
      ...request,
      columns: Array.isArray(request.columns) ? [...request.columns] : request.columns,
    });
  return request;
}
