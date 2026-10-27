import { randomUUID } from 'node:crypto';
import { refusal } from '../core/refusal-codes.mjs';

/** One-window capabilities. A tab is stable; a replaced worker gets a new handle. */
export class WorkbookTabRegistry {
  #records = new Map();
  #active;
  constructor(limit = 4) {
    if (!Number.isInteger(limit) || limit < 1 || limit > 4)
      throw refusal('internal-invalid-field', 'Invalid workbook tab limit.');
    this.limit = limit;
  }
  get size() {
    return this.#records.size;
  }
  get active() {
    return this.#records.get(this.#active) ?? null;
  }
  records() {
    return [...this.#records.values()];
  }
  contains(record) {
    return this.#records.get(record?.tabId) === record;
  }
  byFilename(filename) {
    return this.records().find((record) => record.filename === filename) ?? null;
  }
  byDocument(documentId) {
    return this.records().find((record) => record.status.documentId === documentId) ?? null;
  }
  selected(id) {
    const record = this.records().find((record) => record.id === id);
    if (!record) throw refusal('session-bad-input', 'Document access is not granted.');
    return record;
  }
  assertRoom() {
    if (this.size >= this.limit)
      throw refusal('session-refused', 'Four workbooks are already open. Close a workbook before opening another.');
  }
  #admit(record, previous) {
    if (!record || typeof record.id !== 'string' || !record.owner || !record.status?.documentId)
      throw refusal('internal-invalid-field', 'Invalid workbook owner record.');
    for (const existing of this.records())
      if (existing !== previous) {
        if (
          existing.id === record.id ||
          existing.status.documentId === record.status.documentId ||
          (record.filename && existing.filename === record.filename)
        )
          throw refusal('session-duplicate', 'This workbook already has an open owner.');
      }
  }
  add(record) {
    this.assertRoom();
    this.#admit(record);
    record.tabId = randomUUID();
    record.epoch = 1;
    this.#records.set(record.tabId, record);
    this.#active = record.tabId;
    return record;
  }
  replace(previous, next) {
    if (!this.contains(previous)) throw refusal('session-stale', 'The workbook owner changed.');
    this.#admit(next, previous);
    next.tabId = previous.tabId;
    next.epoch = previous.epoch + 1;
    this.#records.set(next.tabId, next);
    return next;
  }
  activate(id) {
    const record = this.selected(id);
    this.#active = record.tabId;
    return record;
  }
  remove(record) {
    if (!this.contains(record)) throw refusal('session-stale', 'The workbook owner changed.');
    const order = this.records(),
      index = order.indexOf(record);
    this.#records.delete(record.tabId);
    if (this.#active === record.tabId) this.#active = (order[index + 1] ?? order[index - 1])?.tabId;
    return this.active;
  }
  assertDestination(record, filename) {
    const other = this.byFilename(filename);
    if (other && other !== record)
      throw refusal(
        'session-unsupported',
        'This destination is open in another workbook tab. Choose a different filename.'
      );
  }
}
