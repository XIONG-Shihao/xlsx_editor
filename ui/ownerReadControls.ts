import type {
  WorkbookAccessibilityControls,
  WorkbookNoteMembershipControls,
  WorkbookNotesControls,
  WorkbookUsedRangeControls,
} from '../core/contracts';
import { validateAccessibility } from '../core/accessibility.mjs';
import { validateNotes } from '../core/notes.mjs';
import { noteMembershipReceipt, noteMembershipSheet, validateNoteMembership } from '../core/note-membership.mjs';
import { usedRangeSheet, validateUsedRange } from '../core/used-range.mjs';

/** The workbook state a set of read controls answers for. */
export interface OwnerReadReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}

export type OwnerRead = 'accessibility' | 'notes' | 'noteMembership' | 'usedRange';

/** What a host gives the editor to read from its workbook owner. */
export interface OwnerReadBinding {
  receipt: OwnerReadReceipt;
  /** Throws OwnerCommandRetired when the shown workbook is no longer the one the receipt names. */
  current: (operation: OwnerRead) => void;
  /** One owner read; the host adds its own routing. Resolves with the owner's reply. */
  request: (operation: OwnerRead, fields: { documentId: string; state: string; sheet?: string }) => Promise<any>;
}

/** The grid's read-only owner controls (accessibility check, notes, note
 * membership, used range), each checked before and after its read and
 * validated against the receipt. The same for every host. */
export function createOwnerReadControls(binding: OwnerReadBinding): {
  accessibilityControls: WorkbookAccessibilityControls;
  notesControls: WorkbookNotesControls;
  noteMembershipControls: WorkbookNoteMembershipControls;
  usedRangeControls: WorkbookUsedRangeControls;
} {
  const receipt = Object.freeze(noteMembershipReceipt({ ...binding.receipt }));
  const fields = (sheet?: string) => ({
    documentId: receipt.documentId,
    state: receipt.state,
    ...(sheet === undefined ? {} : { sheet }),
  });
  const read = async (operation: OwnerRead, sheet?: string) => {
    binding.current(operation);
    const result = await binding.request(operation, fields(sheet));
    binding.current(operation);
    return result;
  };
  return {
    accessibilityControls: {
      receipt,
      read: async () => validateAccessibility((await read('accessibility')).accessibility, receipt),
    },
    notesControls: {
      receipt,
      read: async (sheet) => validateNotes((await read('notes', sheet)).notes, receipt, { sheet }),
    },
    noteMembershipControls: {
      receipt,
      read: async (sheet) => {
        noteMembershipSheet(sheet);
        return validateNoteMembership((await read('noteMembership', sheet)).noteMembership, { sheet, receipt });
      },
    },
    usedRangeControls: {
      receipt,
      read: async (sheet) => {
        usedRangeSheet(sheet);
        return validateUsedRange((await read('usedRange', sheet)).usedRange, { sheet, receipt });
      },
    },
  };
}
