import {
  type FormulaDraftHandle,
  type NamedFormulaChoice,
  insertFormulaName,
  namedFormulaChoices,
} from '../namedFormulaDraft';
import type { GridNamePicker, GridSetNamePicker } from './gridTypes';
import type { ReadRange } from '../rangeCells';
import type { ViewPreferences } from '../viewPreferences';
import type { WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import { inRange } from './rangeAddress';
import { parseGridInput } from '../sheetModel';
import { targetProtection } from '../functionInsertion';

export interface NamePickerActionsInputs {
  activeFrame: (work: () => void) => number;
  cancelCellFocus: () => void;
  formulaComposing: boolean;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  namedDraftKey: string;
  namedDraftLive: React.MutableRefObject<{ key: string; editable: boolean }>;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  namePicker: GridNamePicker;
  namePickerSerial: React.MutableRefObject<number>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  readRange: ReadRange | undefined;
  selected: string;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setNamePicker: GridSetNamePicker;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  specialSelection: Set<string> | null;
  view: ViewPreferences;
}

export function buildNamePickerActions({
  activeFrame,
  cancelCellFocus,
  formulaComposing,
  formulaDraftHandle,
  namedDraftKey,
  namedDraftLive,
  namedRanges,
  namePicker,
  namePickerSerial,
  onCommandBatch,
  ownerIdentity,
  readRange,
  selected,
  setGridError,
  setNamePicker,
  sheet,
  sheets,
  specialSelection,
  view,
}: NamePickerActionsInputs) {
  const currentNamePicker =
    namePicker && namePicker.key === namedDraftKey && namedDraftLive.current.editable && !formulaComposing
      ? namePicker
      : null;
  const closeNamePicker = (): void => {
    const picker = currentNamePicker,
      serial = ++namePickerSerial.current;
    setNamePicker(null);
    if (picker)
      activeFrame(() => {
        if (namePickerSerial.current === serial && namedDraftLive.current.key === picker.key) picker.handle.focus();
      });
  };
  const nameDraftReason =
    !onCommandBatch || !ownerIdentity
      ? 'Name insertion is unavailable in this editor.'
      : !view.formulaBar
        ? 'Show the formula bar to insert a name.'
        : formulaComposing
          ? 'Finish composing the formula first.'
          : specialSelection
            ? 'Click a regular cell to insert a name.'
            : undefined;
  const openNamePicker = (): void => {
    if (nameDraftReason || !namedDraftLive.current.editable)
      throw Error(nameDraftReason ?? 'This command requires an editable XLSX workbook.');
    const handle = formulaDraftHandle.current;
    if (!handle) throw Error('The formula editor is unavailable.');
    const field = document.activeElement as HTMLElement | null;
    if (
      field?.matches('input,textarea,select,[contenteditable="true"]') &&
      !handle.ownsFocus() &&
      !field.closest('[data-formula-draft-zone]')
    )
      throw Error('Finish or cancel the current input before inserting a name.');
    const snapshot = handle.capture();
    if (snapshot.composing) throw Error('Finish composing the formula first.');
    if (snapshot.text !== '' && !snapshot.text.startsWith('='))
      throw Error('Start a formula with = before inserting a name.');
    const key = namedDraftKey,
      target = selected,
      ownerSheet = sheet,
      reader = readRange,
      accept = onCommandBatch!;
    const isCurrent = () => namedDraftLive.current.key === key && namedDraftLive.current.editable;
    const choices = namedFormulaChoices(
      namedRanges,
      sheet.name,
      sheets.map((item) => item.name)
    );
    handle.park();
    handle.attach({
      isCurrent,
      accept: async (text) => {
        if (!isCurrent()) throw Error('The formula draft target changed.');
        const cells = reader
          ? await reader(ownerSheet.name, target + ':' + target)
          : { [target]: ownerSheet.cells[target] };
        if (!isCurrent()) throw Error('The formula draft target changed.');
        if (
          !cells ||
          typeof cells !== 'object' ||
          Array.isArray(cells) ||
          Object.keys(cells).some((ref) => ref !== target)
        )
          throw Error('The target reply did not match the formula cell.');
        const reason = targetProtection({
          cell: cells[target],
          ...(ownerSheet.merges.some((range) => inRange(range, target))
            ? { protectedReason: 'Choose an unmerged cell for name insertion.' }
            : {}),
        });
        if (reason) throw Error(reason);
        if (cells[target]?.format === 49 || cells[target]?.format === '@')
          throw Error('Choose a cell with a formula-compatible number format.');
        await accept([{ sheet: ownerSheet.name, cell: target, ...parseGridInput(text, false) }]);
      },
      error: (error) => {
        if (namedDraftLive.current.key === key) setGridError(String(error instanceof Error ? error.message : error));
      },
    });
    namePickerSerial.current++;
    cancelCellFocus();
    setGridError('');
    setNamePicker({ key, target: sheet.name + '!' + target, snapshot, choices, handle });
  };
  const chooseFormulaName = (choice: NamedFormulaChoice): void => {
    const picker = currentNamePicker;
    if (!picker || namedDraftLive.current.key !== picker.key || !namedDraftLive.current.editable) return;
    try {
      if (!picker.choices.some((item) => item.id === choice.id && !item.disabledReason))
        throw Error('The selected name changed. Reopen the picker.');
      const next = insertFormulaName(picker.snapshot, choice);
      picker.handle.replace(picker.snapshot, next.text, next.caret);
      setGridError('');
      closeNamePicker();
    } catch (error) {
      if (namedDraftLive.current.key === picker.key)
        setGridError(String(error instanceof Error ? error.message : error));
    }
  };
  return { currentNamePicker, closeNamePicker, nameDraftReason, openNamePicker, chooseFormulaName } as const;
}
