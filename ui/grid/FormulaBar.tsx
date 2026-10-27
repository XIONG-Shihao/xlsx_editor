import { CellInput } from './CellInput';
import type { FormulaDraftHandle, FormulaDraftSnapshot } from '../namedFormulaDraft';
import { FormulaDraftPreview } from '../WorkbookFormulaPreview';
import type { FormulaDraftPreviewController, FormulaPreviewFrame } from '../formulaDraftPreview';
import type { RibbonCommandHandler, RibbonCommandId } from '../ribbonCommands';
import type { ViewPreferences } from '../viewPreferences';
import type {
  WorkbookCell,
  WorkbookCommand,
  WorkbookEdit,
  WorkbookSheet,
  WorkbookValidationRule,
} from '../../core/contracts';
import { inputDisplay } from '../sheetModel';

export interface FormulaBarInputs {
  busy: boolean;
  cancelCellFocus: () => void;
  canEditRef: React.MutableRefObject<boolean>;
  editRef: React.MutableRefObject<(edit: WorkbookEdit) => void>;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  formulaPreviewController: FormulaDraftPreviewController;
  formulaPreviewDraft: { scopeKey: string; snapshot: FormulaDraftSnapshot; active: boolean } | null;
  formulaPreviewFrame: React.MutableRefObject<FormulaPreviewFrame | null>;
  formulaPreviewKey: string | null;
  gridRoot: React.RefObject<HTMLDivElement>;
  invoke: (id: RibbonCommandId, handler: RibbonCommandHandler) => void;
  jump: string;
  jumpError: string;
  jumpToCell: () => void;
  literal: boolean | undefined;
  liveValidationScope: React.MutableRefObject<string>;
  namedDraftKey: string;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  openNamePicker: () => void;
  ownerIdentity: string | undefined;
  passiveNotice: string | undefined;
  range: string;
  readOnly: boolean;
  reportFormulaPreviewDraft: (scopeKey: string, snapshot: FormulaDraftSnapshot, active: boolean) => void;
  selected: string;
  selectedCell: WorkbookCell;
  selectedValidation: WorkbookValidationRule | undefined;
  setFormulaComposing: React.Dispatch<React.SetStateAction<boolean>>;
  setJump: React.Dispatch<React.SetStateAction<string>>;
  setValue: (cell: string, value: string) => void;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  validationScope: string;
  view: ViewPreferences;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  words: (en: string) => string;
}

export function FormulaBar({
  busy,
  cancelCellFocus,
  canEditRef,
  editRef,
  formulaDraftHandle,
  formulaPreviewController,
  formulaPreviewDraft,
  formulaPreviewFrame,
  formulaPreviewKey,
  gridRoot,
  invoke,
  jump,
  jumpError,
  jumpToCell,
  literal,
  liveValidationScope,
  namedDraftKey,
  onCommandBatch,
  openNamePicker,
  ownerIdentity,
  passiveNotice,
  range,
  readOnly,
  reportFormulaPreviewDraft,
  selected,
  selectedCell,
  selectedValidation,
  setFormulaComposing,
  setJump,
  setValue,
  sheet,
  specialSelection,
  validationScope,
  view,
  viewMode,
  words,
}: FormulaBarInputs): JSX.Element {
  return (
    <>
      <div className="sheet-formula-bar" hidden={!view.formulaBar || viewMode !== 'normal'}>
        <input
          className="sheet-cell-reference"
          aria-label={words('Go to cell')}
          value={jump}
          placeholder={specialSelection ? selected : range}
          onFocus={cancelCellFocus}
          onChange={(event) => setJump(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
              event.preventDefault();
              jumpToCell();
            }
          }}
        />
        <button onClick={jumpToCell}>{words('Go')}</button>
        {jumpError && <span role="alert">{jumpError}</span>}
        <span aria-hidden="true" className="sheet-fx">
          fx
        </span>
        <CellInput
          pending={busy}
          formulaBridge={
            !literal && !passiveNotice && onCommandBatch && ownerIdentity
              ? {
                  handle: formulaDraftHandle,
                  scopeKey: namedDraftKey,
                  onComposition: setFormulaComposing,
                  open: () => {
                    [...(gridRoot.current?.querySelectorAll<HTMLButtonElement>('.sheet-ribbon-tab') ?? [])]
                      .find((button) => button.textContent === words('Formulas'))
                      ?.click();
                    invoke('names.use-in-formula', openNamePicker);
                  },
                }
              : undefined
          }
          previewBridge={
            formulaPreviewKey && !literal && !passiveNotice
              ? { scopeKey: formulaPreviewKey, onDraft: reportFormulaPreviewDraft }
              : undefined
          }
          aria-keyshortcuts={!literal && !passiveNotice ? 'Alt+ArrowDown' : undefined}
          aria-label={literal ? words('Cell value') : words('Formula or value')}
          onFocus={cancelCellFocus}
          value={inputDisplay(selectedCell)}
          disabled={readOnly}
          onCommit={(value) => setValue(selected, value)}
        />
      </div>
      {formulaPreviewFrame.current && (
        <FormulaDraftPreview
          controller={formulaPreviewController}
          consumer="bar"
          input={
            formulaPreviewDraft?.scopeKey === formulaPreviewKey
              ? {
                  scope: formulaPreviewFrame.current.scope,
                  formula: formulaPreviewDraft.snapshot.text.startsWith('=')
                    ? formulaPreviewDraft.snapshot.text.slice(1)
                    : '',
                  generation: formulaPreviewDraft.snapshot.generation,
                  active:
                    formulaPreviewDraft.active && view.formulaBar && formulaPreviewDraft.snapshot.text.startsWith('='),
                  composing: formulaPreviewDraft.snapshot.composing,
                  invalidReason:
                    formulaPreviewDraft.snapshot.text && !formulaPreviewDraft.snapshot.text.startsWith('=')
                      ? 'Start a formula with = to preview its result.'
                      : undefined,
                }
              : null
          }
          words={words}
        />
      )}
      {selectedValidation && (
        <div className="sheet-validation-input">
          {selectedValidation.kind === 'list' && selectedValidation.showDropdown && (
            <label>
              {words('Validation choice')}{' '}
              <select
                aria-label={words('Validation choice')}
                disabled={readOnly || busy || !!specialSelection}
                value={
                  typeof selectedCell?.value === 'string' && selectedValidation.values?.includes(selectedCell.value)
                    ? selectedCell.value
                    : ''
                }
                onChange={(event) => {
                  if (
                    !canEditRef.current ||
                    liveValidationScope.current !== validationScope ||
                    specialSelection ||
                    !selectedValidation.values?.includes(event.target.value)
                  )
                    return;
                  editRef.current({ sheet: sheet.name, cell: selected, value: event.target.value });
                }}
              >
                <option value="" disabled>
                  {words('Choose a value…')}
                </option>
                {selectedValidation.values?.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
          )}
          {(selectedValidation.promptTitle || selectedValidation.prompt) && (
            <span role="status">
              {selectedValidation.promptTitle && <strong>{selectedValidation.promptTitle}: </strong>}
              {selectedValidation.prompt}
            </span>
          )}
        </div>
      )}
    </>
  );
}
