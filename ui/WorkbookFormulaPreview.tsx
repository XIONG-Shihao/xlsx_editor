import { useLayoutEffect, useSyncExternalStore } from 'react';
import {
  formulaPreviewInputKey,
  type FormulaDraftPreviewController,
  type FormulaPreviewConsumer,
  type FormulaPreviewInput,
} from './formulaDraftPreview';
export interface FormulaDraftPreviewProps {
  controller: FormulaDraftPreviewController;
  consumer: FormulaPreviewConsumer;
  input: FormulaPreviewInput | null;
  words?: (text: string) => string;
}
const identity = (text: string) => text;
export function FormulaDraftPreview({
  controller,
  consumer,
  input,
  words = identity,
}: FormulaDraftPreviewProps): JSX.Element | null {
  const pair = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot),
    key = input ? formulaPreviewInputKey(input) : null;
  useLayoutEffect(() => {
    controller.setInput(consumer, input);
    return () => controller.setInput(consumer, null);
  }, [controller, consumer, key]);
  if (
    !input?.active ||
    input.composing ||
    !input.scope.active ||
    !input.scope.normal ||
    input.scope.busy ||
    input.scope.profile !== 'xlsx'
  )
    return null;
  const snapshot = pair[consumer],
    current = snapshot.key === key ? snapshot : null,
    result = current?.reply?.result;
  let text = 'Preparing formula preview…',
    status: string = current?.phase ?? 'waiting';
  if (result?.status === 'value') {
    status = result.type;
    text =
      result.type === 'text'
        ? 'Text: ' + JSON.stringify(result.value)
        : result.type === 'boolean'
          ? 'Boolean: ' + (result.value ? 'TRUE' : 'FALSE')
          : result.type === 'error'
            ? 'Error: ' + String(result.value)
            : 'Number: ' + (Object.is(result.value, -0) ? '-0' : String(result.value));
  } else if (result) {
    status = result.status;
    text = result.reason;
  } else if (current?.reason) text = current.reason;
  else if (current?.phase === 'running') text = 'Calculating draft preview…';
  else if (current?.phase === 'idle') return null;
  return (
    <div
      className="sheet-formula-result-preview"
      data-formula-draft-zone={consumer === 'bar' ? '' : undefined}
      data-formula-preview-state={input.scope.state}
      data-formula-preview-target={input.scope.sheet + '!' + input.scope.cell}
      data-formula-preview-status={status}
    >
      <output aria-label={words('Formula result preview')} role="status">
        {words(text)}
      </output>
      {current && ['waiting', 'running', 'stopping'].includes(current.phase) && (
        <button type="button" onClick={() => controller.cancel(consumer, key!)}>
          {words('Cancel preview')}
        </button>
      )}
    </div>
  );
}
