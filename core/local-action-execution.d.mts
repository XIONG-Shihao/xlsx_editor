export type LocalExecutionAction = 'unhide-all' | 'freeze-selection';
export interface LocalExecutionPreview {
  action: LocalExecutionAction;
  label: string;
  sheet: string;
  selection: string;
  commands: Array<
    { operation: 'unhideAll'; sheet: string } | { operation: 'freeze'; sheet: string; rows: number; columns: number }
  >;
}
export const LOCAL_EXECUTION_ACTIONS: readonly LocalExecutionAction[];
export function localExecutionPreview(
  action: LocalExecutionAction,
  scope: { sheet: string; selection: string }
): LocalExecutionPreview;
