import type { WorkbookTool } from '../SpreadsheetRibbon';

/** First-use mounting preserves drafts, scroll positions and in-flight host jobs. */
export function RetainedWorkbookTool({
  tool,
  visible,
  children,
}: {
  tool: WorkbookTool;
  visible: boolean;
  children: import('react').ReactNode;
}): JSX.Element {
  return (
    <div
      className="sheet-tool-content"
      data-workbook-panel={tool}
      hidden={!visible}
      ref={(element) => {
        if (element) element.inert = !visible;
      }}
    >
      {children}
    </div>
  );
}
