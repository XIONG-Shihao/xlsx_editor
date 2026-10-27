import { WorkbookInkPalette } from './WorkbookInkPalette';
import type { InkPenId } from '../core/ink-stroke.mjs';
import { WorkbookViewPresets, type ViewPresetControls } from './WorkbookViewPresets';
import { createContext, useContext, useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { Icon } from './host';
import { useDocumentWords } from './words';
import {
  dispatchRibbonCommand,
  ribbonCapability,
  ribbonCommand,
  ribbonCommandTools,
  ribbonTabOrder,
  ribbonToolCommands,
  type WorkbookTool,
  type RibbonTab,
  type RibbonCommandId,
  type RibbonCommandHandler,
  type RibbonDispatchScope,
  type RibbonDispatchResult,
  type RibbonArguments,
  type RibbonBinding,
} from './ribbonCommands';

const RibbonContext = createContext({ width: 1400, openGroup: '', setOpenGroup: (_name: string) => {} });

export type { WorkbookTool } from './ribbonCommands';
export interface WorkbookRibbonTool {
  id: WorkbookTool;
  label: string;
  disabled?: boolean;
}
export interface RibbonCommandControl {
  id: RibbonCommandId;
  handler: RibbonCommandHandler;
  label?: string;
  disabled?: boolean;
  reason?: string;
  hint?: string;
  pressed?: boolean | null;
  popover?: ReactNode;
}
export interface RibbonViewControls {
  presets?: ViewPresetControls;
  zoom: number;
  gridlines: boolean;
  headings: boolean;
  formulaBar: boolean;
  showFormulas: boolean;
  onChange: (
    patch: Partial<Pick<RibbonViewControls, 'zoom' | 'gridlines' | 'headings' | 'formulaBar' | 'showFormulas'>>
  ) => void | Promise<unknown>;
  onZoomToSelection: () => void | Promise<unknown>;
  onNormal?: () => void | Promise<unknown>;
  mode?: 'normal' | 'pageLayout' | 'pageBreak';
  onFreeze: (mode: 'selection' | 'row' | 'column' | 'none') => void | Promise<unknown>;
  canFreeze: boolean;
  freezeReason?: string;
}
const toolIcons: Record<WorkbookTool, ReactNode> = {
  accessibility: (
    <>
      <circle cx="12" cy="4" r="2" />
      <path d="M3 9h18m-9 0v6m0 0-5 7m5-7 5 7" />
    </>
  ),
  notes: (
    <>
      <path d="M4 3h16v18H4zM8 8h8M8 12h8M8 16h5" />
    </>
  ),
  illustrations: (
    <>
      <rect x="3" y="3" width="11" height="11" />
      <ellipse cx="16" cy="16" rx="6" ry="5" />
    </>
  ),
  pageStyles: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 9h18M9 9v12m4-8h5m-5 4h5" />
    </>
  ),
  checkbox: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="m7 12 3 3 7-7" />
    </>
  ),
  symbols: (
    <>
      <path d="M4 20h6v-3a8 8 0 1 1 4 0v3h6" />
    </>
  ),
  hyperlink: (
    <>
      <path d="m9 15 6-6m-4-4 2-2a5 5 0 0 1 7 7l-4 4a5 5 0 0 1-7 0m4 5-2 2a5 5 0 0 1-7-7l4-4a5 5 0 0 1 7 0" />
    </>
  ),
  pivotDesign: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 8h18M3 14h18M8 3v18" />
    </>
  ),
  pivot: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 8h18M8 8v13M3 14h18m9-3 6 3-6 3" />
    </>
  ),
  special: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 9h18M9 3v18m4-7 2 2 4-4" />
    </>
  ),
  replace: (
    <>
      <circle cx="9" cy="9" r="6" />
      <path d="m14 14 7 7M5 9h8m-3-3 3 3-3 3" />
    </>
  ),
  fill: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 9h18M9 3v18m4-7 3 3 3-3m-3-3v6" />
    </>
  ),
  actions: (
    <>
      <path d="M5 3h10l4 4v14H5V3Zm10 0v4h4M8 11h8M8 15h8M8 18h4" />
    </>
  ),
  review: (
    <>
      <path d="M4 3h12l4 4v14H4V3Zm12 0v4h4M7 11h10M7 15h6" />
      <path d="m13 18 2 2 4-4" />
    </>
  ),
  duplicates: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 8h18M8 3v18m4-9 6 6m0-6-6 6" />
    </>
  ),
  sort: (
    <>
      <path d="M4 5h11M4 11h8M4 17h5m9-12v14m-3-3 3 3 3-3" />
    </>
  ),
  calculation: (
    <>
      <rect x="3" y="2" width="18" height="20" rx="2" />
      <path d="M6 6h12M6 11h2m4 0h2m4 0h1M6 16h2m4 0h2m4 0h1" />
    </>
  ),
  trace: (
    <>
      <rect x="2" y="3" width="6" height="6" />
      <rect x="16" y="15" width="6" height="6" />
      <path d="M8 6h5v12h3m-3-3 3 3-3 3" />
    </>
  ),
  textColumns: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M9 3v18m6-18v18M3 9h18" />
    </>
  ),
  errors: (
    <>
      <path d="m12 3 10 18H2L12 3Z" />
      <path d="M12 9v5m0 3v1" />
    </>
  ),
  watch: (
    <>
      <path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12Z" />
      <circle cx="12" cy="12" r="3" />
    </>
  ),
  functions: (
    <>
      <path d="M14 4c-2-2-5-1-5 2L7 18c-.5 3-3 3-4 2M5 10h8m2 3 6 6m0-6-6 6" />
    </>
  ),
  createNames: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 8h18M8 8v13m3-7h7m-3-3v6" />
    </>
  ),
  names: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="M3 8h18M9 8v13m4-5 2-5 2 5m-3-2h2" />
    </>
  ),
  objects: (
    <>
      <path d="M3 3v18h18" />
      <rect x="6" y="12" width="3" height="6" />
      <rect x="11" y="8" width="3" height="10" />
      <rect x="16" y="4" width="3" height="14" />
    </>
  ),
  print: (
    <>
      <path d="M6 9V3h12v6M6 17H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2" />
      <rect x="6" y="14" width="12" height="7" />
      <path d="M17 12h2" />
    </>
  ),
  conditional: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="1" />
      <path d="M3 9h18M9 3v18M3 15h6" />
      <path d="m12 15 2 2 4-5" />
    </>
  ),
  validation: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <path d="m6 8 2 2 3-4m2 2h5M6 15h5m3 0 2 2 3-3" />
    </>
  ),
  references: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="1" />
      <path d="M3 8h18M9 8v13M15 8v13M3 14h18" />
    </>
  ),
  queries: (
    <>
      <ellipse cx="12" cy="5" rx="8" ry="3" />
      <path d="M4 5v12c0 4 16 4 16 0V5M4 11c0 4 16 4 16 0" />
      <path d="m9 15 3 3 3-3m-3-4v7" />
    </>
  ),
};
const functionIcons: Partial<Record<RibbonCommandId, ReactNode>> = {
  'draw.pen': (
    <>
      <path d="m4 17 1-5L17 1l5 5L10 18l-5 1m0-7 5 6M3 21c4-4 7 2 11-1" />
    </>
  ),
  'draw.select-strokes': (
    <>
      <path d="M3 3h7M3 3v7M21 3h-7M21 3v7M3 21h7M3 21v-7m18 7h-7m7 0v-7" />
    </>
  ),
  'draw.eraser': (
    <>
      <path d="m3 15 11-12 8 8-9 10H8l-5-6Zm7-7 8 8M3 21h18" />
    </>
  ),
  'draw.select-cells': (
    <>
      <path d="m5 3 14 10-7 1-3 7L5 3Z" />
    </>
  ),
  'review.accessibility': toolIcons.accessibility,
  'data.refresh-all': (
    <>
      <path d="M19 8a8 8 0 0 0-13-3L3 8m0-5v5h5m-3 8a8 8 0 0 0 13 3l3-3m0 5v-5h-5" />
    </>
  ),
  'actions.unhide-all': (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 9h18M9 3v18m4-7 3-3 3 3m-3-3v8" />
    </>
  ),
  'actions.freeze-selection': (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 9h18M9 3v18m4-5h6" />
    </>
  ),
  'actions.new': (
    <>
      <path d="M5 3h10l4 4v14H5V3Zm10 0v4h4M8 11h8M8 15h4M16 15v6M13 18h6" />
    </>
  ),
  'actions.library': (
    <>
      <path d="M4 3h16v18H4V3ZM8 3v18M11 8h6M11 12h6M11 16h6" />
    </>
  ),
  'review.statistics': (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 8h18M8 8v13M12 12h5M12 16h5" />
    </>
  ),
  'review.spelling': (
    <>
      <path d="m3 12 4-9 4 9M5 8h4M13 4h4a3 3 0 0 1 0 6h-4V4Zm0 0v8m0 5 3 3 5-6" />
    </>
  ),
  'formulas.calculate-sheet': (
    <>
      <rect x="3" y="3" width="14" height="18" />
      <path d="M3 8h14M8 8v13m11-10-4 6h4l-2 5" />
    </>
  ),
  'formulas.calculate-now': (
    <>
      <rect x="3" y="2" width="14" height="20" rx="2" />
      <path d="M6 6h8M6 10h2m3 0h2M6 14h2m-2 4h2m10-7-4 6h4l-2 5" />
    </>
  ),
  'formulas.trace-precedents': (
    <>
      <rect x="2" y="3" width="6" height="6" />
      <path d="M8 6h5v12h7m-3-3 3 3-3 3" />
    </>
  ),
  'formulas.trace-dependents': (
    <>
      <rect x="16" y="15" width="6" height="6" />
      <path d="M16 18h-5V6H4m3-3L4 6l3 3" />
    </>
  ),
  'formulas.remove-arrows': (
    <>
      <path d="M3 6h10v12h8m-3-3 3 3-3 3M4 13l6 6m0-6-6 6" />
    </>
  ),
  'formulas.auto-sum': <path d="M18 4H6l7 8-7 8h12" />,
  'formulas.recently-used': <path d="m12 3 3 6 7 1-5 5 1 7-6-3-6 3 1-7-5-5 7-1 3-6Z" />,
  'formulas.financial': (
    <>
      <ellipse cx="12" cy="5" rx="7" ry="3" />
      <path d="M5 5v5c0 4 14 4 14 0V5M5 10v5c0 4 14 4 14 0v-5M5 15v4c0 4 14 4 14 0v-4" />
    </>
  ),
  'formulas.logical': (
    <>
      <path d="M8 8a4 4 0 1 1 6 3c-2 1-2 2-2 4" />
      <circle cx="12" cy="19" r=".5" />
    </>
  ),
  'formulas.text': (
    <>
      <path d="m5 20 7-16 7 16M8 14h8" />
    </>
  ),
  'formulas.date-time': (
    <>
      <rect x="3" y="4" width="18" height="17" rx="2" />
      <path d="M3 9h18M8 2v4m8-4v4M12 12v4l3 2" />
    </>
  ),
  'formulas.lookup-reference': (
    <>
      <circle cx="10" cy="10" r="6" />
      <path d="m15 15 6 6M7 10h6m-3-3v6" />
    </>
  ),
  'formulas.math-trig': (
    <>
      <ellipse cx="12" cy="12" rx="6" ry="9" />
      <path d="M6 12h12" />
    </>
  ),
  'formulas.more-functions': (
    <>
      <circle cx="5" cy="12" r="1" />
      <circle cx="12" cy="12" r="1" />
      <circle cx="19" cy="12" r="1" />
    </>
  ),
};

/** Keep familiar groups in place; shrink whole groups as the document pane narrows. */
export function SpreadsheetRibbon({
  children,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  busy,
  tools = [],
  activeTool,
  onToolToggle,
  taskPaneId,
  dataCommands,
  formulaCommands,
  viewCommands,
  commands = [],
  commandScope,
  onCommandResult,
  onTabChange,
  ink,
  view,
}: {
  children: ReactNode;
  canUndo: boolean;
  canRedo: boolean;
  onUndo?: () => void;
  onRedo?: () => void;
  busy?: boolean;
  tools?: WorkbookRibbonTool[];
  activeTool?: WorkbookTool | null;
  onToolToggle?: (id: WorkbookTool, binding?: RibbonBinding) => void;
  taskPaneId?: string;
  dataCommands?: ReactNode;
  formulaCommands?: ReactNode;
  viewCommands?: ReactNode;
  commands?: RibbonCommandControl[];
  commandScope?: RibbonDispatchScope;
  onCommandResult?: (id: RibbonCommandId, result: RibbonDispatchResult<unknown>) => void;
  view?: RibbonViewControls;
  ink?: {
    selected: InkPenId;
    disabled: boolean;
    reason?: string;
    warning?: string | null;
    onChoose: RibbonCommandHandler;
  };
  onTabChange?: (tab: RibbonTab) => void;
}): JSX.Element {
  const words = useDocumentWords();
  const ribbonId = useId();
  const root = useRef<HTMLDivElement>(null);
  const tabButtons = useRef<Partial<Record<RibbonTab, HTMLButtonElement | null>>>({});
  const [width, setWidth] = useState(1400);
  const [openGroup, setOpenGroup] = useState('');
  const previousNormalTab = useRef<RibbonTab>('Home');
  const [selectedTab, setSelectedTab] = useState<RibbonTab>('Home');
  const hasDataCommands = dataCommands !== undefined && dataCommands !== null && dataCommands !== false;
  const capability = (id: RibbonCommandId, hasHandler: boolean, disabled = false, reason?: string) =>
    commandScope
      ? ribbonCapability(id, {
          ...commandScope.availability(id),
          hasHandler,
          ...(disabled
            ? {
                capability: {
                  state: 'disabled' as const,
                  reason: reason ?? 'This command is unavailable for this selection.',
                },
              }
            : {}),
        })
      : { state: 'disabled' as const, reason: 'This host has not bound this command to a workbook.' };
  const directCommands = commands
    .filter((command) => capability(command.id, !!command.handler, command.disabled, command.reason).state !== 'hidden')
    .sort((a, b) => ribbonCommand(a.id).order - ribbonCommand(b.id).order);
  const availableTools = tools.filter(
    (tool) => !commandScope || capability(ribbonToolCommands[tool.id], !!onToolToggle, tool.disabled).state !== 'hidden'
  );
  const availableTabs = ribbonTabOrder.filter(
    (tab) =>
      tab === 'Home' ||
      (tab === 'Data' && hasDataCommands) ||
      (tab === 'Formulas' && (!!formulaCommands || !!view)) ||
      (tab === 'View' && (!!view || !!viewCommands)) ||
      availableTools.some((tool) => ribbonCommand(ribbonToolCommands[tool.id]).tab === tab) ||
      directCommands.some((command) => ribbonCommand(command.id).tab === tab)
  );
  const tab = availableTabs.includes(selectedTab)
    ? selectedTab
    : availableTabs.includes(previousNormalTab.current)
      ? previousNormalTab.current
      : 'Home';
  const closeMenus = (): void => {
    setOpenGroup('');
    root.current?.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((menu) => {
      menu.open = false;
    });
  };
  const selectTab = (next: RibbonTab): void => {
    closeMenus();
    if (next !== 'PivotTable Analyze' && next !== 'Design') previousNormalTab.current = next;
    setSelectedTab(next);
    onTabChange?.(next);
  };
  useEffect(() => {
    if (!availableTabs.includes(selectedTab))
      setSelectedTab(availableTabs.includes(previousNormalTab.current) ? previousNormalTab.current : 'Home');
  }, [selectedTab, availableTabs.join('|')]);
  useEffect(() => {
    closeMenus();
  }, [tab]);
  const execute = async (id: RibbonCommandId, handler: RibbonCommandHandler, args?: RibbonArguments): Promise<void> => {
    if (!commandScope) return;
    try {
      const result = await dispatchRibbonCommand(id, commandScope, handler, args);
      onCommandResult?.(id, result);
    } catch (error) {
      onCommandResult?.(id, { outcome: 'refused', reason: error instanceof Error ? error.message : String(error) });
    }
  };
  const toolButton = (tool: WorkbookRibbonTool): JSX.Element => {
    // Opening or closing a pane does not mutate the workbook. Keep running-job
    // controls reachable; each pane fences its own edits and publication.
    const definition = ribbonCommand(ribbonToolCommands[tool.id]);
    const available = commandScope ? capability(definition.id, !!onToolToggle, tool.disabled) : undefined;
    const disabled = available ? available.state !== 'enabled' : !!tool.disabled || !onToolToggle;
    const unavailable = !onToolToggle ? 'Unavailable in this host.' : 'Unavailable for this workbook or selection.';
    return (
      <button
        type="button"
        data-workbook-tool={tool.id}
        data-ribbon-command={definition.id}
        className={`sheet-ribbon-tool${activeTool === tool.id ? ' is-active' : ''}`}
        disabled={disabled}
        aria-label={words(tool.label)}
        aria-pressed={activeTool === tool.id}
        aria-controls={taskPaneId}
        title={`${words(tool.label)}. ${words(disabled ? (available?.reason ?? unavailable) : definition.hint)}`}
        onClick={() => {
          if (commandScope) void execute(definition.id, (binding) => onToolToggle?.(tool.id, binding));
          else onToolToggle?.(tool.id);
        }}
      >
        <svg
          className="sheet-ribbon-tool-icon"
          width={25}
          height={25}
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth={1.5}
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          {toolIcons[tool.id]}
        </svg>
        <span className="sheet-ribbon-tool-label">{words(tool.label)}</span>
      </button>
    );
  };
  const groupsFor = (panelTab: RibbonTab): JSX.Element[] => {
    const groups = new Map<string, { order: number; nodes: Array<{ order: number; node: ReactNode }> }>();
    const add = (id: RibbonCommandId, node: ReactNode): void => {
      const definition = ribbonCommand(id);
      if (definition.tab !== panelTab) return;
      const group = groups.get(definition.group) ?? { order: definition.order, nodes: [] };
      group.order = Math.min(group.order, definition.order);
      group.nodes.push({ order: definition.order, node });
      groups.set(definition.group, group);
    };
    if (panelTab === 'Formulas' && view) add('view.show-formulas', viewToggle('showFormulas', 'view.show-formulas'));
    if (panelTab === 'View' && view?.onNormal)
      add('view.normal', viewButton('view.normal', view.onNormal, (view.mode ?? 'normal') === 'normal'));
    for (const tool of availableTools)
      if (!directCommands.some((command) => command.id === ribbonToolCommands[tool.id]))
        add(
          ribbonToolCommands[tool.id],
          <span key={`tool-${tool.id}`} data-ribbon-item={ribbonToolCommands[tool.id]}>
            {toolButton(tool)}
          </span>
        );
    for (const command of directCommands) {
      const definition = ribbonCommand(command.id),
        available = capability(command.id, !!command.handler, command.disabled, command.reason),
        pane = ribbonCommandTools[command.id];
      const icon = functionIcons[command.id];
      const button = (
        <button
          key={command.id}
          type="button"
          data-ribbon-item={command.id}
          data-ribbon-command={command.id}
          data-formula-draft-zone={
            command.id === 'names.use-in-formula' ||
            command.id.startsWith('review.') ||
            command.id.startsWith('actions.')
              ? ''
              : undefined
          }
          className={`sheet-ribbon-command${icon ? ' sheet-ribbon-tool' : ''}`}
          aria-haspopup={command.id === 'names.use-in-formula' ? 'dialog' : undefined}
          aria-expanded={command.id === 'names.use-in-formula' ? !!command.popover : undefined}
          disabled={available.state !== 'enabled'}
          aria-pressed={
            command.pressed === null ? undefined : (command.pressed ?? (pane ? activeTool === pane : undefined))
          }
          aria-controls={pane ? taskPaneId : undefined}
          title={words(available.reason ?? command.hint ?? definition.hint)}
          onClick={() => void execute(command.id, command.handler)}
        >
          {icon && (
            <svg
              className="sheet-ribbon-tool-icon"
              width={25}
              height={25}
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth={1.5}
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              {icon}
            </svg>
          )}
          <span className={icon ? 'sheet-ribbon-tool-label' : undefined}>
            {words(command.label ?? definition.label)}
          </span>
        </button>
      );
      add(
        command.id,
        command.id === 'names.use-in-formula' ? (
          <span key={command.id} data-ribbon-item={command.id} className="sheet-name-picker-anchor">
            {button}
            {command.popover}
          </span>
        ) : (
          button
        )
      );
    }
    return [...groups]
      .sort((a, b) => a[1].order - b[1].order)
      .map(([group, contents]) => (
        <RibbonGroup key={group} name={group} icon="table" collapseBelow={0}>
          <div className="sheet-ribbon-body sheet-ribbon-tools">
            {contents.nodes.sort((a, b) => a.order - b.order).map((item) => item.node)}
          </div>
          <small>{words(group)}</small>
        </RibbonGroup>
      ));
  };
  const viewAvailability = (id: RibbonCommandId) =>
    capability(id, !!view, id === 'view.freeze' && !view?.canFreeze, view?.freezeReason);
  const viewToggle = (
    field: 'gridlines' | 'headings' | 'formulaBar' | 'showFormulas',
    id: RibbonCommandId
  ): JSX.Element => {
    const available = viewAvailability(id);
    return (
      <label key={field} data-ribbon-item={id} title={words(available.reason ?? ribbonCommand(id).hint)}>
        <input
          type="checkbox"
          data-ribbon-command={id}
          aria-label={words(ribbonCommand(id).label)}
          checked={!!view?.[field]}
          disabled={available.state !== 'enabled'}
          onChange={(event) => {
            const value = event.target.checked;
            void execute(id, (_binding, args) => view!.onChange({ [field]: (args as { value: boolean }).value }), {
              value,
            });
          }}
        />
        {field === 'showFormulas' && (
          <svg
            className="sheet-ribbon-formula-toggle-icon"
            width={22}
            height={22}
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth={1.5}
            aria-hidden="true"
          >
            {toolIcons.functions}
          </svg>
        )}
        <span>{words(ribbonCommand(id).label)}</span>
      </label>
    );
  };
  const viewButton = (id: RibbonCommandId, handler: () => void | Promise<unknown>, pressed?: boolean): JSX.Element => {
    const available = viewAvailability(id);
    return (
      <button
        key={id}
        type="button"
        data-ribbon-command={id}
        disabled={available.state !== 'enabled'}
        aria-pressed={pressed}
        title={words(available.reason ?? ribbonCommand(id).hint)}
        onClick={() => void execute(id, handler)}
      >
        {words(ribbonCommand(id).label)}
      </button>
    );
  };
  const viewToolbar = view && (
    <>
      {view.presets && (
        <RibbonGroup name="Local views" icon="table" collapseBelow={0}>
          <div className="sheet-ribbon-body">
            <WorkbookViewPresets controls={view.presets} words={words} run={(id, handler) => execute(id, handler)} />
          </div>
          <small>{words('Local views')}</small>
        </RibbonGroup>
      )}
      <RibbonGroup name="Show" icon="table" collapseBelow={0}>
        <div className="sheet-ribbon-body sheet-ribbon-stack">
          {viewToggle('gridlines', 'view.gridlines')}
          {viewToggle('headings', 'view.headings')}
          {viewToggle('formulaBar', 'view.formula-bar')}
          {viewToggle('showFormulas', 'view.show-formulas')}
        </div>
        <small>{words('Show')}</small>
      </RibbonGroup>
      <RibbonGroup name="Zoom" icon="search" collapseBelow={0}>
        <div className="sheet-ribbon-body sheet-ribbon-stack">
          <RibbonZoomInput
            zoom={view.zoom}
            disabled={viewAvailability('view.zoom').state !== 'enabled'}
            reason={viewAvailability('view.zoom').reason}
            onCommit={(zoom) =>
              execute('view.zoom', (_binding, args) => view.onChange({ zoom: (args as { zoom: number }).zoom }), {
                zoom,
              })
            }
          />
          {viewButton('view.zoom-100', () => view.onChange({ zoom: 100 }))}
          {viewButton('view.zoom-selection', view.onZoomToSelection)}
        </div>
        <small>{words('Zoom')}</small>
      </RibbonGroup>
      {viewAvailability('view.freeze').state !== 'hidden' && (
        <RibbonGroup name="Window" icon="table" collapseBelow={0}>
          <div className="sheet-ribbon-body">
            <details className="sheet-ribbon-menu">
              <summary
                title={words(viewAvailability('view.freeze').reason ?? 'Freeze rows or columns while scrolling.')}
              >
                {words('Freeze Panes')} <Icon name="chevron" size={11} />
              </summary>
              <div>
                {(
                  [
                    ['selection', 'Freeze here'],
                    ['row', 'Freeze first row'],
                    ['column', 'Freeze first column'],
                    ['none', 'Unfreeze'],
                  ] as const
                ).map(([mode, label]) => (
                  <button
                    key={mode}
                    type="button"
                    data-ribbon-command="view.freeze"
                    data-freeze-mode={mode}
                    disabled={viewAvailability('view.freeze').state !== 'enabled'}
                    title={words(viewAvailability('view.freeze').reason ?? label)}
                    onClick={() =>
                      void execute(
                        'view.freeze',
                        (_binding, args) => view.onFreeze((args as { mode: typeof mode }).mode),
                        { mode }
                      )
                    }
                  >
                    {words(label)}
                  </button>
                ))}
              </div>
            </details>
          </div>
          <small>{words('Window')}</small>
        </RibbonGroup>
      )}
    </>
  );
  useEffect(() => {
    root.current?.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((menu) => {
      menu.open = false;
    });
  }, [openGroup]);
  useEffect(() => {
    const element = root.current!;
    const observer = new ResizeObserver(([entry]) => {
      setWidth(entry.contentRect.width);
      setOpenGroup('');
      element.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((menu) => {
        menu.open = false;
      });
    });
    observer.observe(element);
    const dismiss = (event: Event): void => {
      if (element.contains(event.target as Node)) return;
      setOpenGroup('');
      element.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((menu) => {
        menu.open = false;
      });
    };
    let openMenu: { details: HTMLDetailsElement; panel: HTMLElement } | null = null;
    let menuObserver: ResizeObserver | null = null;
    let menuFrame: number | null = null;
    let menuScheduled = false;
    const stopPositioning = (): void => {
      menuObserver?.disconnect();
      menuObserver = null;
      if (menuFrame !== null) cancelAnimationFrame(menuFrame);
      menuFrame = null;
      menuScheduled = false;
      openMenu?.panel.style.removeProperty('max-height');
      openMenu = null;
    };
    const repositionMenu = (): void => {
      if (!openMenu || !openMenu.details.open || !element.contains(openMenu.details)) return;
      const { details, panel } = openMenu;
      const margin = 8,
        gap = 4,
        height = window.innerHeight;
      const anchor = details.getBoundingClientRect(),
        box = panel.getBoundingClientRect();
      // scrollHeight retains the unconstrained content extent once the menu
      // scrolls; use border-box height for short menus without overflow.
      const contentHeight = Math.max(box.height, panel.scrollHeight + Math.max(0, box.height - panel.clientHeight));
      const belowTop = Math.max(margin, Math.min(anchor.bottom + gap, height - margin));
      const aboveBottom = Math.max(margin, Math.min(anchor.top - gap, height - margin));
      const below = Math.max(0, height - margin - belowTop),
        above = Math.max(0, aboveBottom - margin);
      const placeAbove = contentHeight > below && (contentHeight <= above || above > below);
      const availableHeight = placeAbove ? above : below;
      panel.style.maxHeight = `${availableHeight}px`;
      panel.style.left = `${Math.max(margin, Math.min(anchor.left, window.innerWidth - box.width - margin))}px`;
      panel.style.top = `${placeAbove ? aboveBottom - Math.min(contentHeight, availableHeight) : belowTop}px`;
    };
    const schedulePosition = (): void => {
      if (!openMenu || !openMenu.details.open || menuScheduled) return;
      menuScheduled = true;
      const frame = requestAnimationFrame(() => {
        menuScheduled = false;
        menuFrame = null;
        repositionMenu();
      });
      // Some embedded test hosts execute RAF synchronously.
      if (menuScheduled) menuFrame = frame;
    };
    const positionMenu = (event: Event): void => {
      const details = event.target;
      if (!(details instanceof HTMLDetailsElement)) return;
      if (!details.open) {
        if (openMenu?.details === details) stopPositioning();
        return;
      }
      element.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((other) => {
        if (other !== details) other.open = false;
      });
      const panel = details.querySelector<HTMLElement>(':scope > div');
      if (!panel) return;
      if (openMenu?.details !== details) {
        stopPositioning();
        openMenu = { details, panel };
        menuObserver = new ResizeObserver(schedulePosition);
        menuObserver.observe(panel);
      }
      schedulePosition();
    };
    element.addEventListener('toggle', positionMenu, true);
    window.addEventListener('resize', schedulePosition);
    document.addEventListener('pointerdown', dismiss);
    document.addEventListener('focusin', dismiss);
    return () => {
      observer.disconnect();
      stopPositioning();
      element.removeEventListener('toggle', positionMenu, true);
      window.removeEventListener('resize', schedulePosition);
      document.removeEventListener('pointerdown', dismiss);
      document.removeEventListener('focusin', dismiss);
    };
  }, []);
  return (
    <RibbonContext.Provider value={{ width, openGroup, setOpenGroup }}>
      <div
        className="sheet-ribbon-shell"
        ref={root}
        onPointerDownCapture={(event) => {
          if ((event.target as HTMLElement).closest('[data-ribbon-command="names.use-in-formula"]'))
            event.preventDefault();
        }}
        onKeyDown={(event) => {
          if (event.key !== 'Escape') return;
          const menu = root.current?.querySelector<HTMLDetailsElement>('details[open]');
          if (menu) {
            menu.open = false;
            menu.querySelector('summary')?.focus();
          } else if (openGroup) {
            root.current?.querySelector<HTMLButtonElement>('.sheet-ribbon-group-toggle[aria-expanded="true"]')?.focus();
            setOpenGroup('');
          } else return;
          event.preventDefault();
          event.stopPropagation();
        }}
      >
        <div className="sheet-quick-access">
          <div className="sheet-history-tools" role="group" aria-label={words('Undo and redo')}>
            <button
              data-home-command="home.undo"
              data-formula-draft-zone
              aria-label={words('Undo')}
              title={words('Undo')}
              disabled={!canUndo}
              onClick={onUndo}
            >
              <Icon name="undo" size={16} />
            </button>
            <button
              data-home-command="home.redo"
              data-formula-draft-zone
              className="sheet-redo"
              aria-label={words('Redo')}
              title={words('Redo')}
              disabled={!canRedo}
              onClick={onRedo}
            >
              <Icon name="undo" size={16} />
            </button>
          </div>
          <div className="sheet-ribbon-tab-list" role="tablist" aria-label={words('Spreadsheet ribbon')}>
            {availableTabs.map((item, index) => (
              <button
                key={item}
                ref={(element) => {
                  tabButtons.current[item] = element;
                }}
                type="button"
                id={`${ribbonId}-tab-${index}`}
                className={`sheet-ribbon-tab${tab === item ? ' is-active' : ''}`}
                role="tab"
                data-formula-draft-zone
                aria-selected={tab === item}
                aria-controls={`${ribbonId}-panel-${index}`}
                tabIndex={tab === item ? 0 : -1}
                onClick={() => selectTab(item)}
                onKeyDown={(event) => {
                  let next: number;
                  if (event.key === 'ArrowRight') next = (index + 1) % availableTabs.length;
                  else if (event.key === 'ArrowLeft') next = (index - 1 + availableTabs.length) % availableTabs.length;
                  else if (event.key === 'Home') next = 0;
                  else if (event.key === 'End') next = availableTabs.length - 1;
                  else return;
                  event.preventDefault();
                  event.stopPropagation();
                  const nextTab = availableTabs[next];
                  selectTab(nextTab);
                  tabButtons.current[nextTab]?.focus();
                }}
              >
                {words(item)}
              </button>
            ))}
          </div>
          {busy && <span role="status">{words('Calculating…')}</span>}
        </div>
        {availableTabs.map((panelTab, index) => (
          <div
            key={panelTab}
            id={`${ribbonId}-panel-${index}`}
            className="sheet-ribbon-tabpanel"
            role="tabpanel"
            data-formula-draft-zone={panelTab === 'Home' ? true : undefined}
            aria-labelledby={`${ribbonId}-tab-${index}`}
            hidden={panelTab !== tab}
          >
            <div
              className="sheet-ribbon"
              data-ribbon-tab={panelTab}
              data-formula-library-layout={
                panelTab === 'Formulas' ? (width >= 1600 ? 'row' : width <= 420 ? 'narrow' : 'compact') : undefined
              }
              role="toolbar"
              aria-label={words(`Spreadsheet ${panelTab}`)}
              onScroll={(event) => {
                if (event.target !== event.currentTarget) return;
                setOpenGroup('');
                root.current?.querySelectorAll<HTMLDetailsElement>('details[open]').forEach((menu) => {
                  menu.open = false;
                });
              }}
              onClick={(event) => {
                const button = (event.target as HTMLElement).closest('details.sheet-ribbon-menu > div button');
                const menu = button?.closest<HTMLDetailsElement>('details');
                if (menu) {
                  menu.open = false;
                  menu.querySelector('summary')?.focus();
                }
                if ((event.target as HTMLElement).closest('.is-collapsed .sheet-ribbon-group-content button')) {
                  root.current
                    ?.querySelector<HTMLButtonElement>('.sheet-ribbon-group-toggle[aria-expanded="true"]')
                    ?.focus();
                  setOpenGroup('');
                }
              }}
            >
              {panelTab === 'Home' && children}
              {panelTab === 'Data' && dataCommands}
              {panelTab === 'Formulas' && formulaCommands}
              {panelTab === 'View' && viewCommands}
              {groupsFor(panelTab)}
              {panelTab === 'View' && viewToolbar}
              {panelTab === 'Draw' && ink && (
                <RibbonGroup name="Pens" icon="table" collapseBelow={0}>
                  <div className="sheet-ribbon-body">
                    <WorkbookInkPalette
                      selected={ink.selected}
                      disabled={capability('draw.preset', !!ink.onChoose, ink.disabled, ink.reason).state !== 'enabled'}
                      reason={ink.reason}
                      warning={ink.warning}
                      words={words}
                      onChoose={(pen) => void execute('draw.preset', ink.onChoose, { pen })}
                    />
                  </div>
                  <small>{words('Pens')}</small>
                </RibbonGroup>
              )}
            </div>
          </div>
        ))}
      </div>
    </RibbonContext.Provider>
  );
}

function RibbonZoomInput({
  zoom,
  disabled,
  reason,
  onCommit,
}: {
  zoom: number;
  disabled: boolean;
  reason?: string;
  onCommit: (zoom: number) => Promise<void>;
}): JSX.Element {
  const words = useDocumentWords(),
    [draft, setDraft] = useState(String(zoom)),
    [error, setError] = useState(''),
    pending = useRef(false);
  useEffect(() => {
    setDraft(String(zoom));
    setError('');
  }, [zoom]);
  const commit = async (): Promise<void> => {
    if (disabled || pending.current) return;
    const value = Number(draft);
    if (!draft.trim() || !Number.isInteger(value) || value < 10 || value > 400) {
      setError(words('Enter a whole percent from 10 to 400.'));
      return;
    }
    setError('');
    if (value === zoom) return;
    pending.current = true;
    try {
      await onCommit(value);
    } finally {
      pending.current = false;
    }
  };
  return (
    <div className="sheet-ribbon-zoom-input">
      <label title={words(reason ?? 'Worksheet zoom from 10 to 400 percent.')}>
        {words('Zoom')}{' '}
        <input
          type="number"
          min={10}
          max={400}
          step={1}
          aria-label={words('Zoom')}
          data-ribbon-command="view.zoom"
          value={draft}
          disabled={disabled}
          onChange={(event) => setDraft(event.target.value)}
          onBlur={() => void commit()}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              event.stopPropagation();
              void commit();
            } else if (event.key === 'Escape') {
              event.preventDefault();
              event.stopPropagation();
              setDraft(String(zoom));
              setError('');
            }
          }}
        />
        <span aria-hidden="true">%</span>
      </label>
      {error && <small role="alert">{error}</small>}
    </div>
  );
}

export function RibbonGroup({
  name,
  icon,
  collapseBelow,
  children,
}: {
  name: string;
  icon: string;
  collapseBelow: number;
  children: ReactNode;
}): JSX.Element {
  const words = useDocumentWords(),
    id = useId();
  const { width, openGroup, setOpenGroup } = useContext(RibbonContext);
  const collapsed = width < collapseBelow,
    open = collapsed && openGroup === name;
  const trigger = useRef<HTMLButtonElement>(null),
    panel = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (!open || !panel.current || !trigger.current) return;
    const anchor = trigger.current.getBoundingClientRect(),
      box = panel.current.getBoundingClientRect();
    panel.current.style.left = `${Math.max(8, Math.min(anchor.left, window.innerWidth - box.width - 8))}px`;
    panel.current.style.top = `${Math.max(8, Math.min(anchor.bottom + 6, window.innerHeight - box.height - 8))}px`;
  }, [open]);
  return (
    <div
      data-ribbon-group={name}
      className={`sheet-ribbon-group sheet-ribbon-${name.toLowerCase()}${collapsed ? ' is-collapsed' : ''}`}
    >
      {collapsed && (
        <button
          ref={trigger}
          className="sheet-ribbon-group-toggle"
          aria-label={words(name)}
          aria-expanded={open}
          aria-controls={id}
          onClick={() => setOpenGroup(open ? '' : name)}
          onKeyDown={(event) => {
            if (!['ArrowDown', 'Enter', ' '].includes(event.key)) return;
            event.preventDefault();
            const opening = event.key === 'ArrowDown' || !open;
            setOpenGroup(opening ? name : '');
            if (opening)
              requestAnimationFrame(() => {
                if (!panel.current?.isConnected || panel.current.closest('[hidden]')) return;
                panel.current
                  .querySelector<HTMLElement>('button:not(:disabled), select:not(:disabled), summary')
                  ?.focus();
              });
          }}
        >
          <Icon name={icon} size={24} />
          <span>{words(name)}</span>
          <Icon name="chevron" size={12} />
        </button>
      )}
      <div
        id={id}
        ref={panel}
        className="sheet-ribbon-group-content"
        hidden={collapsed && !open}
        role="group"
        aria-label={words(name)}
      >
        {children}
      </div>
    </div>
  );
}
