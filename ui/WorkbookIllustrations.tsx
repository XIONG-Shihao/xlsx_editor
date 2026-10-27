import { useEffect, useMemo, useRef } from 'react';
import { WorkbookShapes, type IllustrationMode } from './WorkbookShapes';
import { WorkbookModel3D } from './WorkbookModel3D';
import { WorkbookSmartArt } from './WorkbookSmartArt';
import { createWorksheetObjectController, type WorksheetObjectFrame } from './worksheetObjectController';
/** One physical inventory/write lane survives every gallery mode and frame retirement. */
export function WorkbookIllustrations({
  frame,
  mode,
  selection,
  words,
}: {
  frame: WorksheetObjectFrame;
  mode: IllustrationMode;
  selection: string;
  words?: (text: string) => string;
}): JSX.Element {
  const controller = useMemo(() => createWorksheetObjectController(), []),
    disposal = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (disposal.current) clearTimeout(disposal.current);
    return () => {
      controller.setFrame(null);
      disposal.current = setTimeout(() => controller.dispose(), 0);
    };
  }, [controller]);
  return mode === 'model3d' ? (
    <WorkbookModel3D frame={frame} selection={selection} controller={controller} words={words} />
  ) : mode === 'smartart' ? (
    <WorkbookSmartArt frame={frame} selection={selection} controller={controller} words={words} />
  ) : (
    <WorkbookShapes frame={frame} mode={mode} selection={selection} controller={controller} words={words} />
  );
}
