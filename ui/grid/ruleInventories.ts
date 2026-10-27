import type { GridConditionalControls2, GridTables, GridValidationControls4 } from './gridTypes';
import type { WorkbookConditionalDefinition, WorkbookSheet, WorkbookValidationDefinition } from '../../core/contracts';
import { useEffect, useRef, useState } from 'react';

export interface RuleInventoriesInputs {
  conditionalControls: GridConditionalControls2;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  ownerBusy: boolean;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  tables: GridTables;
  validationControls: GridValidationControls4;
}

export function useRuleInventories({
  conditionalControls,
  namedRanges,
  ownerBusy,
  setSheetName,
  sheet,
  sheets,
  tables,
  validationControls,
}: RuleInventoriesInputs) {
  const conditionalScope = conditionalControls
    ? `${conditionalControls.documentId}\0${conditionalControls.revision}\0${sheet?.name}`
    : '';
  const liveConditionalScope = useRef(conditionalScope);
  liveConditionalScope.current = conditionalScope;
  const conditionalReader = useRef(conditionalControls?.inventory);
  conditionalReader.current = conditionalControls?.inventory;
  const [acceptedConditional, setAcceptedConditional] = useState<{
    scope: string;
    definition: WorkbookConditionalDefinition;
  } | null>(null);
  useEffect(() => {
    let cancelled = false;
    setAcceptedConditional(null);
    if (!conditionalScope || !sheet || ownerBusy || !conditionalReader.current) return;
    void conditionalReader
      .current(sheet.name)
      .then((definition) => {
        if (!cancelled && liveConditionalScope.current === conditionalScope)
          setAcceptedConditional({ scope: conditionalScope, definition });
      })
      .catch((error) => {
        if (!cancelled && liveConditionalScope.current === conditionalScope)
          console.warn('xlsx.conditional-inventory', String(error));
      });
    return () => {
      cancelled = true;
    };
  }, [conditionalScope, ownerBusy]);
  const validationScope = validationControls
    ? `${validationControls.documentId}\0${validationControls.revision}\0${sheet?.name}`
    : '';
  const liveValidationScope = useRef(validationScope);
  liveValidationScope.current = validationScope;
  const validationReader = useRef(validationControls?.inventory);
  validationReader.current = validationControls?.inventory;
  const [acceptedValidation, setAcceptedValidation] = useState<{
    scope: string;
    definition: WorkbookValidationDefinition;
  } | null>(null);
  useEffect(() => {
    let cancelled = false;
    setAcceptedValidation(null);
    if (!validationScope || !sheet || ownerBusy || !validationReader.current) return;
    void validationReader
      .current(sheet.name)
      .then((definition) => {
        if (!cancelled && liveValidationScope.current === validationScope)
          setAcceptedValidation({ scope: validationScope, definition });
      })
      .catch((error) => {
        if (!cancelled && liveValidationScope.current === validationScope)
          console.warn('xlsx.validation-inventory', String(error));
      });
    return () => {
      cancelled = true;
    };
  }, [validationScope, ownerBusy]);
  const [namePage, setNamePage] = useState(0);
  const [tablePage, setTablePage] = useState(0);
  useEffect(
    () => setNamePage((page) => Math.min(page, Math.max(0, Math.ceil(namedRanges.length / 100) - 1))),
    [namedRanges.length]
  );
  useEffect(
    () => setTablePage((page) => Math.min(page, Math.max(0, Math.ceil(tables.length / 100) - 1))),
    [tables.length]
  );
  const [newRangeName, setNewRangeName] = useState('');
  const [nameScope, setNameScope] = useState<'workbook' | 'worksheet'>('workbook');
  const pendingSheetSelection = useRef<string | null>(null);
  useEffect(() => {
    if (pendingSheetSelection.current && sheets.some((item) => item.name === pendingSheetSelection.current)) {
      setSheetName(pendingSheetSelection.current);
      pendingSheetSelection.current = null;
    }
  }, [sheets]);
  return {
    conditionalScope,
    liveConditionalScope,
    acceptedConditional,
    setAcceptedConditional,
    validationScope,
    liveValidationScope,
    acceptedValidation,
    namePage,
    setNamePage,
    tablePage,
    setTablePage,
    newRangeName,
    setNewRangeName,
    nameScope,
    setNameScope,
    pendingSheetSelection,
  } as const;
}
