/**
 * ExcelApp.tsx
 * 
 * Complete Excel application shell with title bar, ribbon, formula bar,
 * spreadsheet area, sheet tabs, and status bar.
 */

import React, { useState, useCallback, useMemo, useEffect, useRef } from 'react';
import type { Workbook, CellComment } from '@cyber-sheet/core';
import type { CanvasRenderer } from '@cyber-sheet/renderer-canvas';
import { CommandManager, DrawingLayer, ClipboardService, ClearCellsCommand, FormulaEngine, DropdownList, FileOperations, SetViewModeCommand, SetHeaderFooterCommand, FormattingController, SetHyperlinkCommand, getDefaultHyperlinkScreenTip, SortCommand, ToggleAutoFilterCommand, ClearFilterCommand, FormatFormControlCommand, type ViewMode, type InsertCellsMode, type DeleteCellsMode, type Address, type CellHyperlink, type FormControlObject, type HeaderFooterSettings } from '@cyber-sheet/core';
import { loadXlsxFromArrayBuffer } from '@cyber-sheet/io-xlsx';
import { TitleBar } from './TitleBar';
import { RibbonTabs } from './RibbonTabs';
import { Ribbon } from '../components/ribbon/Ribbon';
import { FormulaBar } from '../FormulaBar';
import { CyberSheet } from '../CyberSheet';
import { SheetTabs } from '../SheetTabs';
import { StatusBar } from './StatusBar';
import { DrawingCanvas, type DrawingCanvasHandle, type PictureInsertTemplate } from './DrawingCanvas';
import type { IconInsertTemplate } from '../utils/createDrawingObject';
import type { FormControlInsertTemplate } from '../utils/formControlFactory';
import type { TextBoxInsertTemplate } from '../utils/textBoxFactory';
import { createWordArtTemplate } from '../utils/wordArtFactory';
import { CellEditOverlay } from './CellEditOverlay';
import { CutRangeOverlay } from './CutRangeOverlay';
import { ContextMenu, ContextMenuItem } from './ContextMenu';
import { MiniToolbar } from './MiniToolbar';
import FormatCellsDialog, { FormattingChanges } from './dialogs/FormatCellsDialog/FormatCellsDialog';
import { FormatControlDialog } from './dialogs/FormatControlDialog';
import { getFormControlDefaultSize } from '../utils/formControlFactory';
import {
  afterFormatControlApplied,
  buildFormControlFormatUpdates,
  type FormControlDialogDraft,
} from '../utils/formatControlApply';
import { formatAddressA1 } from '../utils/parseA1Reference';
import {
  formatSelectionAsReference,
  isFormulaEditValue,
  mergePointReference,
} from '../utils/formulaEdit';
import { executeAutoSum, type AutoSumFunction } from '../utils/autoSum';
import FindReplaceDialog from './dialogs/FindReplaceDialog';
import { CellCommentDialog } from './dialogs/CellCommentDialog';
import { HeaderFooterDialog } from './dialogs/HeaderFooterDialog';
import { WordArtGalleryDialog } from './dialogs/WordArtGalleryDialog';
import { CommentPanel } from './CommentPanel';
import { InsertHyperlinkDialog, type InsertHyperlinkDialogResult } from './dialogs/InsertHyperlinkDialog';
import { InsertDeleteCellsDialog } from './dialogs/InsertDeleteCellsDialog';
import { CreatePivotTableDialog } from './dialogs/CreatePivotTableDialog';
import { CreateTableDialog } from './dialogs/CreateTableDialog';
import {
  executeInsertCells,
  executeDeleteCells,
  getInsertDeleteInvalidateBounds,
  type CellRange,
} from '../utils/insertDeleteCells';
import { BackstageContainer, BackstagePanel } from './backstage/BackstageContainer';
import { debugEdit, debugRender, debugMenu } from '../utils/debug';
import { configureCyberSheet } from '../config/globalConfig';
import { CyberSheetConfigProvider, useCyberSheetAppConfig, useCyberSheetConfig } from '../config/CyberSheetConfigContext';
import { getVisibleRibbonTabLabels, isTabEnabled, type CyberSheetConfigInput } from '../config/appConfig';
import { canEditCellValue, canStartCellEdit } from '../config/commandPermissions';
import { createGuardedCommandManager } from '../config/guardedCommandManager';
import { attachOnEventForwarder } from '../config/eventForwarder';
import { handleSpreadsheetHistoryShortcut } from '../utils/spreadsheetHistory';
import {
  performCopy,
  performCut,
  performPasteAsync,
  normalizeSelectionRange,
  type ClipboardCopyMode,
} from '../utils/clipboardOperations';
import { normalizeRange } from '../utils/sortFilterCommands';
import { insertPivotTable } from '../utils/insertPivotTable';
import { applyTableFormat, DEFAULT_INSERT_TABLE_STYLE } from '../utils/createTable';
import { resolveTableDataRange } from '../utils/conditionalFormattingRibbon';
import { openHyperlinkTarget, downloadBlankWorkbook } from '../utils/hyperlinkNavigation';
import {
  isCtrlLetter,
  isCtrlPhysicalKey,
  isArrowKey,
  getTypedCharacter,
  hasCtrlOrMeta,
} from '../utils/keyboardLayout';
import './excel-app.css';

type NumberFormatCategory = 'general' | 'currency' | 'percentage' | 'number' | 'comma';

const MAX_EXPANDED_SELECTION_ADDRESSES = 10000;

function addressesFromRange(range: CellRange | null | undefined): Address[] {
  if (!range?.start || !range?.end) return [];
  const r1 = Math.min(range.start.row, range.end.row);
  const r2 = Math.max(range.start.row, range.end.row);
  const c1 = Math.min(range.start.col, range.end.col);
  const c2 = Math.max(range.start.col, range.end.col);
  const cellCount = (r2 - r1 + 1) * (c2 - c1 + 1);

  if (cellCount > MAX_EXPANDED_SELECTION_ADDRESSES) {
    const start = { row: r1, col: c1 };
    const end = { row: r2, col: c2 };
    return start.row === end.row && start.col === end.col ? [start] : [start, end];
  }

  const addresses: Address[] = [];
  for (let row = r1; row <= r2; row++) {
    for (let col = c1; col <= c2; col++) {
      addresses.push({ row, col });
    }
  }
  return addresses;
}

function getUsedDataRange(sheet: any): CellRange {
  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxCol = 0;
  const forEachNonEmptyCell = sheet?.forEachNonEmptyCell;

  if (typeof forEachNonEmptyCell === 'function') {
    forEachNonEmptyCell.call(sheet, (row: number, col: number, cell: { value?: unknown; formula?: string }) => {
      const hasValue = cell?.value !== undefined && cell.value !== null && cell.value !== '';
      if (!hasValue && !cell?.formula) return;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    });
  }

  if (maxRow === 0 || maxCol === 0) {
    return { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } };
  }

  return {
    start: { row: minRow, col: minCol },
    end: { row: maxRow, col: maxCol },
  };
}

function resolveStyleColor(color: unknown): string {
  if (typeof color === 'string') return color;
  return '#000000';
}

function resolveStyleFillColor(fill: unknown): string {
  if (typeof fill === 'string') return fill;
  if (fill && typeof fill === 'object' && 'fgColor' in fill) {
    const fg = (fill as { fgColor?: unknown }).fgColor;
    if (typeof fg === 'string') return fg;
  }
  return '#FFFF00';
}

function getNumberFormatCategory(format: string | undefined): NumberFormatCategory {
  if (!format || format === 'General') return 'general';
  if (format.includes('$')) return 'currency';
  if (format.includes('%')) return 'percentage';
  if (format === '#,##0') return 'comma';
  return 'number';
}

function createUniqueSheetName(workbook: Workbook): string {
  const names = new Set(workbook.getSheetNames());
  let index = 1;
  while (names.has(`Sheet${index}`)) {
    index += 1;
  }
  return `Sheet${index}`;
}

export interface ExcelAppProps {
  workbook: Workbook;
  fileName?: string;
  onSave?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onWorkbookLoaded?: (workbook: Workbook) => void;
  config?: CyberSheetConfigInput;
  onEvent?: (event: string, payload: unknown) => void;
  style?: React.CSSProperties;
}

/**
 * ExcelApp - Complete Excel application shell
 * 
 * Provides the full Excel-like interface with:
 * - Title bar with quick access toolbar
 * - Ribbon with tab navigation
 * - Formula bar
 * - Spreadsheet canvas
 * - Sheet tabs
 * - Status bar with zoom controls
 */
const ExcelAppView: React.FC<ExcelAppProps> = ({
  workbook,
  fileName = 'Book1',
  onSave,
  onUndo,
  onRedo,
  onWorkbookLoaded,
  onEvent,
  style,
}) => {
  const appConfig = useCyberSheetAppConfig();
  const cyberSheetConfig = useCyberSheetConfig();

  const [activeTab, setActiveTab] = useState<string>(() => {
    const visible = getVisibleRibbonTabLabels(appConfig);
    return visible[0] ?? 'Home';
  });
  const [activeSheet, setActiveSheet] = useState<string>(workbook.activeSheet?.name || 'Sheet1');
  const [sheetRevision, setSheetRevision] = useState(0);
  const [renameSheetTrigger, setRenameSheetTrigger] = useState(0);
  const [zoom, setZoom] = useState<number>(100);
  const [viewMode, setViewMode] = useState<ViewMode>(() => {
    return ((workbook.activeSheet as any)?.viewMode as ViewMode) || 'normal';
  });
  const [selectedCell, setSelectedCell] = useState<any>(null);
  const [cellValue, setCellValue] = useState<any>('');
  const [cellFormula, setCellFormula] = useState<string>('');
  const [isEditingFormula, setIsEditingFormula] = useState<boolean>(false);
  const [formulaBarValue, setFormulaBarValue] = useState<string>(''); // Track formula bar input
  const [selection, setSelection] = useState<any>(null);
  const [renderer, setRenderer] = useState<CanvasRenderer | null>(null);
  const [scrollLeft, setScrollLeft] = useState<number>(0);
  const [scrollTop, setScrollTop] = useState<number>(0);
  const [viewportWidth, setViewportWidth] = useState<number>(1000);
  const [viewportHeight, setViewportHeight] = useState<number>(600);
  const [isAutoFilterEnabled, setIsAutoFilterEnabled] = useState<boolean>(true);
  const [headerFilterIcons, setHeaderFilterIcons] = useState<Array<{ col: number; x: number; w: number }>>([]);
  const [filterMenu, setFilterMenu] = useState<{
    col: number;
    x: number;
    y: number;
    values: Array<{ value: string; count: number }>;
    apply: (selected: string[]) => void;
    clear: () => void;
  } | null>(null);
  const [filterSearch, setFilterSearch] = useState('');
  const [filterSel, setFilterSel] = useState<Set<string>>(new Set());
  const [historyVersion, setHistoryVersion] = useState(0);

  useEffect(() => {
    configureCyberSheet({ fonts: cyberSheetConfig.fonts });
  }, [cyberSheetConfig.fonts]);

  useEffect(() => {
    if (!onEvent) return;
    return attachOnEventForwarder(workbook.eventBus, onEvent, appConfig.eventFilter);
  }, [workbook, onEvent, appConfig.eventFilter]);

  useEffect(() => {
    if (!isTabEnabled(appConfig, activeTab)) {
      const visible = getVisibleRibbonTabLabels(appConfig);
      if (visible.length > 0) setActiveTab(visible[0]!);
    }
  }, [appConfig, activeTab]);

  // In-cell editing state
  const [inCellEdit, setInCellEdit] = useState<{
    cell: { row: number; col: number };
    bounds: { x: number; y: number; width: number; height: number };
    initialValue: string;
    currentValue: string;
    cursorPosition: number;
  } | null>(null);
  
  // Debug: Log when edit mode changes
  useEffect(() => {
    console.log('📊 inCellEdit state changed:', inCellEdit ? `cell (${inCellEdit.cell.row}, ${inCellEdit.cell.col})` : 'null');
  }, [inCellEdit]);

  // Cell reference picking mode (when editing formula)
  const [isPickingReference, setIsPickingReference] = useState<boolean>(false);

  // Context menu and mini toolbar state
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number } | null>(null);
  const [contextMenuLayout, setContextMenuLayout] = useState<{ x: number; y: number; width: number; height: number } | null>(null);
  const [miniToolbar, setMiniToolbar] = useState<{ x: number; y: number } | null>(null);
  const [insertDeleteDialog, setInsertDeleteDialog] = useState<{
    type: 'insert' | 'delete';
    range: CellRange;
  } | null>(null);
  
  // Format Cells dialog state
  const [isFormatDialogOpen, setIsFormatDialogOpen] = useState<boolean>(false);

  // Format Control dialog state
  const [formatControlObjectId, setFormatControlObjectId] = useState<string | null>(null);
  const [formControlContextMenu, setFormControlContextMenu] = useState<{
    objectId: string;
    x: number;
    y: number;
  } | null>(null);
  const [formControlCellLinkPicker, setFormControlCellLinkPicker] = useState(false);
  const [pickedFormControlCellLink, setPickedFormControlCellLink] = useState<string | null>(null);
  
  // Find/Replace dialog state
  const [isFindReplaceOpen, setIsFindReplaceOpen] = useState<boolean>(false);
  const [findReplaceTab, setFindReplaceTab] = useState<'find' | 'replace'>('find');
  const [isCreatePivotTableOpen, setIsCreatePivotTableOpen] = useState(false);
  const [isCreateTableOpen, setIsCreateTableOpen] = useState(false);
  const [tableInsertError, setTableInsertError] = useState<string | null>(null);
  const [pivotTableError, setPivotTableError] = useState<string | null>(null);
  const [commentDialog, setCommentDialog] = useState<{
    address: { row: number; col: number };
    comments: CellComment[];
  } | null>(null);
  const [headerFooterDialogOpen, setHeaderFooterDialogOpen] = useState(false);
  const [wordArtGalleryOpen, setWordArtGalleryOpen] = useState(false);
  const [showCommentPanel, setShowCommentPanel] = useState(appConfig.showCommentPanel);
  const [hyperlinkDialog, setHyperlinkDialog] = useState<{
    address: Address;
    existingHyperlink: CellHyperlink | null;
    displayText: string;
  } | null>(null);
  const [commentTooltip, setCommentTooltip] = useState<{
    x: number;
    y: number;
    text: string;
  } | null>(null);
  const [hyperlinkTooltip, setHyperlinkTooltip] = useState<{
    x: number;
    y: number;
    text: string;
  } | null>(null);
  const [statusBarLink, setStatusBarLink] = useState<string | null>(null);
  
  // Recent colors tracking (up to 10 each)
  const [recentFillColors, setRecentFillColors] = useState<string[]>([]);
  const [recentFontColors, setRecentFontColors] = useState<string[]>([]);
  
  // Cut range tracking (for visual indication with marching ants border)
  const [cutRange, setCutRange] = useState<{ start: { row: number; col: number }; end: { row: number; col: number } } | null>(null);
  
  // Data validation dropdown state
  const [validationDropdown, setValidationDropdown] = useState<{
    items: string[];
    cellAddress: { row: number; col: number };
    cellBounds: { x: number; y: number; width: number; height: number };
    currentValue: string;
  } | null>(null);
  
  // Backstage file menu state
  const [backstageOpen, setBackstageOpen] = useState<boolean>(false);
  const [backstagePanel, setBackstagePanel] = useState<BackstagePanel>('new');
  
  // FileOperations instance for backstage panels
  const fileOperations = useMemo(() => new FileOperations({
    id: `workbook-${Date.now()}`,
    name: fileName,
    path: `/Documents/${fileName}.xlsx`,
    location: 'local',
    size: 0,
    created: new Date(),
    lastModified: new Date(),
    lastModifiedBy: 'User',
    author: 'User',
    sheets: workbook.getSheetNames().length,
    tags: [],
    isProtected: false,
    isMarkedFinal: false,
  }), [fileName, workbook]);
  
  const workbookMetadata = useMemo(() => fileOperations.getMetadata(), [fileOperations]);
  
  // Refs for keyboard handler to avoid re-attaching listeners on every state change
  const selectedCellRef = useRef(selectedCell);
  const selectionRef = useRef(selection);
  const inCellEditRef = useRef(inCellEdit);
  const isEditingFormulaRef = useRef(isEditingFormula);
  const formulaBarValueRef = useRef(formulaBarValue);
  const formulaBarCursorRef = useRef(0);
  const isPickingReferenceRef = useRef(isPickingReference);
  const isFindReplaceOpenRef = useRef(isFindReplaceOpen);
  // Context menu selection: captured at menu open time and frozen until menu closes
  const contextMenuSelectionRef = useRef<any>(null);
  const contextMenuTargetCellRef = useRef<Address | null>(null);
  
  // Keep refs in sync with state (but NOT contextMenuSelectionRef - that's set by context menu handler)
  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);
  
  useEffect(() => {
    selectedCellRef.current = selectedCell;
  }, [selectedCell]);
  
  useEffect(() => {
    inCellEditRef.current = inCellEdit;
  }, [inCellEdit]);
  
  useEffect(() => {
    isEditingFormulaRef.current = isEditingFormula;
  }, [isEditingFormula]);

  useEffect(() => {
    formulaBarValueRef.current = formulaBarValue;
  }, [formulaBarValue]);

  useEffect(() => {
    isPickingReferenceRef.current = isPickingReference;
  }, [isPickingReference]);

  useEffect(() => {
    isFindReplaceOpenRef.current = isFindReplaceOpen;
  }, [isFindReplaceOpen]);
  
  // Helper to add color to recent list
  const addRecentColor = (color: string, setRecentColors: React.Dispatch<React.SetStateAction<string[]>>) => {
    setRecentColors((prev) => {
      const filtered = prev.filter((c) => c !== color);
      return [color, ...filtered].slice(0, 10);
    });
  };

  const isAddressInSelection = useCallback((addr: Address, range: { start: Address; end: Address }) => {
    const r1 = Math.min(range.start.row, range.end.row);
    const r2 = Math.max(range.start.row, range.end.row);
    const c1 = Math.min(range.start.col, range.end.col);
    const c2 = Math.max(range.start.col, range.end.col);
    return addr.row >= r1 && addr.row <= r2 && addr.col >= c1 && addr.col <= c2;
  }, []);

  const contextMenuTarget = useCallback(() => {
    return contextMenuTargetCellRef.current ?? contextMenuSelectionRef.current?.start ?? null;
  }, []);

  const formatCellAddress = useCallback((addr: { row: number; col: number }) => {
    let colNum = Math.max(1, addr.col);
    let colLabel = '';
    while (colNum > 0) {
      const rem = (colNum - 1) % 26;
      colLabel = String.fromCharCode(65 + rem) + colLabel;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return `${colLabel || 'A'}${Math.max(1, addr.row)}`;
  }, []);

  const openCommentEditorForAddress = useCallback((address: { row: number; col: number }) => {
    if (!appConfig.enableComments) return;
    const sheet = workbook.activeSheet;
    if (!sheet) return;
    const existing = sheet.getComments(address);
    setCommentDialog({ address, comments: existing });
  }, [workbook, appConfig.enableComments]);

  const openHyperlinkEditorForAddress = useCallback((address: Address) => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;
    const existing = sheet.getHyperlink(address);
    const raw = sheet.getCellValue(address);
    const displayText = raw != null && raw !== '' ? String(raw) : '';
    setHyperlinkDialog({
      address,
      existingHyperlink: existing ?? null,
      displayText,
    });
  }, [workbook]);

  const recomputeHeaderFilterIcons = useCallback(() => {
    const r = renderer;
    if (!r || !isAutoFilterEnabled) {
      setHeaderFilterIcons([]);
      return;
    }

    try {
      const opts = r.optionsReadonly ?? { headerWidth: 48, headerHeight: 24 };
      const headerWidth = opts.headerWidth;
      const vp = r.getViewportSize?.() || { width: 0, height: 0 };
      const sheet = r.sheetReadonly;
      const sx = scrollLeft;
      const zoomFactor = typeof (r as any).getZoom === 'function' ? (r as any).getZoom() : 1;
      const icons: Array<{ col: number; x: number; w: number }> = [];

      let col = 1;
      let x = headerWidth - sx;
      while (col <= sheet.colCount) {
        const cw = sheet.getColumnWidth(col) * zoomFactor;
        if (x + cw > headerWidth) break;
        x += cw;
        col++;
      }

      const limit = headerWidth + vp.width;
      while (col <= sheet.colCount && x < limit) {
        const cw = sheet.getColumnWidth(col) * zoomFactor;
        icons.push({ col, x, w: cw });
        x += cw;
        col++;
      }

      setHeaderFilterIcons(icons);
    } catch {
      setHeaderFilterIcons([]);
    }
  }, [renderer, isAutoFilterEnabled, scrollLeft]);

  const baseCommandManager = useMemo(() => {
    const sheet = workbook.activeSheet;
    return new CommandManager(100, sheet, workbook.eventBus);
  }, [workbook]);

  const commandManager = useMemo(
    () => createGuardedCommandManager(baseCommandManager, appConfig),
    [baseCommandManager, appConfig],
  );

  const formattingController = useMemo(() => {
    const sheet = workbook.activeSheet;
    if (!sheet) return null;
    return new FormattingController(sheet, commandManager);
  }, [workbook.activeSheet, commandManager]);

  const selectionAddresses = useMemo(() => addressesFromRange(selection), [selection]);

  const selectionRange = useMemo(() => normalizeSelectionRange(selection), [selection]);

  const miniToolbarAddresses = useMemo(() => {
    if (!contextMenu && !miniToolbar) return [];
    return addressesFromRange(contextMenuSelectionRef.current);
  }, [contextMenu, miniToolbar, historyVersion]);

  const miniToolbarStyleCell = miniToolbarAddresses[0] ?? selection?.start ?? null;
  const miniToolbarCellStyle = miniToolbarStyleCell && workbook.activeSheet
    ? workbook.activeSheet.getCellStyle(miniToolbarStyleCell)
    : undefined;

  const applyContextMenuFormatting = useCallback(
    (apply: (addresses: Address[]) => void) => {
      const addresses = miniToolbarAddresses.length > 0 ? miniToolbarAddresses : selectionAddresses;
      if (!formattingController || addresses.length === 0) return;
      apply(addresses);
      const rows = addresses.map((a) => a.row);
      const cols = addresses.map((a) => a.col);
      renderer?.invalidateRange(
        Math.min(...rows),
        Math.min(...cols),
        Math.max(...rows),
        Math.max(...cols),
      );
      renderer?.scheduleRedraw?.();
    },
    [miniToolbarAddresses, selectionAddresses, formattingController, renderer],
  );

  useEffect(() => {
    return commandManager.subscribe(() => {
      const sheet = workbook.activeSheet;
      const cell = selectedCellRef.current;
      if (sheet && cell) {
        const formula = (sheet as any).getCellFormula?.(cell);
        const value = sheet.getCellValue(cell);
        setCellValue(formula != null && formula !== '' ? formula : (value ?? ''));
      }
      renderer?.redraw?.();
      setHistoryVersion((version) => version + 1);
    });
  }, [commandManager, workbook, renderer]);

  useEffect(() => {
    const handleHistoryShortcut = (event: KeyboardEvent) => {
      handleSpreadsheetHistoryShortcut(event, commandManager);
    };
    window.addEventListener('keydown', handleHistoryShortcut, true);
    return () => window.removeEventListener('keydown', handleHistoryShortcut, true);
  }, [commandManager]);

  const executeFilterStateCommand = useCallback(
    (applyChange: (sheet: any) => void, description: string) => {
      const sheet = workbook.activeSheet as any;
      if (!sheet) return;
      const before = new Map(sheet.getAllFilters?.() ?? []);
      const refresh = () => renderer?.redraw?.();

      commandManager.execute({
        description,
        execute: () => {
          applyChange(sheet);
          refresh();
        },
        undo: () => {
          sheet.clearAllFilters?.();
          for (const [col, filter] of before) {
            sheet.setColumnFilter(col, filter);
          }
          refresh();
        },
      });
    },
    [workbook, commandManager, renderer]
  );

  const handleRibbonFilter = useCallback((action: 'toggle' | 'clear' | 'reapply') => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    if (action === 'toggle') {
      if (!selection?.start || !selection?.end) return;
      const filterRange = normalizeRange({ start: selection.start, end: selection.end });
      const enabling = !isAutoFilterEnabled;

      if (enabling) {
        commandManager.execute(new ToggleAutoFilterCommand(sheet, filterRange, true));
        setIsAutoFilterEnabled(true);
        setTimeout(() => recomputeHeaderFilterIcons(), 0);
      } else {
        commandManager.execute(new ToggleAutoFilterCommand(sheet, filterRange, false));
        executeFilterStateCommand(
          (activeSheet) => activeSheet.clearAllFilters?.(),
          'Clear All Filters',
        );
        setIsAutoFilterEnabled(false);
        setFilterMenu(null);
      }
      renderer?.redraw?.();
      return;
    }

    if (action === 'clear') {
      commandManager.execute(new ClearFilterCommand(sheet));
      renderer?.redraw?.();
      return;
    }

    if (action === 'reapply') {
      renderer?.redraw?.();
    }
  }, [
    workbook,
    selection,
    isAutoFilterEnabled,
    commandManager,
    executeFilterStateCommand,
    recomputeHeaderFilterIcons,
    renderer,
  ]);

  const handleInsertPivotTable = useCallback(() => {
    setPivotTableError(null);
    setIsCreatePivotTableOpen(true);
  }, []);

  const handleInsertTable = useCallback(() => {
    setTableInsertError(null);
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    const range = resolveTableDataRange(sheet, selectionRange);
    if (!range) {
      setTableInsertError('Select a data range with at least one cell, then try again.');
      return;
    }

    setIsCreateTableOpen(true);
  }, [workbook, selectionRange]);

  const handleCreateTable = useCallback((range: CellRange, hasHeaders: boolean) => {
    const sheet = workbook.activeSheet;
    if (!sheet || !formattingController) return;

    applyTableFormat(formattingController, commandManager, sheet, range, hasHeaders);
    setIsCreateTableOpen(false);
    setTableInsertError(null);
    renderer?.invalidateRange?.(range.start.row, range.start.col, range.end.row, range.end.col);
    renderer?.scheduleRedraw?.();
    setHistoryVersion((version) => version + 1);
  }, [workbook, formattingController, commandManager, renderer]);

  const handleCreatePivotTable = useCallback((sourceRange: CellRange, target: Address) => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    const result = insertPivotTable(sheet, commandManager, sourceRange, target);
    if (!result.ok) {
      setPivotTableError(result.error);
      return;
    }

    setIsCreatePivotTableOpen(false);
    setPivotTableError(null);
    setSelectedCell(target);
    setSelection({ start: target, end: result.outputRange.end });
    renderer?.setSelection?.({ start: target, end: result.outputRange.end });
    renderer?.invalidateRange?.(
      result.outputRange.start.row,
      result.outputRange.start.col,
      result.outputRange.end.row,
      result.outputRange.end.col,
    );
    renderer?.scrollToCell?.({ row: target.row, col: target.col });
    renderer?.scheduleRedraw?.();
    setHistoryVersion((version) => version + 1);
  }, [workbook.activeSheet, commandManager, renderer]);

  const handleDataCommand = useCallback((command: { type: string; [key: string]: unknown }) => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    switch (command.type) {
      case 'sort':
        commandManager.execute(
          new SortCommand(
            sheet,
            command.range as any,
            command.sortBy as any,
            (command.hasHeaders as boolean | undefined) ?? true,
          ),
        );
        renderer?.redraw?.();
        break;
      case 'toggleAutoFilter': {
        const enabling = (command.enabled as boolean | undefined) ?? true;
        commandManager.execute(
          new ToggleAutoFilterCommand(sheet, command.range as any, enabling),
        );
        setIsAutoFilterEnabled(enabling);
        if (!enabling) {
          executeFilterStateCommand(
            (activeSheet) => activeSheet.clearAllFilters?.(),
            'Clear All Filters',
          );
          setFilterMenu(null);
        } else {
          setTimeout(() => recomputeHeaderFilterIcons(), 0);
        }
        renderer?.redraw?.();
        break;
      }
      case 'clearFilter':
        commandManager.execute(new ClearFilterCommand(sheet));
        renderer?.redraw?.();
        break;
      case 'reapplyFilter':
        renderer?.redraw?.();
        break;
      default:
        console.warn('Unhandled data command:', command.type);
    }
  }, [
    workbook,
    commandManager,
    renderer,
    executeFilterStateCommand,
    recomputeHeaderFilterIcons,
  ]);

  const openColumnFilterMenu = useCallback((col: number, x: number, y: number) => {
    const r = renderer;
    if (!r) return;

    const rows = (r.sheetReadonly as any).getVisibleRowIndicesExcluding?.(col) || [];
    const seen = new Map<string, number>();
    for (const rr of rows) {
      const v = r.sheetReadonly.getCellValue({ row: rr, col });
      const key = v == null ? '' : String(v);
      seen.set(key, (seen.get(key) || 0) + 1);
    }
    const values = Array.from(seen.entries())
      .map(([value, count]) => ({ value, count }))
      .sort((a, b) => a.value.localeCompare(b.value));

    const existing = r.sheetReadonly.getColumnFilter?.(col);
    const allValues = values.map((v) => v.value);
    let selected: Set<string>;
    if (existing && existing.type === 'in' && Array.isArray(existing.value)) {
      selected = new Set((existing.value as any[]).map(String));
    } else {
      const present = new Set<string>();
      try {
        const visibleRows: number[] = r.sheetReadonly.getVisibleRowIndices();
        for (const rr of visibleRows) {
          const v = r.sheetReadonly.getCellValue({ row: rr, col });
          present.add(v == null ? '' : String(v));
        }
      } catch {}
      selected = present.size ? present : new Set(allValues);
    }

    setFilterSel(selected);
    setFilterSearch('');
    setFilterMenu({
      col,
      x,
      y,
      values,
      apply: (sel) => {
        executeFilterStateCommand(
          (sheet) => {
            if (sel.length === allValues.length) {
              sheet.clearColumnFilter(col);
            } else {
              sheet.setColumnFilter(col, { type: 'in', value: sel });
            }
          },
          `Filter Column ${col}`
        );
        setFilterMenu(null);
      },
      clear: () => {
        executeFilterStateCommand(
          (sheet) => {
            sheet.clearColumnFilter(col);
          },
          `Clear Filter Column ${col}`
        );
        setFilterMenu(null);
      },
    });
  }, [renderer, executeFilterStateCommand]);

  // Create formula engine for autocomplete and formula evaluation
  const formulaEngine = useMemo(() => new FormulaEngine(), []);

  useEffect(() => {
    workbook.setFormulaEngine(formulaEngine);
  }, [workbook, formulaEngine]);

  // Create drawing layer instance
  const drawingLayer = useMemo(() => new DrawingLayer(), []);
  const drawingCanvasRef = useRef<DrawingCanvasHandle>(null);

  const handleDrawingChange = useCallback(() => {
    renderer?.scheduleRedraw?.();
  }, [renderer]);

  const clearDrawingSelection = useCallback(() => {
    if (drawingLayer.getSelectedIds().length === 0) return;
    drawingLayer.deselectAll();
  }, [drawingLayer]);

  const handleBeginShapeInsert = useCallback((shapeType: string) => {
    drawingCanvasRef.current?.startShapeInsert(shapeType);
  }, []);

  const handleBeginPictureInsert = useCallback((template: PictureInsertTemplate) => {
    drawingCanvasRef.current?.startPictureInsert(template);
  }, []);

  const handleBeginIconInsert = useCallback((template: IconInsertTemplate) => {
    drawingCanvasRef.current?.startIconInsert(template);
  }, []);

  const handleBeginFormControlInsert = useCallback((template: FormControlInsertTemplate) => {
    drawingCanvasRef.current?.startFormControlInsert(template);
  }, []);

  const handleBeginTextBoxInsert = useCallback((template: TextBoxInsertTemplate) => {
    drawingCanvasRef.current?.startTextBoxInsert(template);
  }, []);

  const handleBeginWordArtInsert = useCallback((template: TextBoxInsertTemplate) => {
    drawingCanvasRef.current?.startWordArtInsert(template);
  }, []);

  const openFormatControlDialog = useCallback((objectId: string) => {
    setFormControlContextMenu(null);
    setPickedFormControlCellLink(null);
    setFormatControlObjectId(objectId);
  }, []);

  const handleApplyFormatControl = useCallback(
    (draft: FormControlDialogDraft) => {
      if (!formatControlObjectId) return;
      const updates = buildFormControlFormatUpdates(draft);
      commandManager.execute(
        new FormatFormControlCommand(drawingLayer, formatControlObjectId, updates),
      );
      afterFormatControlApplied(
        formatControlObjectId,
        draft,
        drawingLayer,
        workbook.activeSheet,
      );
      handleDrawingChange();
    },
    [formatControlObjectId, commandManager, drawingLayer, workbook, handleDrawingChange],
  );
  
  // Create clipboard service instance
  const clipboardService = useMemo(() => new ClipboardService(), []);

  // Get all sheet names
  const sheets = useMemo(
    () => workbook.getSheetNames(),
    [workbook, sheetRevision],
  );

  // Handle sheet change
  const handleSheetChange = useCallback((sheetName: string) => {
    // Set active sheet by setting activeSheetName property
    workbook.activeSheetName = sheetName;
    setActiveSheet(sheetName);
    const sheet = workbook.getSheet(sheetName);
    const mode = ((sheet as any)?.viewMode as ViewMode) || 'normal';
    setViewMode(mode);
    renderer?.setViewMode(mode);
  }, [workbook, renderer]);

  const applyLoadedWorkbook = useCallback((newWorkbook: Workbook) => {
    onWorkbookLoaded?.(newWorkbook);
    setActiveSheet(newWorkbook.activeSheet?.name || 'Sheet1');
    setSheetRevision((v) => v + 1);
    renderer?.scheduleRedraw();
  }, [onWorkbookLoaded, renderer]);

  const handleSpreadsheetDragOver = useCallback((e: React.DragEvent) => {
    if (!appConfig.allowOpen || !onWorkbookLoaded) return;
    if (!e.dataTransfer.types.includes('Files')) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  }, [appConfig.allowOpen, onWorkbookLoaded]);

  const handleSpreadsheetDrop = useCallback(async (e: React.DragEvent) => {
    e.preventDefault();
    if (!appConfig.allowOpen || !onWorkbookLoaded) return;

    const file = e.dataTransfer.files?.[0];
    if (!file) return;

    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = loadXlsxFromArrayBuffer(new Uint8Array(arrayBuffer));
      applyLoadedWorkbook(workbook);
    } catch (err) {
      console.error('Failed to load dropped file:', err);
      alert('Failed to load file. Please ensure it is a valid XLSX file.');
    }
  }, [appConfig.allowOpen, onWorkbookLoaded, applyLoadedWorkbook]);

  const navigateToCommentCell = useCallback((sheetName: string, address: Address) => {
    if (workbook.activeSheet?.name !== sheetName) {
      handleSheetChange(sheetName);
    }
    setSelection({ start: address, end: address });
    setSelectedCell(address);
    renderer?.setSelection({ start: address, end: address });
    renderer?.scrollToCell(address, 'nearest');
    renderer?.scheduleRedraw?.();
  }, [workbook, handleSheetChange, renderer]);

  const handleReviewCommand = useCallback((command: {
    type: string;
    cell?: Address;
    option?: 'show' | 'hide' | 'showIndicator';
  }) => {
    if (!appConfig.enableComments) return;

    const sheet = workbook.activeSheet;
    const target = command.cell ?? selectedCell ?? contextMenuTarget() ?? { row: 1, col: 1 };

    switch (command.type) {
      case 'newComment':
        openCommentEditorForAddress(target);
        break;
      case 'deleteComment': {
        if (!sheet) break;
        const comments = sheet.getComments(target);
        if (comments.length === 0) break;
        const latest = comments[comments.length - 1]!;
        sheet.deleteComment(target, latest.id, true);
        renderer?.invalidateRange(target.row, target.col, target.row, target.col);
        renderer?.scheduleRedraw?.();
        break;
      }
      case 'previousComment': {
        if (!sheet || !selectedCell) break;
        const prev = sheet.getNextCommentCell(selectedCell, 'prev');
        if (prev) navigateToCommentCell(sheet.name, prev);
        break;
      }
      case 'nextComment': {
        if (!sheet || !selectedCell) break;
        const next = sheet.getNextCommentCell(selectedCell, 'next');
        if (next) navigateToCommentCell(sheet.name, next);
        break;
      }
      case 'toggleComments':
        if (command.option === 'show') {
          setShowCommentPanel(true);
        } else if (command.option === 'hide') {
          setShowCommentPanel(false);
        }
        break;
      case 'toggleCommentPanel':
        setShowCommentPanel((open) => !open);
        break;
      default:
        break;
    }
  }, [
    appConfig.enableComments,
    workbook,
    selectedCell,
    contextMenuTarget,
    openCommentEditorForAddress,
    navigateToCommentCell,
    renderer,
  ]);

  const navigateHyperlink = useCallback((target: string) => {
    openHyperlinkTarget(target, workbook, renderer, handleSheetChange, (address) => {
      setSelectedCell(address);
      setSelection({ start: address, end: address });
    });
    renderer?.scheduleRedraw?.();
  }, [workbook, renderer, handleSheetChange]);

  const handleViewModeChange = useCallback((mode: ViewMode) => {
    commandManager.execute(new SetViewModeCommand(workbook, mode));
    setViewMode(mode);
    renderer?.setViewMode(mode);
  }, [workbook, commandManager, renderer]);

  const handleInsertHeaderFooter = useCallback(() => {
    handleViewModeChange('pageLayout');
    setHeaderFooterDialogOpen(true);
  }, [handleViewModeChange]);

  const handleSaveHeaderFooter = useCallback((settings: HeaderFooterSettings) => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;
    commandManager.execute(new SetHeaderFooterCommand(sheet, settings));
    setHeaderFooterDialogOpen(false);
    renderer?.scheduleRedraw?.();
  }, [workbook, commandManager, renderer]);

  const handleInsertWordArt = useCallback(() => {
    setWordArtGalleryOpen(true);
  }, []);

  const handleWordArtPresetSelected = useCallback((presetId: string) => {
    setWordArtGalleryOpen(false);
    handleBeginWordArtInsert(createWordArtTemplate(presetId));
  }, [handleBeginWordArtInsert]);

  const handleAutoSum = useCallback((functionType: string) => {
    if (functionType.toLowerCase() === 'more') return;

    const sheet = workbook.activeSheet;
    if (!sheet || !selection?.start || !selection?.end) return;

    const range = normalizeRange({ start: selection.start, end: selection.end });
    const activeCell = { row: selection.start.row, col: selection.start.col };
    const fn = functionType.toUpperCase() as AutoSumFunction;

    const outputCells = executeAutoSum(commandManager, sheet, range, activeCell, fn);
    if (outputCells.length === 0) return;

    for (const addr of outputCells) {
      renderer?.invalidateRange(addr.row, addr.col, addr.row, addr.col);
    }
    renderer?.redraw?.();

    const target = outputCells[outputCells.length - 1];
    setSelection({ start: target, end: target });
    setSelectedCell(target);
    renderer?.setSelection?.({ start: target, end: target });

    const cell = sheet.getCell(target);
    const value = cell?.value ?? '';
    const formula = cell?.formula ?? '';
    setCellValue(value);
    setCellFormula(formula);
    setFormulaBarValue(formula || String(value ?? ''));
    setHistoryVersion((version) => version + 1);
  }, [workbook, selection, commandManager, renderer]);

  const handleInsertSheet = useCallback(() => {
    const newName = createUniqueSheetName(workbook);
    workbook.addSheet(newName);
    handleSheetChange(newName);
    setSheetRevision((v) => v + 1);
    setSelectedCell({ row: 1, col: 1 });
    setSelection({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    renderer?.setSelection({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    renderer?.scheduleRedraw?.();
  }, [workbook, handleSheetChange, renderer]);

  const handleAddSheet = handleInsertSheet;

  const handleDeleteSheet = useCallback((sheetName: string) => {
    if (workbook.getSheetNames().length <= 1) return;
    try {
      const wasActive = activeSheet === sheetName;
      workbook.deleteSheet(sheetName);
      setSheetRevision((v) => v + 1);
      if (wasActive) {
        const next = workbook.activeSheet?.name;
        if (next) handleSheetChange(next);
      }
      renderer?.scheduleRedraw?.();
    } catch {
      // last sheet guard
    }
  }, [workbook, activeSheet, handleSheetChange, renderer]);

  const handleDeleteActiveSheet = useCallback(() => {
    const name = workbook.activeSheet?.name;
    if (name) handleDeleteSheet(name);
  }, [workbook, handleDeleteSheet]);

  const handleRenameSheet = useCallback((oldName: string, newName: string): string | null => {
    try {
      workbook.renameSheet(oldName, newName);
      setSheetRevision((v) => v + 1);
      if (activeSheet === oldName) {
        setActiveSheet(newName.trim());
      }
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : 'Could not rename sheet.';
    }
  }, [workbook, activeSheet]);

  // Handle zoom change
  const handleZoomChange = useCallback((newZoom: number) => {
    const clampedZoom = Math.max(10, Math.min(400, newZoom));
    setZoom(clampedZoom);
  }, []);

  // Cancel in-cell edit
  const cancelInCellEdit = useCallback(() => {
    setInCellEdit(null);
    setIsPickingReference(false);
  }, []);

  const applyEditedCellValue = useCallback((address: Address, value: string) => {
    const sheet = workbook.activeSheet;
    if (!sheet || !renderer) return;

    if (value.startsWith('=')) {
      sheet.setCellFormula(address, value);
      try {
        sheet.autoRecalculate();
      } catch {
        // Formula engine may not be attached.
      }
    } else {
      sheet.setCellFormula(address, '');
      sheet.setCellValue(address, value);
    }

    renderer.invalidateRange(address.row, address.col, address.row, address.col);
    renderer.redraw?.();
  }, [workbook, renderer]);

  const syncFormulaDraft = useCallback((value: string, cursor: number) => {
    setFormulaBarValue(value);
    setCellFormula(value.startsWith('=') ? value : '');
    setCellValue(value.startsWith('=') ? '' : value);
    setIsPickingReference(isFormulaEditValue(value));
    formulaBarCursorRef.current = cursor;
  }, []);

  const insertFormulaReference = useCallback((start: Address, end: Address) => {
    const pickedRef = formatSelectionAsReference(start, end);
    const pickedIsRange = start.row !== end.row || start.col !== end.col;

    if (inCellEditRef.current) {
      const edit = inCellEditRef.current;
      const cursor = edit.cursorPosition ?? edit.currentValue.length;
      const merged = mergePointReference(edit.currentValue, cursor, pickedRef, pickedIsRange);
      setInCellEdit((prev) =>
        prev
          ? {
              ...prev,
              initialValue: merged.text,
              currentValue: merged.text,
              cursorPosition: merged.cursor,
            }
          : null,
      );
      syncFormulaDraft(merged.text, merged.cursor);
      return;
    }

    if (isEditingFormulaRef.current) {
      const current = formulaBarValueRef.current;
      const cursor = formulaBarCursorRef.current ?? current.length;
      const merged = mergePointReference(current, cursor, pickedRef, pickedIsRange);
      syncFormulaDraft(merged.text, merged.cursor);
    }
  }, [syncFormulaDraft]);

  const handleFormulaBarValueChange = useCallback((value: string) => {
    setFormulaBarValue(value);
    setCellFormula(value.startsWith('=') ? value : '');
    setCellValue(value.startsWith('=') ? '' : value);
    setIsPickingReference(isFormulaEditValue(value));
    if (inCellEditRef.current) {
      setInCellEdit((prev) =>
        prev
          ? {
              ...prev,
              initialValue: value,
              currentValue: value,
              cursorPosition: formulaBarCursorRef.current,
            }
          : null,
      );
    }
  }, []);

  const handleFormulaBarCursorChange = useCallback((position: number) => {
    if (formulaBarCursorRef.current === position) return;
    formulaBarCursorRef.current = position;
    if (inCellEditRef.current && inCellEditRef.current.cursorPosition !== position) {
      setInCellEdit((prev) => (prev ? { ...prev, cursorPosition: position } : null));
    }
  }, []);

  // Handle selection change
  const handleSelectionChange = useCallback((sel: any) => {
    clearDrawingSelection();

    if (
      isPickingReferenceRef.current &&
      (inCellEditRef.current || isEditingFormulaRef.current)
    ) {
      insertFormulaReference(sel.start, sel.end);
      setSelection(sel);
      return;
    }

    if (inCellEdit) {
      const isDifferentCell =
        sel.start.row !== inCellEdit.cell.row || sel.start.col !== inCellEdit.cell.col;
      if (isDifferentCell) {
        applyEditedCellValue(
          { row: inCellEdit.cell.row, col: inCellEdit.cell.col },
          inCellEdit.currentValue,
        );
        setInCellEdit(null);
        inCellEditRef.current = null;
        setIsPickingReference(false);
        isPickingReferenceRef.current = false;
      } else {
        return;
      }
    }

    if (isEditingFormula && !isPickingReferenceRef.current) {
      const formula = formulaBarValueRef.current;
      const editCell = selectedCellRef.current;
      if (editCell && canEditCellValue(appConfig, formula)) {
        applyEditedCellValue({ row: editCell.row, col: editCell.col }, formula);
      }
      setIsEditingFormula(false);
      isEditingFormulaRef.current = false;
    }

    setSelection(sel);
    setSelectedCell({ row: sel.start.row, col: sel.start.col });

    if (!inCellEditRef.current && !isEditingFormulaRef.current) {
      const sheet = workbook.activeSheet;
      if (sheet) {
        const address = { row: sel.start.row, col: sel.start.col };
        const cell = sheet.getCell(address);
        const value = cell?.value || '';
        const formula = cell?.formula || '';
        setCellValue(value);
        setCellFormula(formula);
        setFormulaBarValue(formula || String(value));
      }
    }
  }, [
    workbook,
    isEditingFormula,
    inCellEdit,
    appConfig,
    clearDrawingSelection,
    insertFormulaReference,
    applyEditedCellValue,
  ]);

  const handleFormulaSubmit = useCallback((formula: string) => {
    if (!selectedCell || !canEditCellValue(appConfig, formula)) return;

    applyEditedCellValue({ row: selectedCell.row, col: selectedCell.col }, formula);
    setCellFormula(formula.startsWith('=') ? formula : '');
    setCellValue(formula.startsWith('=') ? '' : formula);
    setFormulaBarValue(formula);
    setIsEditingFormula(false);
    setIsPickingReference(false);
  }, [selectedCell, appConfig, applyEditedCellValue]);

  const handleFormulaEditModeChange = useCallback((editing: boolean) => {
    if (editing && !canStartCellEdit(appConfig)) return;
    setIsEditingFormula(editing);
  }, [appConfig]);

  // Start in-cell editing
  const startInCellEdit = useCallback((cell: { row: number; col: number }, initialValue?: string) => {
    if (!canStartCellEdit(appConfig)) return;
    debugEdit('🚀 startInCellEdit called:', cell, 'renderer:', !!renderer);
    if (!renderer) return;

    const bounds = renderer.getCellBounds({ row: cell.row, col: cell.col });
    debugEdit('📏 Cell bounds:', bounds);
    if (!bounds) return;

    let value: string;
    if (initialValue !== undefined) {
      value = initialValue;
    } else {
      const sheet = workbook.activeSheet;
      const cellData = sheet?.getCell({ row: cell.row, col: cell.col });
      value = cellData?.formula || String(cellData?.value || '');
    }

    setIsEditingFormula(false);
    setInCellEdit({
      cell,
      bounds,
      initialValue: value,
      currentValue: value,
      cursorPosition: value.length,
    });
    syncFormulaDraft(value, value.length);
  }, [renderer, workbook, appConfig, syncFormulaDraft]);

  // Commit in-cell edit
  const commitInCellEdit = useCallback((value: string) => {
    if (!inCellEdit || !canEditCellValue(appConfig, value)) return;

    applyEditedCellValue({ row: inCellEdit.cell.row, col: inCellEdit.cell.col }, value);
    setCellFormula(value.startsWith('=') ? value : '');
    setCellValue(value.startsWith('=') ? '' : value);
    setFormulaBarValue(value);
    setInCellEdit(null);
    setIsPickingReference(false);
  }, [inCellEdit, appConfig, applyEditedCellValue]);

  const cancelInCellEditAndNavigateVertical = useCallback((direction: 'up' | 'down') => {
    const edit = inCellEditRef.current;
    const sheet = workbook.activeSheet;
    if (!edit || !sheet) return;

    const nextCell = {
      row: direction === 'up'
        ? Math.max(1, edit.cell.row - 1)
        : Math.min(sheet.rowCount || 1, edit.cell.row + 1),
      col: edit.cell.col,
    };

    setInCellEdit(null);
    inCellEditRef.current = null;
    setIsPickingReference(false);
    isPickingReferenceRef.current = false;
    setSelectedCell(nextCell);
    setSelection({ start: nextCell, end: nextCell });
    renderer?.setSelection({ start: nextCell, end: nextCell });
    renderer?.scrollToCell?.(nextCell);
  }, [workbook, renderer]);

  const applyInsertDelete = useCallback((
    type: 'insert' | 'delete',
    range: CellRange,
    mode: InsertCellsMode | DeleteCellsMode
  ) => {
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    if (type === 'insert') {
      executeInsertCells(sheet, commandManager, range, mode as InsertCellsMode);
    } else {
      executeDeleteCells(sheet, commandManager, range, mode as DeleteCellsMode);
    }

    const bounds = getInsertDeleteInvalidateBounds(sheet, range, mode);
    renderer?.invalidateRange(bounds.r1, bounds.c1, bounds.r2, bounds.c2);
    renderer?.scheduleRedraw?.();
  }, [workbook, commandManager, renderer]);

  const openInsertDeleteDialog = useCallback((
    type: 'insert' | 'delete',
    range: CellRange | null
  ) => {
    if (!range?.start || !range?.end) return;
    setInsertDeleteDialog({ type, range });
  }, []);

  const pasteAtSelection = useCallback(async (sel: { start: Address; end: Address } | null) => {
    const sheet = workbook.activeSheet;
    if (!sheet || !sel) return false;
    return performPasteAsync(sheet, commandManager, clipboardService, sel.start, {
      renderer,
      onCutComplete: () => setCutRange(null),
    });
  }, [workbook, commandManager, clipboardService, renderer]);

  const copySelection = useCallback((
    sel: { start: Address; end: Address } | null,
    mode: ClipboardCopyMode = 'all'
  ) => {
    const sheet = workbook.activeSheet;
    const range = normalizeSelectionRange(sel);
    if (!sheet || !range) return false;
    performCopy(sheet, clipboardService, range, { mode });
    return true;
  }, [workbook, clipboardService]);

  const cutSelection = useCallback((
    sel: { start: Address; end: Address } | null,
    mode: ClipboardCopyMode = 'all'
  ) => {
    const sheet = workbook.activeSheet;
    const range = normalizeSelectionRange(sel);
    if (!sheet || !range) return false;
    performCut(sheet, clipboardService, range, { mode });
    setCutRange(range);
    return true;
  }, [workbook, clipboardService]);

  // Get context menu items
  const getContextMenuItems = useCallback((): ContextMenuItem[] => {
    const sheet = workbook.activeSheet;
    if (!sheet) return [];

    return [
      {
        id: 'cut',
        label: 'Cut',
        icon: '✂️',
        submenu: [
          {
            id: 'cut-all',
            label: 'Cut',
            shortcut: 'Ctrl+X',
            onClick: () => cutSelection(contextMenuSelectionRef.current, 'all'),
          },
          {
            id: 'cut-values',
            label: 'Cut Values',
            onClick: () => cutSelection(contextMenuSelectionRef.current, 'values'),
          },
        ],
      },
      {
        id: 'copy',
        label: 'Copy',
        icon: '📋',
        submenu: [
          {
            id: 'copy-all',
            label: 'Copy',
            shortcut: 'Ctrl+C',
            onClick: () => copySelection(contextMenuSelectionRef.current, 'all'),
          },
          {
            id: 'copy-values',
            label: 'Copy Values',
            onClick: () => copySelection(contextMenuSelectionRef.current, 'values'),
          },
        ],
      },
      {
        id: 'paste',
        label: 'Paste',
        icon: '📄',
        shortcut: 'Ctrl+V',
        onClick: () => {
          debugMenu('Paste clicked');
          void pasteAtSelection(contextMenuSelectionRef.current);
        },
      },
      { id: 'sep1', label: '', separator: true },
      {
        id: 'insert',
        label: 'Insert...',
        icon: '➕',
        shortcut: 'Ctrl+Shift+=',
        onClick: () => {
          const selection = contextMenuSelectionRef.current;
          debugMenu('Insert clicked');
          if (selection) {
            openInsertDeleteDialog('insert', {
              start: selection.start,
              end: selection.end,
            });
          }
        },
      },
      {
        id: 'delete',
        label: 'Delete...',
        icon: '➖',
        shortcut: 'Ctrl+-',
        onClick: () => {
          const selection = contextMenuSelectionRef.current;
          debugMenu('Delete clicked');
          if (selection) {
            openInsertDeleteDialog('delete', {
              start: selection.start,
              end: selection.end,
            });
          }
        },
      },
      {
        id: 'clearContents',
        label: 'Clear Contents',
        icon: '🧹',
        shortcut: 'Delete',
        onClick: () => {
          const selection = contextMenuSelectionRef.current;
          debugMenu('Clear Contents clicked');
          if (selection && sheet) {
            // Clear cells using command (for undo/redo)
            const clearCmd = new ClearCellsCommand(sheet, { start: selection.start, end: selection.end });
            commandManager.execute(clearCmd);
            
            // Invalidate and redraw
            const r1 = Math.min(selection.start.row, selection.end.row);
            const r2 = Math.max(selection.start.row, selection.end.row);
            const c1 = Math.min(selection.start.col, selection.end.col);
            const c2 = Math.max(selection.start.col, selection.end.col);
            renderer?.invalidateRange(r1, c1, r2, c2);
            renderer?.scheduleRedraw();
          }
        },
      },
      { id: 'sep2', label: '', separator: true },
      {
        id: 'sort',
        label: 'Sort',
        icon: '🔤',
        onClick: () => {
          const selection = contextMenuSelectionRef.current;
          debugMenu('Sort clicked');
        },
      },
      {
        id: 'filter',
        label: 'Filter',
        icon: '🔽',
        onClick: () => {
          const currentSelection = contextMenuSelectionRef.current;
          debugMenu('Filter clicked');
          setIsAutoFilterEnabled(true);
          recomputeHeaderFilterIcons();
          if (currentSelection?.start?.col && renderer) {
            const headerHeight = renderer.optionsReadonly?.headerHeight ?? 24;
            const icon = headerFilterIcons.find((h) => h.col === currentSelection.start.col);
            const x = icon ? Math.round(icon.x + icon.w - 14) : 80;
            openColumnFilterMenu(currentSelection.start.col, x, headerHeight);
          }
        },
      },
      { id: 'sep3', label: '', separator: true },
      {
        id: 'formatCells',
        label: 'Format Cells...',
        icon: '⚙️',
        shortcut: 'Ctrl+1',
        onClick: () => {
          const selection = contextMenuSelectionRef.current;
          debugMenu('Format Cells clicked');
          setIsFormatDialogOpen(true);
        },
      },
      {
        id: 'numberFormat',
        label: 'Number Format...',
        icon: '🔢',
        onClick: () => {
          debugMenu('Number Format clicked');
        },
      },
      { id: 'sep4', label: '', separator: true },
      {
        id: 'hyperlink',
        label: (() => {
          const target = contextMenuTarget();
          if (!target || !sheet) return 'Hyperlink...';
          return sheet.getHyperlink(target)?.target ? 'Edit Hyperlink...' : 'Hyperlink...';
        })(),
        icon: '🔗',
        shortcut: 'Ctrl+K',
        onClick: () => {
          const target = contextMenuTarget();
          if (!target) return;
          openHyperlinkEditorForAddress(target);
        },
      },
      {
        id: 'openHyperlink',
        label: 'Open Hyperlink',
        icon: '🌐',
        disabled: (() => {
          const target = contextMenuTarget();
          if (!target || !sheet) return true;
          return !sheet.getHyperlink(target)?.target;
        })(),
        onClick: () => {
          const target = contextMenuTarget();
          if (!target) return;
          const link = sheet.getHyperlink(target);
          if (!link?.target) return;
          navigateHyperlink(link.target);
        },
      },
      {
        id: 'removeHyperlink',
        label: 'Remove Hyperlink',
        icon: '⛓️‍💥',
        disabled: (() => {
          const target = contextMenuTarget();
          if (!target || !sheet) return true;
          return !sheet.getHyperlink(target)?.target;
        })(),
        onClick: () => {
          const target = contextMenuTarget();
          if (!target) return;
          const link = sheet.getHyperlink(target);
          if (!link?.target) return;
          commandManager.execute(new SetHyperlinkCommand(sheet, target, null));
          renderer?.invalidateRange(target.row, target.col, target.row, target.col);
          renderer?.scheduleRedraw();
        },
      },
      { id: 'sep5', label: '', separator: true },
      {
        id: 'insertComment',
        label: (() => {
          const target = contextMenuTarget();
          if (!target || !sheet) return 'New Comment';
          const comments = sheet.getComments(target);
          return comments.length > 0 ? 'Edit Comment' : 'New Comment';
        })(),
        icon: '💬',
        shortcut: 'Shift+F2',
        onClick: () => {
          debugMenu('New/Edit Comment clicked');
          const target = contextMenuTarget();
          if (!target) return;
          openCommentEditorForAddress(target);
        },
      },
      {
        id: 'deleteComment',
        label: 'Delete Comment',
        icon: '🗑️',
        disabled: (() => {
          const target = contextMenuTarget();
          if (!target || !sheet) return true;
          return sheet.getComments(target).length === 0;
        })(),
        onClick: () => {
          const target = contextMenuTarget();
          if (!target) return;
          const comments = sheet.getComments(target);
          if (comments.length === 0) return;
          const latest = comments[comments.length - 1];
          sheet.deleteComment(target, latest.id, true);
          renderer?.invalidateRange(target.row, target.col, target.row, target.col);
          renderer?.scheduleRedraw();
        },
      },
    ];
  }, [workbook, selection, renderer, clipboardService, commandManager, setIsFormatDialogOpen, openCommentEditorForAddress, openHyperlinkEditorForAddress, navigateHyperlink, contextMenuTarget, recomputeHeaderFilterIcons, headerFilterIcons, openColumnFilterMenu, openInsertDeleteDialog, cutSelection, copySelection, pasteAtSelection]);

  // Handle right-click (context menu)
  useEffect(() => {
    if (!renderer) return;

    const handleContextMenu = (e: MouseEvent) => {
      if (!appConfig.showContextMenu) return;
      e.preventDefault();
      const rendererSelections = renderer.getSelections();
      const frozenSelection = rendererSelections.length > 0 ? rendererSelections[0] : renderer['selection'] || selection;
      contextMenuSelectionRef.current = frozenSelection;

      const hit = renderer.hitTest(e.clientX, e.clientY);
      const clickedCell = hit?.type === 'cell' ? hit.addr : null;
      contextMenuTargetCellRef.current = clickedCell;

      // Move active cell to right-clicked cell without shrinking a multi-cell selection
      if (clickedCell && frozenSelection && isAddressInSelection(clickedCell, frozenSelection)) {
        setSelectedCell(clickedCell);
      }

      setContextMenu({ x: e.clientX, y: e.clientY });
      setContextMenuLayout(null);
      setMiniToolbar({ x: e.clientX, y: e.clientY });
    };

    const canvas = renderer['canvas'];
    if (canvas) {
      canvas.addEventListener('contextmenu', handleContextMenu);
      return () => canvas.removeEventListener('contextmenu', handleContextMenu);
    }
  }, [renderer, selection, isAddressInSelection, appConfig.showContextMenu]);

  // Handle renderer ready
  const handleRendererReady = useCallback((r: CanvasRenderer) => {
    debugRender('=== Renderer Ready ===');
    
    // Expose renderer globally for debugging
    (window as any).__renderer = r;
    
    setRenderer(r);
    
    // Initialize viewport size
    const vp = r.getViewportSize();
    setViewportWidth(vp.width);
    setViewportHeight(vp.height);
    
    // Initialize scroll position
    const scroll = r.getScroll();
    setScrollLeft(scroll.x);
    setScrollTop(scroll.y);
  }, []);

  // Subscribe to renderer scroll events
  useEffect(() => {
    if (!renderer) return;
    
    const onScroll = (event: { x: number; y: number }) => {
      setScrollLeft(event.x);
      setScrollTop(event.y);
    };
    
    const unsubscribe = renderer.onScrollChange(onScroll);
    return () => unsubscribe.dispose();
  }, [renderer]);

  useEffect(() => {
    renderer?.setViewMode(viewMode);
  }, [renderer, viewMode]);

  useEffect(() => {
    recomputeHeaderFilterIcons();
  }, [recomputeHeaderFilterIcons, viewportWidth, viewportHeight, activeSheet, zoom]);

  useEffect(() => {
    if (!filterMenu) return;
    const onDocClick = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (!target.closest('.excel-filter-menu') && !target.closest('.excel-filter-icon')) {
        setFilterMenu(null);
      }
    };
    document.addEventListener('mousedown', onDocClick);
    return () => document.removeEventListener('mousedown', onDocClick);
  }, [filterMenu]);

  // Handle double-click to start in-cell editing
  const isHandlingDoubleClick = useRef(false);
  
  useEffect(() => {
    console.log('🔧 Double-click handler effect running, renderer:', !!renderer, 'workbook:', !!workbook);
    if (!renderer) return;

    const sheet = renderer['sheet'];
    if (!sheet) return;

    const handleCellDoubleClick = (event: any) => {
      // Only handle double-click events (ignore other events)
      if (event.type !== 'cell-double-click') return;
      if (!canStartCellEdit(appConfig)) return;
      console.log('🎯 handleCellDoubleClick received event:', event.type);
      
      // Prevent multiple simultaneous handlers
      if (isHandlingDoubleClick.current) {
        console.log('⏭️  Already handling double-click, skipping');
        return;
      }
      
      isHandlingDoubleClick.current = true;
      
      const { address } = event.event;
      console.log('🖱️ Double-click detected on cell:', address);
      
      // Start edit mode
      const bounds = renderer.getCellBounds({ row: address.row, col: address.col });
      if (!bounds) {
        isHandlingDoubleClick.current = false;
        return;
      }

      // Determine the value to edit
      const sheet = workbook.activeSheet;
      const cellData = sheet?.getCell({ row: address.row, col: address.col });
      const value = cellData?.formula || String(cellData?.value || '');

      // Set edit mode directly to avoid callback changes
      setIsEditingFormula(false);
      console.log('✍️ [DOUBLE-CLICK HANDLER] Activating edit mode for cell', address, 'value:', value);
      setInCellEdit({
        cell: { row: address.row, col: address.col },
        bounds,
        initialValue: value,
        currentValue: value,
        cursorPosition: value.length,
      });
      syncFormulaDraft(value, value.length);
      console.log('✅ Edit mode activated');
      
      // Reset flag after a short delay
      setTimeout(() => {
        isHandlingDoubleClick.current = false;
      }, 100);
    };

    const disposable = sheet.on(handleCellDoubleClick);

    console.log('✅ Double-click listener attached');
    return () => {
      console.log('🧹 Double-click listener cleanup');
      disposable.dispose();
    };
  }, [renderer, workbook, appConfig, syncFormulaDraft]);

  // Comment + hyperlink hover tooltips
  useEffect(() => {
    if (!renderer) return;
    const sheet = renderer['sheet'];
    if (!sheet || typeof sheet.on !== 'function') return;

    const disposable = sheet.on((event: any) => {
      if (event?.type === 'cell-hover') {
        const addr = event.event?.address;
        const bounds = event.event?.bounds;
        if (!addr || !bounds) return;

        const link = sheet.getHyperlink?.(addr);
        if (link?.target) {
          setHyperlinkTooltip({
            x: bounds.x + 4,
            y: Math.max(4, bounds.y - 28),
            text: getDefaultHyperlinkScreenTip(link),
          });
          setStatusBarLink(link.target);
          setCommentTooltip(null);
          return;
        }

        setHyperlinkTooltip(null);
        setStatusBarLink(null);

        const comments = sheet.getComments(addr);
        if (comments.length === 0) {
          setCommentTooltip(null);
          return;
        }
        const latest = comments[comments.length - 1];
        setCommentTooltip({
          x: bounds.x + bounds.width + 8,
          y: bounds.y + 4,
          text: `${latest.author}: ${latest.text}`,
        });
      } else if (event?.type === 'cell-hover-end') {
        setCommentTooltip(null);
        setHyperlinkTooltip(null);
        setStatusBarLink(null);
      }
    });

    return () => {
      disposable.dispose();
    };
  }, [renderer]);

  // Global keyboard shortcut handler
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Log EVERY key event for debugging
      console.log('⌨️ [ExcelApp] Key event', {
        key: e.key,
        code: e.code,
        ctrl: e.ctrlKey,
        meta: e.metaKey,
        shift: e.shiftKey,
        alt: e.altKey,
        target: (e.target as HTMLElement)?.tagName,
        isComposing: e.isComposing,
      });
      
      // Get current values from refs to avoid re-attaching listeners on state changes
      const selectedCell = selectedCellRef.current;
      const selection = selectionRef.current;
      const inCellEdit = inCellEditRef.current;
      const isEditingFormula = isEditingFormulaRef.current;
      
      const target = e.target as HTMLElement;
      const isInInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable;
      const isFindReplaceOpen = isFindReplaceOpenRef.current;

      // When Find/Replace is open, only allow dialog input and Ctrl+F/Ctrl+H reopen shortcuts
      if (isFindReplaceOpen) {
        if (isInInput || target.closest('.find-replace-dialog')) return;
        if (isCtrlLetter(e, 'f') || isCtrlLetter(e, 'h')) {
          // fall through to open/switch tab
        } else {
          return;
        }
      }
      
      // ===== FILE MENU SHORTCUTS (ALWAYS ACTIVE) =====
      
      // Ctrl+N (New workbook)
      if (isCtrlLetter(e, 'n')) {
        if (!appConfig.allowOpen) return;
        e.preventDefault();
        setBackstagePanel('new');
        setBackstageOpen(true);
        return;
      }
      
      // Ctrl+O (Open workbook)
      if (isCtrlLetter(e, 'o')) {
        if (!appConfig.allowOpen) return;
        e.preventDefault();
        setBackstagePanel('open');
        setBackstageOpen(true);
        return;
      }
      
      // Ctrl+S (Save workbook)
      if (isCtrlLetter(e, 's')) {
        if (!appConfig.allowSave && !onSave) return;
        e.preventDefault();
        if (onSave) {
          onSave();
        } else if (appConfig.allowSave || appConfig.allowExport) {
          // If no save handler, open export panel
          setBackstagePanel('export');
          setBackstageOpen(true);
        }
        return;
      }

      // Ctrl+F (Find) - works everywhere except in input fields
      if (isCtrlLetter(e, 'f') && !isInInput) {
        e.preventDefault();
        setFindReplaceTab('find');
        setIsFindReplaceOpen(true);
        return;
      }
      
      // Ctrl+H (Replace) - works everywhere except in input fields
      if (isCtrlLetter(e, 'h') && !isInInput) {
        e.preventDefault();
        setFindReplaceTab('replace');
        setIsFindReplaceOpen(true);
        return;
      }

      // Shift+F11 (Insert Sheet)
      if (e.code === 'F11' && e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && !isInInput) {
        e.preventDefault();
        handleInsertSheet();
        return;
      }
      
      // Escape - cancel in-cell editing
      if (e.code === 'Escape' && inCellEdit) {
        e.preventDefault();
        setInCellEdit(null);
        inCellEditRef.current = null;
        setIsPickingReference(false);
        isPickingReferenceRef.current = false;
        return;
      }
      
      // Enter - move down (works when not editing)
      if (e.code === 'Enter' && !e.shiftKey && !isEditingFormula && !inCellEdit && !isInInput && selectedCell) {
        e.preventDefault();
        const newRow = Math.min(selectedCell.row + 1, workbook.activeSheet?.rowCount || 1000);
        setSelectedCell({ row: newRow, col: selectedCell.col });
        renderer?.setSelection({ 
          start: { row: newRow, col: selectedCell.col }, 
          end: { row: newRow, col: selectedCell.col } 
        });
        return;
      }
      
      // Shift+Enter - move up (works when not editing)
      if (e.code === 'Enter' && e.shiftKey && !isEditingFormula && !inCellEdit && !isInInput && selectedCell) {
        e.preventDefault();
        const newRow = Math.max(selectedCell.row - 1, 0);
        setSelectedCell({ row: newRow, col: selectedCell.col });
        renderer?.setSelection({ 
          start: { row: newRow, col: selectedCell.col }, 
          end: { row: newRow, col: selectedCell.col } 
        });
        return;
      }
      
      // Tab - move right (works when not editing)
      if (e.code === 'Tab' && !e.shiftKey && !isEditingFormula && !inCellEdit && !isInInput && selectedCell) {
        e.preventDefault();
        const newCol = Math.min(selectedCell.col + 1, workbook.activeSheet?.colCount || 100);
        setSelectedCell({ row: selectedCell.row, col: newCol });
        renderer?.setSelection({ 
          start: { row: selectedCell.row, col: newCol }, 
          end: { row: selectedCell.row, col: newCol } 
        });
        return;
      }
      
      // Shift+Tab - move left (works when not editing)
      if (e.code === 'Tab' && e.shiftKey && !isEditingFormula && !inCellEdit && !isInInput && selectedCell) {
        e.preventDefault();
        const newCol = Math.max(selectedCell.col - 1, 0);
        setSelectedCell({ row: selectedCell.row, col: newCol });
        renderer?.setSelection({ 
          start: { row: selectedCell.row, col: newCol }, 
          end: { row: selectedCell.row, col: newCol } 
        });
        return;
      }
      
      // ===== SHORTCUTS THAT ONLY WORK ON SHEET CANVAS =====
      
      // Skip if editing in formula bar or in input field
      if (isEditingFormula || inCellEdit || isInInput) return;
      
      const sheet = workbook.activeSheet;
      if (!sheet) return;
      
      // --- Clipboard Operations ---
      
      // Ctrl+C (Copy)
      if (isCtrlLetter(e, 'c')) {
        e.preventDefault();
        copySelection(selection);
        return;
      }
      
      // Ctrl+X (Cut)
      if (isCtrlLetter(e, 'x')) {
        e.preventDefault();
        cutSelection(selection);
        return;
      }
      
      // Ctrl+V (Paste)
      if (isCtrlLetter(e, 'v')) {
        e.preventDefault();
        void pasteAtSelection(selection);
        return;
      }
      
      // --- Formatting Operations ---
      
      // Ctrl+B (Bold)
      if (isCtrlLetter(e, 'b')) {
        e.preventDefault();
        if (formattingController && selectionAddresses.length > 0) {
          formattingController.toggleBold(selectionAddresses);
        }
        return;
      }
      
      // Ctrl+I (Italic)
      if (isCtrlLetter(e, 'i')) {
        e.preventDefault();
        if (formattingController && selectionAddresses.length > 0) {
          formattingController.toggleItalic(selectionAddresses);
        }
        return;
      }
      
      // Ctrl+U (Underline)
      if (isCtrlLetter(e, 'u')) {
        e.preventDefault();
        if (formattingController && selectionAddresses.length > 0) {
          formattingController.toggleUnderline(selectionAddresses);
        }
        return;
      }
      
      // Ctrl+S (Save)
      if (isCtrlLetter(e, 's')) {
        e.preventDefault();
        if (onSave) onSave();
        return;
      }
      
      // Ctrl+1 (Format Cells dialog)
      if (isCtrlPhysicalKey(e, 'Digit1')) {
        e.preventDefault();
        setIsFormatDialogOpen(true);
        return;
      }

      // Alt+= (AutoSum)
      if (
        !isInInput &&
        e.altKey &&
        !e.ctrlKey &&
        !e.metaKey &&
        !e.shiftKey &&
        selection?.start &&
        selection?.end &&
        (e.code === 'Equal' || e.code === 'NumpadAdd')
      ) {
        e.preventDefault();
        handleAutoSum('SUM');
        return;
      }

      // Ctrl+Shift+= / Ctrl+Shift++ (Insert cells dialog)
      if (
        !isInInput &&
        selection?.start &&
        selection?.end &&
        hasCtrlOrMeta(e) &&
        e.shiftKey &&
        (e.code === 'Equal' || e.code === 'NumpadAdd')
      ) {
        e.preventDefault();
        openInsertDeleteDialog('insert', { start: selection.start, end: selection.end });
        return;
      }

      // Ctrl+- (Delete cells dialog)
      if (
        !isInInput &&
        selection?.start &&
        selection?.end &&
        hasCtrlOrMeta(e) &&
        !e.shiftKey &&
        (e.code === 'Minus' || e.code === 'NumpadSubtract')
      ) {
        e.preventDefault();
        openInsertDeleteDialog('delete', { start: selection.start, end: selection.end });
        return;
      }

      // Ctrl+Shift+L (toggle AutoFilter)
      if (isCtrlPhysicalKey(e, 'KeyL', { shift: true })) {
        e.preventDefault();
        setIsAutoFilterEnabled((prev) => {
          const next = !prev;
          if (!next) {
            executeFilterStateCommand(
              (activeSheet) => activeSheet.clearAllFilters?.(),
              'Clear All Filters'
            );
            setFilterMenu(null);
          } else {
            setTimeout(() => recomputeHeaderFilterIcons(), 0);
          }
          return next;
        });
        return;
      }
      
      // --- Navigation Operations ---
      
      // Ctrl+Home (Jump to A1)
      if (hasCtrlOrMeta(e) && e.code === 'Home') {
        e.preventDefault();
        setSelectedCell({ row: 0, col: 0 });
        renderer?.setSelection({ 
          start: { row: 0, col: 0 }, 
          end: { row: 0, col: 0 } 
        });
        // Scroll to top-left
        renderer?.setScroll(0, 0);
        return;
      }
      
      // Ctrl+End (Jump to last used cell)
      if (hasCtrlOrMeta(e) && e.code === 'End') {
        e.preventDefault();
        // Find last used cell using sparse cell iteration (O(n) not O(rows*cols))
        let lastRow = 0;
        let lastCol = 0;
        sheet.forEachNonEmptyCell((row: number, col: number) => {
          lastRow = Math.max(lastRow, row);
          lastCol = Math.max(lastCol, col);
        });
        setSelectedCell({ row: lastRow, col: lastCol });
        renderer?.setSelection({ 
          start: { row: lastRow, col: lastCol }, 
          end: { row: lastRow, col: lastCol } 
        });
        return;
      }
      
      // Ctrl+A (Select used data range)
      if (isCtrlLetter(e, 'a')) {
        e.preventDefault();

        const usedRange = getUsedDataRange(sheet);

        console.log('📋 [ExcelApp] Ctrl+A - Selecting used data range:', usedRange);

        setSelectedCell({ row: usedRange.start.row, col: usedRange.start.col });
        setSelection(usedRange);
        renderer?.setSelection(usedRange);
        return;
      }
      
      // Alt+ArrowDown (Show validation dropdown for list validation)
      if (e.altKey && e.code === 'ArrowDown' && !isInInput && selectedCell) {
        e.preventDefault();
        const validationEngine = (renderer as any)?.getValidationEngine?.();
        if (validationEngine) {
          const items = validationEngine.getDropdownItems(selectedCell);
          if (items && items.length > 0) {
            // Get cell bounds for positioning dropdown
            const cellBounds = (renderer as any)?.rectForCell?.(selectedCell.row, selectedCell.col);
            if (cellBounds) {
              const currentValue = sheet.getCellValue(selectedCell)?.toString() || '';
              setValidationDropdown({
                items,
                cellAddress: selectedCell,
                cellBounds,
                currentValue,
              });
            }
          }
        }
        return;
      }
      
      // Ctrl+Arrow keys (Jump to edge of data region)
      if (hasCtrlOrMeta(e) && isArrowKey(e) && !e.shiftKey) {
        e.preventDefault();
        if (!selectedCell) return;
        
        const lastRow = sheet.rowCount || 1000;
        const lastCol = sheet.colCount || 100;
        
        // Helper to check if cell is empty
        const isEmpty = (r: number, c: number): boolean => {
          if (r < 0 || c < 0 || r >= lastRow || c >= lastCol) return true;
          const value = sheet.getCellValue({ row: r, col: c });
          return value === null || value === undefined || value === '';
        };
        
        let newRow = selectedCell.row;
        let newCol = selectedCell.col;
        const currentIsEmpty = isEmpty(newRow, newCol);
        
        switch (e.code) {
          case 'ArrowUp':
            if (currentIsEmpty) {
              // Jump to first non-empty cell above
              while (newRow > 0 && isEmpty(newRow - 1, newCol)) {
                newRow--;
              }
              // Now we're at the last empty cell before data, move into data
              if (newRow > 0) newRow--;
            } else {
              // Jump to top edge of current data region
              while (newRow > 0 && !isEmpty(newRow - 1, newCol)) {
                newRow--;
              }
            }
            break;
            
          case 'ArrowDown':
            if (currentIsEmpty) {
              // Jump to first non-empty cell below
              while (newRow < lastRow - 1 && isEmpty(newRow + 1, newCol)) {
                newRow++;
              }
              // Now we're at the last empty cell before data, move into data
              if (newRow < lastRow - 1) newRow++;
            } else {
              // Jump to bottom edge of current data region
              while (newRow < lastRow - 1 && !isEmpty(newRow + 1, newCol)) {
                newRow++;
              }
            }
            break;
            
          case 'ArrowLeft':
            if (currentIsEmpty) {
              // Jump to first non-empty cell to the left
              while (newCol > 0 && isEmpty(newRow, newCol - 1)) {
                newCol--;
              }
              // Now we're at the last empty cell before data, move into data
              if (newCol > 0) newCol--;
            } else {
              // Jump to left edge of current data region
              while (newCol > 0 && !isEmpty(newRow, newCol - 1)) {
                newCol--;
              }
            }
            break;
            
          case 'ArrowRight':
            if (currentIsEmpty) {
              // Jump to first non-empty cell to the right
              while (newCol < lastCol - 1 && isEmpty(newRow, newCol + 1)) {
                newCol++;
              }
              // Now we're at the last empty cell before data, move into data
              if (newCol < lastCol - 1) newCol++;
            } else {
              // Jump to right edge of current data region
              while (newCol < lastCol - 1 && !isEmpty(newRow, newCol + 1)) {
                newCol++;
              }
            }
            break;
        }
        
        setSelectedCell({ row: newRow, col: newCol });
        renderer?.setSelection({ 
          start: { row: newRow, col: newCol }, 
          end: { row: newRow, col: newCol } 
        });
        return;
      }
      
      // Shift+Arrow keys (Extend selection)
      if (e.shiftKey && isArrowKey(e)) {
        e.preventDefault();
        if (!selection) return;
        
        const currentEnd = selection.end;
        let newRow = currentEnd.row;
        let newCol = currentEnd.col;
        
        switch (e.code) {
          case 'ArrowUp':
            newRow = Math.max(0, currentEnd.row - 1);
            break;
          case 'ArrowDown':
            newRow = Math.min((sheet.rowCount || 1000) - 1, currentEnd.row + 1);
            break;
          case 'ArrowLeft':
            newCol = Math.max(0, currentEnd.col - 1);
            break;
          case 'ArrowRight':
            newCol = Math.min((sheet.colCount || 100) - 1, currentEnd.col + 1);
            break;
        }
        
        const newSelection = {
          start: selection.start,
          end: { row: newRow, col: newCol }
        };
        
        setSelection(newSelection);
        renderer?.setSelection(newSelection);
        return;
      }
      
      // --- Editing Operations ---
      
      // Enter — follow hyperlink when cell has a link (Excel behavior)
      if (e.code === 'Enter' && !e.ctrlKey && !e.altKey && !e.shiftKey && selectedCell && !isInInput && !inCellEdit) {
        const link = sheet.getHyperlink?.(selectedCell);
        if (link?.target) {
          e.preventDefault();
          navigateHyperlink(link.target);
          return;
        }
      }

      // Ctrl+K (Insert/Edit Hyperlink)
      if (isCtrlLetter(e, 'k') && selectedCell) {
        e.preventDefault();
        openHyperlinkEditorForAddress(selectedCell);
        return;
      }

      // F2 to start editing with existing content
      if (e.code === 'F2' && selectedCell && renderer) {
        if (e.shiftKey) {
          if (!appConfig.enableComments) return;
          e.preventDefault();
          openCommentEditorForAddress(selectedCell);
          return;
        }
        if (!canStartCellEdit(appConfig)) return;
        e.preventDefault();
        
        // Inline edit mode activation
        const bounds = renderer.getCellBounds(selectedCell);
        if (bounds) {
          const cellData = sheet?.getCell(selectedCell);
          const value = cellData?.formula || String(cellData?.value || '');
          
          console.log('✍️ [F2 KEY] Activating edit mode for cell', selectedCell, 'value:', value);
          setIsEditingFormula(false);
          setInCellEdit({
            cell: selectedCell,
            bounds,
            initialValue: value,
            currentValue: value,
            cursorPosition: value.length,
          });
          syncFormulaDraft(value, value.length);
        }
        return;
      }
      
      // Delete or Backspace to clear cell
      if ((e.code === 'Delete' || e.code === 'Backspace') && selection) {
        e.preventDefault();
        
        // Clear cells using command (for undo/redo)
        const clearCmd = new ClearCellsCommand(sheet, { start: selection.start, end: selection.end });
        commandManager.execute(clearCmd);
        
        // Invalidate and redraw
        const r1 = Math.min(selection.start.row, selection.end.row);
        const r2 = Math.max(selection.start.row, selection.end.row);
        const c1 = Math.min(selection.start.col, selection.end.col);
        const c2 = Math.max(selection.start.col, selection.end.col);
        renderer?.invalidateRange(r1, c1, r2, c2);
        renderer?.scheduleRedraw();
        return;
      }
      
      // Any printable character starts editing with that character
      const typedCharacter = getTypedCharacter(e);
      
      if (selectedCell && typedCharacter && renderer) {
        if (!canStartCellEdit(appConfig)) return;
        if (typedCharacter === '=' && appConfig.allowFormulaEdit === false) return;
        e.preventDefault();
        
        // Inline edit mode activation with initial character
        const bounds = renderer.getCellBounds(selectedCell);
        if (bounds) {
          console.log('✍️ [TYPING KEY] Activating edit mode for cell', selectedCell, 'key:', typedCharacter);
          setIsEditingFormula(false);
          setInCellEdit({
            cell: selectedCell,
            bounds,
            initialValue: typedCharacter,
            currentValue: typedCharacter,
            cursorPosition: typedCharacter.length,
          });
          syncFormulaDraft(typedCharacter, typedCharacter.length);
        }
      }
    };

    console.log('🎹 Keyboard handler effect running - attaching listener');
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      console.log('🧹 Keyboard handler effect cleanup - removing listener');
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [workbook, renderer, clipboardService, commandManager, onSave, openCommentEditorForAddress, openHyperlinkEditorForAddress, navigateHyperlink, recomputeHeaderFilterIcons, executeFilterStateCommand, formattingController, selectionAddresses, openInsertDeleteDialog, copySelection, cutSelection, pasteAtSelection, handleInsertSheet, handleAutoSum, appConfig, syncFormulaDraft]);

  // Form Control dialog — cell link range picker
  useEffect(() => {
    if (!renderer || !formControlCellLinkPicker) return;

    const sheet = renderer['sheet'];
    if (!sheet || typeof sheet.on !== 'function') return;

    const handleCellClick = (event: { type: string; event?: { address?: Address } }) => {
      if (event.type !== 'cell-click') return;
      const { address } = event.event ?? {};
      if (!address) return;

      setPickedFormControlCellLink(formatAddressA1(address));
      setFormControlCellLinkPicker(false);
      setFormatControlObjectId((id) => id);
    };

    const subscription = sheet.on(handleCellClick);
    return () => subscription.dispose();
  }, [renderer, formControlCellLinkPicker]);

  // Ctrl+Click on a hyperlinked cell opens the link (Excel behavior)
  useEffect(() => {
    if (!renderer) return;

    const sheet = renderer['sheet'];
    if (!sheet || typeof sheet.on !== 'function') return;

    const handleHyperlinkClick = (event: any) => {
      if (event.type !== 'cell-click') return;
      const { address, originalEvent } = event.event ?? {};
      if (!address || !originalEvent) return;
      if (!originalEvent.ctrlKey && !originalEvent.metaKey) return;

      const link = sheet.getHyperlink?.(address);
      if (!link?.target) return;

      originalEvent.preventDefault();
      navigateHyperlink(link.target);
    };

    const subscription = sheet.on(handleHyperlinkClick);

    return () => {
      subscription.dispose();
    };
  }, [renderer, navigateHyperlink]);

  const handleCommentDialogSave = useCallback((text: string) => {
    if (!commentDialog) return;
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    const trimmed = text.trim();
    const existing = sheet.getComments(commentDialog.address);
    const latest = existing.length > 0 ? existing[existing.length - 1] : null;

    if (trimmed.length === 0) {
      if (latest) {
        sheet.deleteComment(commentDialog.address, latest.id, true);
      }
      setCommentDialog(null);
      renderer?.invalidateRange(commentDialog.address.row, commentDialog.address.col, commentDialog.address.row, commentDialog.address.col);
      renderer?.scheduleRedraw();
      return;
    }

    if (latest) {
      sheet.updateComment(commentDialog.address, latest.id, { text: trimmed });
    } else {
      sheet.addComment(commentDialog.address, {
        text: trimmed,
        author: appConfig.authorName,
        authorAvatar: appConfig.authorAvatar,
      });
    }

    setCommentDialog(null);
    renderer?.invalidateRange(commentDialog.address.row, commentDialog.address.col, commentDialog.address.row, commentDialog.address.col);
    renderer?.scheduleRedraw();
  }, [commentDialog, workbook, renderer, appConfig.authorName, appConfig.authorAvatar]);

  const handleHyperlinkDialogSave = useCallback((result: InsertHyperlinkDialogResult) => {
    if (!hyperlinkDialog) return;
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    commandManager.execute(new SetHyperlinkCommand(
      sheet,
      hyperlinkDialog.address,
      result.hyperlink,
      { displayText: result.displayText }
    ));

    if (result.editNewDocumentNow && result.hyperlink.newDocumentName) {
      downloadBlankWorkbook(result.hyperlink.newDocumentName);
    }

    const { row, col } = hyperlinkDialog.address;
    setHyperlinkDialog(null);
    renderer?.invalidateRange(row, col, row, col);
    renderer?.scheduleRedraw();
  }, [hyperlinkDialog, workbook, commandManager, renderer]);

  const handleHyperlinkDialogRemove = useCallback(() => {
    if (!hyperlinkDialog) return;
    const sheet = workbook.activeSheet;
    if (!sheet) return;

    commandManager.execute(new SetHyperlinkCommand(sheet, hyperlinkDialog.address, null));

    const { row, col } = hyperlinkDialog.address;
    setHyperlinkDialog(null);
    renderer?.invalidateRange(row, col, row, col);
    renderer?.scheduleRedraw();
  }, [hyperlinkDialog, workbook, commandManager, renderer]);

  const handleCommentDialogDelete = useCallback(() => {
    if (!commentDialog) return;
    const sheet = workbook.activeSheet;
    if (!sheet) return;
    const existing = sheet.getComments(commentDialog.address);
    const latest = existing.length > 0 ? existing[existing.length - 1] : null;
    if (!latest) {
      setCommentDialog(null);
      return;
    }
    sheet.deleteComment(commentDialog.address, latest.id, true);
    setCommentDialog(null);
    renderer?.invalidateRange(commentDialog.address.row, commentDialog.address.col, commentDialog.address.row, commentDialog.address.col);
    renderer?.scheduleRedraw();
  }, [commentDialog, workbook, renderer]);

  return (
    <div className="excel-app" style={style}>
      {/* Title Bar */}
      <TitleBar
        fileName={`${fileName} - Excel`}
        onSave={onSave}
        onUndo={() => commandManager.undo()}
        onRedo={() => commandManager.redo()}
      />

      {/* Ribbon Tabs */}
      {appConfig.showRibbon && (
        <>
      <RibbonTabs
        activeTab={activeTab}
        onTabChange={setActiveTab}
        onFileClick={() => {
          if (!appConfig.allowOpen) return;
          setBackstagePanel('new');
          setBackstageOpen(true);
        }}
      />

      {/* Ribbon Content */}
      <Ribbon
        commandManager={commandManager}
        selection={selection}
        activeTab={activeTab}
        drawingLayer={drawingLayer}
        workbook={workbook}
        clipboardService={clipboardService}
        onAfterCommand={() => renderer?.redraw?.()}
        historyVersion={historyVersion}
        viewMode={viewMode}
        onViewModeChange={handleViewModeChange}
        zoom={zoom}
        onZoomChange={handleZoomChange}
        onInsertHyperlink={() => {
          if (selectedCell) {
            openHyperlinkEditorForAddress(selectedCell);
          }
        }}
        onOpenFindReplace={(tab) => {
          setFindReplaceTab(tab);
          setIsFindReplaceOpen(true);
        }}
        onInsertSheet={handleInsertSheet}
        onDeleteSheet={handleDeleteActiveSheet}
        onRequestRenameSheet={() => setRenameSheetTrigger((v) => v + 1)}
        onFilter={handleRibbonFilter}
        onDataCommand={handleDataCommand}
        selectedCells={selectionAddresses}
        onInsertPivotTable={handleInsertPivotTable}
        onInsertTable={handleInsertTable}
        formattingController={formattingController}
        onReviewCommand={handleReviewCommand}
        worksheet={workbook.activeSheet}
        onDrawingChange={handleDrawingChange}
        onBeginShapeInsert={handleBeginShapeInsert}
        onBeginPictureInsert={handleBeginPictureInsert}
        onBeginIconInsert={handleBeginIconInsert}
        onBeginFormControlInsert={handleBeginFormControlInsert}
        onBeginTextBoxInsert={handleBeginTextBoxInsert}
        onInsertHeaderFooter={handleInsertHeaderFooter}
        onInsertWordArt={handleInsertWordArt}
        onAutoSum={handleAutoSum}
      />
        </>
      )}

      {/* Formula Bar */}
      {appConfig.showFormulaBar && (
        <FormulaBar
          selectedCell={selectedCell}
          cellValue={cellValue}
          cellFormula={cellFormula}
          onFormulaSubmit={handleFormulaSubmit}
          onValueChange={handleFormulaBarValueChange}
          onCursorChange={handleFormulaBarCursorChange}
          isEditing={isEditingFormula}
          onEditModeChange={handleFormulaEditModeChange}
          onReferencePickingChange={setIsPickingReference}
          isPickingReference={isPickingReference}
          functionRegistry={formulaEngine.functions}
        />
      )}

      {/* Spreadsheet Area */}
      <div
        className="spreadsheet-area"
        style={{ position: 'relative', display: 'flex', minHeight: 0 }}
        onDragOver={handleSpreadsheetDragOver}
        onDrop={handleSpreadsheetDrop}
      >
        <div style={{ position: 'relative', flex: 1, minWidth: 0 }}>
        <CyberSheet
          workbook={workbook}
          sheetName={activeSheet}
          commandManager={commandManager}
          enableCustomComponents={appConfig.enableCustomComponents}
          zoom={zoom / 100}
          viewMode={viewMode}
          fontFamily={cyberSheetConfig.fonts.defaultFamily}
          fontSize={cyberSheetConfig.fonts.defaultSize}
          rendererOptions={{
            onRequestColumnFilterMenu: ({ col, anchor, values, apply, clear }: any) => {
              if (!isAutoFilterEnabled) return;
              const existing = renderer?.sheetReadonly?.getColumnFilter?.(col);
              const allValues = values.map((v: { value: string; count: number }) => v.value);
              let selected: Set<string>;
              if (existing && existing.type === 'in' && Array.isArray(existing.value)) {
                selected = new Set((existing.value as any[]).map(String));
              } else {
                const present = new Set<string>();
                try {
                  const visibleRows: number[] = renderer?.sheetReadonly?.getVisibleRowIndices?.() || [];
                  for (const rr of visibleRows) {
                    const v = renderer?.sheetReadonly?.getCellValue({ row: rr, col });
                    present.add(v == null ? '' : String(v));
                  }
                } catch {}
                selected = present.size ? present : new Set(allValues);
              }
              setFilterSel(selected);
              setFilterSearch('');
              setFilterMenu({
                col,
                x: anchor.x,
                y: anchor.y,
                values,
                apply: (sel) => {
                  executeFilterStateCommand(
                    (sheet) => {
                      if (sel.length === allValues.length) clear();
                      else apply(sel);
                    },
                    `Filter Column ${col}`
                  );
                  setFilterMenu(null);
                },
                clear: () => {
                  executeFilterStateCommand(
                    () => clear(),
                    `Clear Filter Column ${col}`
                  );
                  setFilterMenu(null);
                },
              });
            },
          }}
          onRendererReady={handleRendererReady}
          onSelectionChange={handleSelectionChange}
          style={{ width: '100%', height: '100%' }}
        />
        <DrawingCanvas
          ref={drawingCanvasRef}
          drawingLayer={drawingLayer}
          canvasWidth={viewportWidth}
          canvasHeight={viewportHeight}
          scrollLeft={scrollLeft}
          scrollTop={scrollTop}
          zoom={zoom / 100}
          worksheet={workbook.activeSheet}
          commandManager={commandManager}
          onObjectChange={handleDrawingChange}
          onFormControlContextMenu={(objectId, x, y) => {
            setFormControlContextMenu({ objectId, x, y });
          }}
          suspendInteraction={formControlCellLinkPicker}
        />

        {/* Cut Range Visual Indication */}
        {cutRange && renderer && (
          <CutRangeOverlay
            cutRange={cutRange}
            renderer={renderer}
            zoom={zoom / 100}
          />
        )}

        {/* In-Cell Edit Overlay */}
        {inCellEdit && (
          <div 
            style={{ 
              position: 'absolute', 
              top: 0, 
              left: 0, 
              width: '100%', 
              height: '100%', 
              // When picking references, allow clicks to pass through to canvas
              // When not picking, catch clicks to commit the edit
              pointerEvents: isPickingReference ? 'none' : 'auto',
              zIndex: 1000,
              background: 'transparent'
            }}
            onMouseDown={(e) => {
              // If clicking outside the input, commit and prevent event propagation
              if (!(e.target as HTMLElement).matches('input')) {
                e.preventDefault();
                e.stopPropagation();
                // Get the current value from the input before committing
                const inputElement = e.currentTarget.querySelector('input');
                const currentValue = inputElement?.value || inCellEdit.currentValue;
                commitInCellEdit(currentValue);
              }
            }}
          >
            <CellEditOverlay
              cell={inCellEdit.cell}
              value={inCellEdit.currentValue}
              bounds={inCellEdit.bounds}
              cursorPosition={inCellEdit.cursorPosition}
              onCommit={commitInCellEdit}
              onCancel={cancelInCellEdit}
              onValueChange={(newValue) => {
                const cursor = formulaBarCursorRef.current;
                setInCellEdit((prev) =>
                  prev
                    ? {
                        ...prev,
                        currentValue: newValue,
                        initialValue: newValue,
                        cursorPosition: cursor,
                      }
                    : null,
                );
                syncFormulaDraft(newValue, cursor);
              }}
              onCursorChange={(position) => {
                if (formulaBarCursorRef.current === position) return;
                formulaBarCursorRef.current = position;
                setInCellEdit((prev) =>
                  prev && prev.cursorPosition !== position
                    ? { ...prev, cursorPosition: position }
                    : prev,
                );
              }}
              onNavigateVertical={cancelInCellEditAndNavigateVertical}
              isPickingReference={isPickingReference}
              onReferencePickingChange={setIsPickingReference}
              fontSize={cyberSheetConfig.fonts.defaultSize}
            />
          </div>
        )}

        {/* Comment hover preview */}
        {commentTooltip && (
          <div
            style={{
              position: 'absolute',
              left: commentTooltip.x,
              top: commentTooltip.y,
              background: '#fffbe6',
              border: '1px solid #d9c27a',
              borderRadius: 4,
              boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
              padding: '6px 8px',
              maxWidth: 280,
              fontSize: 12,
              lineHeight: 1.35,
              zIndex: 10003,
              pointerEvents: 'none',
              whiteSpace: 'pre-wrap',
            }}
          >
            {commentTooltip.text}
          </div>
        )}

        {hyperlinkTooltip && (
          <div
            style={{
              position: 'absolute',
              left: hyperlinkTooltip.x,
              top: hyperlinkTooltip.y,
              background: '#ffffe1',
              border: '1px solid #aca899',
              borderRadius: 0,
              boxShadow: '1px 1px 3px rgba(0,0,0,0.15)',
              padding: '4px 8px',
              maxWidth: 360,
              fontSize: 11,
              fontFamily: 'Segoe UI, Arial, sans-serif',
              lineHeight: 1.3,
              zIndex: 10004,
              pointerEvents: 'none',
              whiteSpace: 'nowrap',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
            }}
          >
            {hyperlinkTooltip.text}
          </div>
        )}

        {/* AutoFilter header icons */}
        {isAutoFilterEnabled && renderer && headerFilterIcons.map((h) => {
          const headerHeight = renderer.optionsReadonly?.headerHeight ?? 24;
          const iconLeft = Math.round(h.x + h.w - 14);
          const iconTop = Math.round(headerHeight / 2 - 6);
          const isActive = !!renderer.sheetReadonly?.getColumnFilter?.(h.col);
          return (
            <div
              key={`filter-col-${h.col}`}
              className="excel-filter-icon"
              onClick={(e) => {
                e.stopPropagation();
                openColumnFilterMenu(h.col, iconLeft, headerHeight);
              }}
              title={`Filter column ${h.col}`}
              style={{
                position: 'absolute',
                left: iconLeft,
                top: iconTop,
                zIndex: 1001,
                width: 12,
                height: 12,
                borderRadius: 2,
                fontSize: 10,
                lineHeight: '12px',
                textAlign: 'center',
                cursor: 'pointer',
                background: isActive ? '#e6f2ff' : 'transparent',
                color: isActive ? '#0f6cbd' : '#666',
              }}
            >
              ▼
            </div>
          );
        })}

        {/* AutoFilter menu */}
        {filterMenu && (
          <div
            className="excel-filter-menu"
            style={{
              position: 'absolute',
              left: filterMenu.x,
              top: filterMenu.y,
              zIndex: 1002,
              width: 260,
              background: '#fff',
              border: '1px solid #c8c8c8',
              borderRadius: 6,
              boxShadow: '0 8px 22px rgba(0,0,0,0.18)',
              overflow: 'hidden',
              fontFamily: 'Segoe UI, Arial, sans-serif',
              fontSize: 12,
            }}
          >
            <div style={{ padding: '8px 10px', borderBottom: '1px solid #eee', fontWeight: 600 }}>
              Filter - Column {filterMenu.col}
            </div>
            <div style={{ padding: 8 }}>
              <input
                type="text"
                placeholder="Search..."
                value={filterSearch}
                onChange={(e) => setFilterSearch(e.target.value)}
                style={{
                  width: '100%',
                  boxSizing: 'border-box',
                  border: '1px solid #c8c8c8',
                  borderRadius: 4,
                  padding: '6px 8px',
                  fontSize: 12,
                }}
              />
            </div>
            <div style={{ maxHeight: 220, overflowY: 'auto', borderTop: '1px solid #f3f3f3', borderBottom: '1px solid #f3f3f3' }}>
              {(() => {
                const allValues = filterMenu.values.map((v) => v.value);
                const isAllSelected = allValues.length > 0 && filterSel.size === allValues.length;
                return (
                  <label style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px', cursor: 'pointer', borderBottom: '1px solid #f5f5f5' }}>
                    <input
                      type="checkbox"
                      checked={isAllSelected}
                      onChange={() => setFilterSel(isAllSelected ? new Set() : new Set(allValues))}
                    />
                    (Select All)
                  </label>
                );
              })()}

              {filterMenu.values
                .filter((item) => !filterSearch || item.value.toLowerCase().includes(filterSearch.toLowerCase()))
                .map((item) => (
                  <label
                    key={`filter-value-${item.value}`}
                    style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px', cursor: 'pointer' }}
                  >
                    <input
                      type="checkbox"
                      checked={filterSel.has(item.value)}
                      onChange={() => {
                        setFilterSel((prev) => {
                          const next = new Set(prev);
                          if (next.has(item.value)) next.delete(item.value);
                          else next.add(item.value);
                          return next;
                        });
                      }}
                    />
                    <span style={{ flex: 1 }}>{item.value || '(Blank)'}</span>
                    <span style={{ color: '#888' }}>({item.count})</span>
                  </label>
                ))}
            </div>
            <div style={{ padding: 8, display: 'flex', justifyContent: 'space-between', gap: 8 }}>
              <button
                type="button"
                onClick={() => {
                  const allValues = filterMenu.values.map((v) => v.value);
                  setFilterSel(new Set(allValues));
                }}
              >
                Reset
              </button>
              <button type="button" onClick={() => filterMenu.clear()}>Clear</button>
              <button type="button" onClick={() => filterMenu.apply(Array.from(filterSel))}>OK</button>
            </div>
          </div>
        )}
        </div>

        {appConfig.enableComments && (
        <CommentPanel
          isOpen={showCommentPanel}
          workbook={workbook}
          activeSheetName={activeSheet}
          defaultAuthor={appConfig.authorName}
          onClose={() => setShowCommentPanel(false)}
          onNavigate={navigateToCommentCell}
          onNewComment={() => {
            const target = selectedCell ?? { row: 1, col: 1 };
            openCommentEditorForAddress(target);
          }}
        />
        )}
      </div>

      {/* Sheet Tabs */}
      {appConfig.showSheetTabs && (
      <SheetTabs
        sheets={sheets}
        activeSheet={activeSheet}
        onSheetChange={handleSheetChange}
        onAddSheet={handleAddSheet}
        onRenameSheet={handleRenameSheet}
        onDeleteSheet={handleDeleteSheet}
        renameActiveSheetTrigger={renameSheetTrigger}
      />
      )}

      {/* Status Bar */}
      {appConfig.showStatusBar && (
      <StatusBar
        zoom={zoom}
        onZoomChange={handleZoomChange}
        selection={selection}
        workbook={workbook}
        viewMode={viewMode}
        onViewModeChange={handleViewModeChange}
        statusMessage={statusBarLink}
      />
      )}

      {/* Mini Toolbar — Excel-style formatting bar above context menu */}
      {miniToolbar && (
        <MiniToolbar
          x={miniToolbar.x}
          y={miniToolbar.y}
          menuTopY={contextMenuLayout?.y ?? contextMenu?.y}
          persistWithContextMenu={!!contextMenu}
          onClose={() => setMiniToolbar(null)}
          recentFillColors={recentFillColors}
          recentFontColors={recentFontColors}
          currentFont={miniToolbarCellStyle?.fontFamily || 'Calibri'}
          currentFontSize={miniToolbarCellStyle?.fontSize || 11}
          isBold={miniToolbarCellStyle?.bold || false}
          isItalic={miniToolbarCellStyle?.italic || false}
          isUnderline={!!miniToolbarCellStyle?.underline}
          currentFontColor={resolveStyleColor(miniToolbarCellStyle?.color)}
          currentFillColor={resolveStyleFillColor(miniToolbarCellStyle?.fill)}
          numberFormatCategory={getNumberFormatCategory(miniToolbarCellStyle?.numberFormat)}
          onFontChange={(font) => {
            applyContextMenuFormatting((addresses) => formattingController?.setFontFamily(addresses, font));
          }}
          onFontSizeChange={(size) => {
            if (!Number.isFinite(size)) return;
            applyContextMenuFormatting((addresses) => formattingController?.setFontSize(addresses, size));
          }}
          onBoldToggle={() => {
            applyContextMenuFormatting((addresses) => formattingController?.toggleBold(addresses));
          }}
          onItalicToggle={() => {
            applyContextMenuFormatting((addresses) => formattingController?.toggleItalic(addresses));
          }}
          onUnderlineToggle={() => {
            applyContextMenuFormatting((addresses) => formattingController?.toggleUnderline(addresses));
          }}
          onFillColor={(color) => {
            applyContextMenuFormatting((addresses) => formattingController?.setFill(addresses, color));
            addRecentColor(color, setRecentFillColors);
          }}
          onFontColor={(color) => {
            applyContextMenuFormatting((addresses) => formattingController?.setFontColor(addresses, color));
            addRecentColor(color, setRecentFontColors);
          }}
          onCurrencyFormat={() => {
            applyContextMenuFormatting((addresses) => formattingController?.applyNumberFormatPreset(addresses, 'currency'));
          }}
          onPercentFormat={() => {
            applyContextMenuFormatting((addresses) => formattingController?.applyNumberFormatPreset(addresses, 'percentage'));
          }}
          onCommaFormat={() => {
            applyContextMenuFormatting((addresses) => formattingController?.setNumberFormat(addresses, '#,##0'));
          }}
          onIncreaseDecimal={() => {
            applyContextMenuFormatting((addresses) => formattingController?.increaseDecimalPlaces(addresses));
          }}
          onDecreaseDecimal={() => {
            applyContextMenuFormatting((addresses) => formattingController?.decreaseDecimalPlaces(addresses));
          }}
        />
      )}

      {/* Form Control context menu */}
      {formControlContextMenu && (
        <ContextMenu
          x={formControlContextMenu.x}
          y={formControlContextMenu.y}
          items={[
            {
              id: 'format-control',
              label: 'Format Control…',
              onClick: () => openFormatControlDialog(formControlContextMenu.objectId),
            },
          ]}
          onClose={() => setFormControlContextMenu(null)}
        />
      )}

      {/* Format Control Dialog */}
      {formatControlObjectId && !formControlCellLinkPicker && (() => {
        const fc = drawingLayer.getObject(formatControlObjectId) as FormControlObject | undefined;
        if (!fc || fc.type !== 'formControl') return null;
        return (
          <FormatControlDialog
            isOpen
            control={fc}
            drawingLayer={drawingLayer}
            defaultSize={getFormControlDefaultSize(fc.controlType)}
            pickedCellLink={pickedFormControlCellLink}
            onClose={() => {
              setFormatControlObjectId(null);
              setPickedFormControlCellLink(null);
            }}
            onApply={handleApplyFormatControl}
            onPickCellLink={() => {
              setFormControlCellLinkPicker(true);
              setFormatControlObjectId(formatControlObjectId);
            }}
          />
        );
      })()}

      {formControlCellLinkPicker && (
        <div
          style={{
            position: 'fixed',
            bottom: 16,
            left: '50%',
            transform: 'translateX(-50%)',
            background: '#fff',
            border: '1px solid #ccc',
            borderRadius: 4,
            padding: '8px 16px',
            boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
            zIndex: 10025,
            fontFamily: 'Segoe UI, sans-serif',
            fontSize: 13,
          }}
        >
          Select a cell for the link, then click to confirm.
          <button
            type="button"
            style={{ marginLeft: 12, padding: '4px 10px', cursor: 'pointer' }}
            onClick={() => {
              setFormControlCellLinkPicker(false);
            }}
          >
            Cancel
          </button>
        </div>
      )}

      {/* Context Menu */}
      {appConfig.showContextMenu && contextMenu && (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          items={getContextMenuItems()}
          onLayoutChange={setContextMenuLayout}
          onClose={() => {
            setContextMenu(null);
            setContextMenuLayout(null);
            contextMenuTargetCellRef.current = null;
            setMiniToolbar(null);
          }}
        />
      )}

      {/* Insert / Delete cells dialog (Excel-style shift direction) */}
      {insertDeleteDialog && (
        <InsertDeleteCellsDialog
          type={insertDeleteDialog.type}
          onCancel={() => setInsertDeleteDialog(null)}
          onConfirm={(mode) => {
            applyInsertDelete(insertDeleteDialog.type, insertDeleteDialog.range, mode);
            setInsertDeleteDialog(null);
          }}
        />
      )}

      {isCreateTableOpen && workbook.activeSheet && (
        <CreateTableDialog
          isOpen={isCreateTableOpen}
          worksheet={workbook.activeSheet}
          selection={selectionRange}
          styleName={DEFAULT_INSERT_TABLE_STYLE.name}
          onClose={() => {
            setIsCreateTableOpen(false);
            setTableInsertError(null);
          }}
          onConfirm={handleCreateTable}
        />
      )}

      {tableInsertError && (
        <div
          style={{
            position: 'fixed',
            bottom: 24,
            left: '50%',
            transform: 'translateX(-50%)',
            background: '#fde7e9',
            color: '#a4262c',
            border: '1px solid #f1aeb5',
            borderRadius: 4,
            padding: '10px 16px',
            zIndex: 10013,
            fontFamily: 'Segoe UI, Arial, sans-serif',
            fontSize: 13,
          }}
        >
          {tableInsertError}
        </div>
      )}

      {/* Create PivotTable Dialog */}
      {isCreatePivotTableOpen && workbook.activeSheet && (
        <CreatePivotTableDialog
          isOpen={isCreatePivotTableOpen}
          worksheet={workbook.activeSheet}
          selection={selectionRange}
          onClose={() => {
            setIsCreatePivotTableOpen(false);
            setPivotTableError(null);
          }}
          onCreate={handleCreatePivotTable}
        />
      )}

      {pivotTableError && (
        <div
          style={{
            position: 'fixed',
            bottom: 24,
            left: '50%',
            transform: 'translateX(-50%)',
            background: '#fde7e9',
            color: '#a4262c',
            border: '1px solid #f1aeb5',
            borderRadius: 4,
            padding: '10px 16px',
            zIndex: 10013,
            fontFamily: 'Segoe UI, Arial, sans-serif',
            fontSize: 13,
          }}
        >
          {pivotTableError}
        </div>
      )}

      {/* Find/Replace Dialog */}
      {isFindReplaceOpen && workbook.activeSheet && (
        <FindReplaceDialog
          isOpen={isFindReplaceOpen}
          onClose={() => setIsFindReplaceOpen(false)}
          worksheet={workbook.activeSheet}
          initialTab={findReplaceTab}
          onMatchSelected={(address) => {
            setSelectedCell(address);
            setSelection({ start: address, end: address });
            renderer?.setSelection({
              start: address,
              end: address,
            });
            renderer?.scrollToCell({ row: address.row, col: address.col });
            renderer?.scheduleRedraw?.();
          }}
        />
      )}

      {/* Format Cells Dialog */}
      {isFormatDialogOpen && selection && (
        <FormatCellsDialog
          isOpen={isFormatDialogOpen}
          onClose={() => setIsFormatDialogOpen(false)}
          onApply={(changes: FormattingChanges) => {
            const sheet = workbook.activeSheet;
            if (!sheet) return;
            
            const r1 = Math.min(selection.start.row, selection.end.row);
            const r2 = Math.max(selection.start.row, selection.end.row);
            const c1 = Math.min(selection.start.col, selection.end.col);
            const c2 = Math.max(selection.start.col, selection.end.col);
            
            // Apply formatting changes to all cells in selection
            for (let r = r1; r <= r2; r++) {
              for (let c = c1; c <= c2; c++) {
                const addr = { row: r, col: c };
                const currentStyle = sheet.getCellStyle(addr) || {};
                const newStyle = { ...currentStyle };
                
                // Apply font changes
                if (changes.font) {
                  if (changes.font.fontFamily !== undefined) newStyle.fontFamily = changes.font.fontFamily;
                  if (changes.font.fontSize !== undefined) newStyle.fontSize = changes.font.fontSize;
                  if (changes.font.bold !== undefined) newStyle.bold = changes.font.bold;
                  if (changes.font.italic !== undefined) newStyle.italic = changes.font.italic;
                  if (changes.font.underline !== undefined) newStyle.underline = changes.font.underline !== 'none';
                  if (changes.font.color !== undefined) newStyle.color = changes.font.color;
                }
                
                // Apply fill changes
                if (changes.fill?.backgroundColor !== undefined) {
                  newStyle.fillColor = changes.fill.backgroundColor;
                }
                
                // Apply alignment changes
                if (changes.alignment) {
                  if (changes.alignment.horizontal !== undefined) newStyle.horizontalAlign = changes.alignment.horizontal;
                  if (changes.alignment.vertical !== undefined) newStyle.verticalAlign = changes.alignment.vertical;
                }
                
                // Apply number format changes
                if (changes.number?.numberFormat !== undefined) {
                  sheet.setCellNumberFormat(addr, changes.number.numberFormat);
                }
                
                // Intern and apply style
                const internedStyle = sheet.internStyle(newStyle);
                sheet.setCellStyle(addr, internedStyle);
              }
            }
            
            renderer?.scheduleRedraw();
            setIsFormatDialogOpen(false);
          }}
          selectedCells={[
            { row: selection.start.row, col: selection.start.col }
          ]}
          currentFormatting={(
            workbook.activeSheet?.getCellStyle(
              { row: selection.start.row, col: selection.start.col }
            ) || {}
          )}
        />
      )}

      {/* Data Validation Dropdown */}
      {validationDropdown && (
        <DropdownList
          items={validationDropdown.items}
          value={validationDropdown.currentValue}
          onSelect={(value) => {
            const sheet = workbook.activeSheet;
            if (sheet && canEditCellValue(appConfig, value)) {
              sheet.setCellValue(validationDropdown.cellAddress, value);
              renderer?.scheduleRedraw();
            }
            setValidationDropdown(null);
          }}
          onClose={() => setValidationDropdown(null)}
          cellX={validationDropdown.cellBounds.x}
          cellY={validationDropdown.cellBounds.y}
          cellWidth={validationDropdown.cellBounds.width}
          cellHeight={validationDropdown.cellBounds.height}
        />
      )}

      {/* Cell Comment Dialog */}
      {commentDialog && (
        <CellCommentDialog
          isOpen={!!commentDialog}
          cellLabel={formatCellAddress(commentDialog.address)}
          existingComments={commentDialog.comments}
          onClose={() => setCommentDialog(null)}
          onSave={handleCommentDialogSave}
          onDelete={handleCommentDialogDelete}
        />
      )}

      <HeaderFooterDialog
        isOpen={headerFooterDialogOpen}
        settings={workbook.activeSheet?.getHeaderFooter() ?? {
          header: { left: '', center: '', right: '' },
          footer: { left: '', center: '', right: '' },
        }}
        onClose={() => setHeaderFooterDialogOpen(false)}
        onSave={handleSaveHeaderFooter}
      />

      <WordArtGalleryDialog
        isOpen={wordArtGalleryOpen}
        onClose={() => setWordArtGalleryOpen(false)}
        onSelectPreset={handleWordArtPresetSelected}
      />

      {hyperlinkDialog && (
        <InsertHyperlinkDialog
          isOpen={!!hyperlinkDialog}
          cellLabel={formatCellAddress(hyperlinkDialog.address)}
          cellAddress={hyperlinkDialog.address}
          initialDisplayText={hyperlinkDialog.displayText}
          existingHyperlink={hyperlinkDialog.existingHyperlink}
          sheetNames={workbook.getSheetNames()}
          activeSheetName={activeSheet}
          onClose={() => setHyperlinkDialog(null)}
          onSave={handleHyperlinkDialogSave}
          onRemove={handleHyperlinkDialogRemove}
        />
      )}

      {/* Backstage File Menu */}
      {backstageOpen && appConfig.allowOpen && (
        <BackstageContainer
          isOpen={backstageOpen}
          onClose={() => setBackstageOpen(false)}
          initialPanel={backstagePanel}
          fileOperations={fileOperations}
          workbookMetadata={workbookMetadata}
          workbook={workbook}
          onWorkbookLoaded={onWorkbookLoaded ? (newWorkbook: Workbook) => {
            applyLoadedWorkbook(newWorkbook);
            setBackstageOpen(false);
          } : undefined}
          onCreateBlankWorkbook={() => {
            // Handle new blank workbook
            console.log('Creating new blank workbook');
            setBackstageOpen(false);
          }}
          onCreateFromTemplate={(templateId) => {
            console.log('Creating workbook from template:', templateId);
            setBackstageOpen(false);
          }}
          onOpenFile={(fileId) => {
            console.log('Opening file:', fileId);
            setBackstageOpen(false);
          }}
          onExportComplete={(blob) => {
            // Trigger download
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${fileName}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
            console.log('Export complete, download triggered');
          }}
          onVersionRestored={() => {
            console.log('Version restored');
            renderer?.scheduleRedraw();
          }}
        />
      )}
    </div>
  );
};

export const ExcelApp: React.FC<ExcelAppProps> = (props) => (
  <CyberSheetConfigProvider config={props.config}>
    <ExcelAppView {...props} />
  </CyberSheetConfigProvider>
);
