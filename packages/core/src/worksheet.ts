import { Address, Cell, CellStyle, CellComment, CellHyperlink, CellIcon, ColumnFilter, MergedRegion, Range, SheetEvents, IFormulaEngine, type CellValue, type DataValidationRule, type SheetProtectionOptions, type FreezeState, type SortKey, type AutoFilterRange, type CustomCellComponent } from './types';
import { ConditionalFormattingRule } from './ConditionalFormattingEngine';
import { Emitter } from './events';
import { SearchOptions, SearchRange, SearchResult, SpecialCellsOptions, SpecialCellValue } from './types/search-types';
import { extractReferences } from './utils/formula-reference-extractor';
import {
  buildMatcher,
  cellValueToString,
  compareRowMajor,
  compareColMajor,
  isStrictlyAfterRowMajor,
  isStrictlyBeforeRowMajor,
  styleMatchesFormat,
} from './search-engine';
import type { ICellStore } from './storage/ICellStore';
import { ColumnarCellStore } from './storage/ColumnarCellStore';
import type { IMergeStore } from './storage/MergeStore';
import { MergeStoreV1, MergeConflictError } from './storage/MergeStore';
import type { IVisibilityStore } from './storage/VisibilityStore';
import { VisibilityStoreV1 } from './storage/VisibilityStore';
import {
  DependencyGraph,
  RecalcCoordinator,
  type CycleDiagnostic,
  type RecalcResult,
  type IterativeRecalcResult,
  type RecalcIterationPolicy,
  packKey,
  unpackKey,
} from './dag/DependencyGraph';
import { FORMAT_VERSION, type WorksheetSnapshot } from './persistence/SnapshotCodec';
import {
  cloneHeaderFooterSettings,
  DEFAULT_HEADER_FOOTER,
  type HeaderFooterSettings,
} from './headerFooter';
export type { WorksheetSnapshot } from './persistence/SnapshotCodec';

export const DEFAULT_WORKSHEET_ROWS = 5000;
export const DEFAULT_WORKSHEET_COLS = 16384;

export class Worksheet {
  readonly name: string;
  /** Cell store — ICellStore boundary; swap implementation without touching any other Worksheet code. */
  private cells: ICellStore = new ColumnarCellStore();
  private colWidths = new Map<number, number>(); // px
  private rowHeights = new Map<number, number>(); // px
  /** Data validation rules, keyed by "row:col". */
  private validationStore = new Map<string, DataValidationRule>();
  /** Sheet-level protection settings, or null when not protected. */
  private sheetProtection: SheetProtectionOptions | null = null;
  /** Freeze-pane state, or null when no panes are frozen. */
  private freezeState: FreezeState | null = null;
  /** Auto-filter range marker (header row + column span), or null if not set. */
  private autoFilterRange: AutoFilterRange | null = null;
  private filters = new Map<number, ColumnFilter>();
  private events = new Emitter<SheetEvents>();
  private formulaEngine?: IFormulaEngine;
  /**
   * Merge store — IMergeStore boundary; swap implementation without touching any other Worksheet code.
   * To rollback: replace MergeStoreV1 with MergeStoreLegacy (same interface, O(n) ops).
   */
  private mergeStore: IMergeStore = new MergeStoreV1();
  /**
   * Visibility store — IVisibilityStore boundary.
   * Stores hidden rows and columns as Set<number> — all queries O(1).
   * To rollback: replace VisibilityStoreV1 with VisibilityStoreLegacy (same interface, O(n) ops).
   * Hidden state is NEVER written into CellRecord — this is an invariant.
   */
  private visibilityStore: IVisibilityStore = new VisibilityStoreV1();
  /**
   * Dependency graph — Phase 4.
   * Tracks formula dependencies as a directed graph: precedent → dependent.
   * Provides BFS dirty propagation + Kahn topo evaluation order + Tarjan SCC cycle detection.
   * Decoupled from ICellStore — the DAG holds only NodeKey (packed row/col keys).
   */
  private readonly dag = new DependencyGraph();
  /**
   * RecalcCoordinator — Phase 4.
   * Owns the recalculation pipeline: register deps → notify changed → recalc(evalFn).
   * Exposes registerDependencies() and recalc() as the public DAG API on Worksheet.
   */
  private readonly recalcCoordinator = new RecalcCoordinator(this.dag);
  private conditionalRules: ConditionalFormattingRule[] = [];
  private workbook?: any; // Reference to parent Workbook (for StyleCache access)
  /** Phase 28: Internal worksheet ID for pivot registry */
  private worksheetId: string;
  /** Reference to SpreadsheetEngine for E2 invariant enforcement */
  private _engine?: { isMutating(): boolean };
  /** Per-sheet header and footer text (left / center / right sections). */
  private headerFooter: HeaderFooterSettings = cloneHeaderFooterSettings(DEFAULT_HEADER_FOOTER);
  private progressiveLoadTotalRows = 0;
  private progressiveLoadDone = true;
  private readonly progressiveLoadedRows = new Set<number>();
  rowCount: number;
  colCount: number;

  // Transaction-level event buffering (Phase 11 enforcement)
  private _inTransaction = false;
  private _pendingEvents: SheetEvents[] = [];

  constructor(name: string, rows = DEFAULT_WORKSHEET_ROWS, cols = DEFAULT_WORKSHEET_COLS, engine?: IFormulaEngine, workbook?: any, spreadsheetEngine?: { isMutating(): boolean }) {
    this.name = name;
    this.rowCount = rows;
    this.colCount = cols;
    this.formulaEngine = engine;
    this.workbook = workbook;
    this._engine = spreadsheetEngine;
    this.worksheetId = `ws-${name}-${Date.now()}`; // Phase 28: Unique ID
  }

  on(listener: (e: SheetEvents) => void) {
    return this.events.on(listener);
  }

  beginProgressiveLoad(totalRows = this.rowCount): void {
    this.progressiveLoadTotalRows = totalRows;
    this.progressiveLoadDone = false;
    this.progressiveLoadedRows.clear();
    this.events.emit({
      type: 'progressive-load-changed',
      loadedRows: 0,
      totalRows,
      done: false,
    });
  }

  markRowsLoaded(startRow: number, endRow: number): void {
    if (endRow < startRow) return;
    for (let row = Math.max(1, startRow); row <= endRow; row++) {
      this.progressiveLoadedRows.add(row);
    }
    this.events.emit({
      type: 'progressive-load-changed',
      loadedRows: this.progressiveLoadedRows.size,
      totalRows: this.progressiveLoadTotalRows || this.rowCount,
      done: this.progressiveLoadDone,
    });
  }

  finishProgressiveLoad(): void {
    this.progressiveLoadDone = true;
    this.events.emit({
      type: 'progressive-load-changed',
      loadedRows: this.progressiveLoadedRows.size,
      totalRows: this.progressiveLoadTotalRows || this.rowCount,
      done: true,
    });
  }

  isProgressiveLoading(): boolean {
    return !this.progressiveLoadDone;
  }

  isRowLoaded(row: number): boolean {
    return this.progressiveLoadDone || this.progressiveLoadedRows.has(row);
  }

  /**
   * Assert that we're in a valid mutation state.
   * Throws if called outside engine.run() (E2 invariant enforcement).
   */
  private assertMutating(operation: string): void {
    if (this._engine && !this._engine.isMutating()) {
      throw new Error(
        `[E2 INVARIANT VIOLATION] ${operation} called outside engine.run(). ` +
        'All mutations must occur within the engine.run() callback.'
      );
    }
  }

  /**
   * Check if a transaction is currently active.
   */
  isTransactionActive(): boolean {
    return this._inTransaction;
  }

  /**
   * Run a batch of mutations as a single transaction.
   * Events are buffered and emitted once after all operations complete.
   * Re-entrant: nested calls reuse the outer transaction.
   */
  runTransaction<T>(fn: () => T): T {
    if (this._inTransaction) {
      // Re-entrant call — reuse existing transaction
      return fn();
    }

    this._inTransaction = true;
    try {
      const result = fn();
      this._inTransaction = false;
      this._flushEvents();
      return result;
    } catch (e) {
      this._inTransaction = false;
      this._pendingEvents = [];
      throw e;
    }
  }

  /**
   * Emit or buffer an event depending on transaction state.
   */
  private _emitOrBuffer(event: SheetEvents): void {
    if (this._inTransaction) {
      this._pendingEvents.push(event);
    } else {
      this.events.emit(event);
    }
  }

  /**
   * Flush pending events after transaction commit.
   * Coalesces events into one final aggregate event.
   */
  private _flushEvents(): void {
    if (this._pendingEvents.length === 0) return;

    // For now, emit the last event (aggregate coalescing can be added later)
    const lastEvent = this._pendingEvents[this._pendingEvents.length - 1];
    this._pendingEvents = [];
    this.events.emit(lastEvent);
  }

  /**
   * Get a cell object (returns IMMUTABLE view to prevent direct mutation).
   * 
   * CRITICAL: This returns a frozen object to enforce E2 invariant.
   * Direct mutation of returned cells is impossible:
   *   const cell = ws.getCell(addr);
   *   cell.value = 10;  // ❌ TypeError: Cannot assign to read only property
   * 
   * All mutations must go through setCellValue(), setCellFormula(), etc.
   */
  getCell(addr: Address): Readonly<Cell> | undefined {
    // Redirect non-anchor merged cells to their anchor.
    const anchor = this.mergeStore.getAnchor(addr.row, addr.col);
    const { row, col } = anchor ?? addr;
    const cell = this.cells.get(row, col);
    
    // Return frozen copy to prevent direct mutation
    // This closes the "read → mutate via returned object" escape hatch
    return cell ? Object.freeze({ ...cell }) : undefined;
  }

  /**
   * Get cell by packed NodeKey (used by GraphInvariantValidator).
   * Convenience method for DAG validation.
   * 
   * @internal DEV/TEST only
   */
  getCellByKey(nodeKey: number): Readonly<Cell> | undefined {
    const addr = unpackKey(nodeKey);
    return this.getCell(addr);
  }

  /**
   * Get-or-create a cell entry in the sparse Map.
   *
   * Eliminates the `getCell() ?? { value: null }` anti-pattern that produced
   * detached phantom objects (Infrastructure Risk logged Day 2).
   *
   * All mutation helpers (setCellValue, setCellFormula, setCellStyle, etc.)
   * must go through this method.
   */
  /**
   * Resolve the effective address for any cell.
   * If the address is a non-anchor member of a merged region, returns the anchor.
   * Otherwise returns the address unchanged.
   * O(1) via MergeStoreV1.getAnchor().
   */
  private resolveAnchor(addr: Address): Address {
    return this.mergeStore.getAnchor(addr.row, addr.col) ?? addr;
  }

  /**
   * Internal: Get-or-create cell for mutation.
   * 
   * SAFETY: This is called ONLY from public mutation methods that already
   * enforce assertMutating(). The mutation choke point is at the public API
   * boundary, not here. This keeps internal implementation flexible while
   * maintaining invariant enforcement at the architectural boundary.
   */
  private ensureCell(addr: Address): Cell {
    const resolved = this.resolveAnchor(addr);
    return this.cells.getOrCreate(resolved.row, resolved.col);
  }

  getCellValue(addr: Address): Cell['value'] {
    return this.getCell(addr)?.value ?? null;
  }

  setCellValue(addr: Address, value: Cell['value']): void {
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setCellValue(addr, value));
    }

    const c = this.ensureCell(addr);
    const previousValue = c.value; // Phase 30b: Capture before mutation
    // Phase 30b: No-op guard — same value, no mutation, no event
    if (previousValue === value) {
      return;
    }
    c.value = value;
    if (this.formulaEngine) this.formulaEngine.onCellChanged?.(addr, c);
    this.recalcCoordinator.notifyChanged(addr.row, addr.col);
    this._emitOrBuffer({ type: 'cell-changed', address: addr, cell: { ...c }, previousValue });
  }

  /**
   * Set a formula (and optionally its pre-computed display value) on a cell.
   *
   * Uses ensureCell() to guarantee the cell object is registered in the Map
   * before the formula is attached — preventing the detached-object phantom
   * that affected lookIn:'formulas' search.
   *
   * @param addr          Cell address
   * @param formula       Formula string, e.g. '=SUM(A1:A10)'
   * @param displayValue  Optional pre-evaluated result (for read-only display)
   */
  setCellFormula(addr: Address, formula: string, displayValue?: Cell['value']): void {
    this.assertMutating('setCellFormula');
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setCellFormula(addr, formula, displayValue));
    }

    const c = this.ensureCell(addr);
    c.formula = formula;
    if (displayValue !== undefined) c.value = displayValue;
    
    // Phase 3: Automatically extract and register formula dependencies
    try {
      const dependencies = extractReferences(formula, addr);
      this.registerDependencies(addr, dependencies);
    } catch (error) {
      // Dependency extraction failed - log but continue
      console.warn(`Failed to extract dependencies from formula at ${addr.row}:${addr.col}:`, error);
    }
    
    if (this.formulaEngine) this.formulaEngine.onCellChanged?.(addr, c);
    this.recalcCoordinator.notifyChanged(addr.row, addr.col);
    this._emitOrBuffer({ type: 'cell-changed', address: addr, cell: { ...c } });
  }

  /**
   * Set spill source metadata on a cell (used exclusively by SpillEngine).
   * Does NOT fire events (spill is internal bookkeeping).
   */
  setSpillSource(addr: Address, source: Cell['spillSource']): void {
    this.assertMutating('setSpillSource');
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setSpillSource(addr, source));
    }

    const c = this.ensureCell(addr);
    const before = c.spillSource;
    c.spillSource = source;
    this._emitOrBuffer({ type: 'spill-source-changed', address: addr, before, after: source });
  }

  /**
   * Set spilledFrom reference on a cell (used exclusively by SpillEngine).
   * Does NOT fire events.
   */
  setSpilledFrom(addr: Address, from: Cell['spilledFrom']): void {
    this.assertMutating('setSpilledFrom');
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setSpilledFrom(addr, from));
    }

    const c = this.ensureCell(addr);
    const before = c.spilledFrom;
    c.spilledFrom = from;
    this._emitOrBuffer({ type: 'spill-from-changed', address: addr, before, after: from });
  }

  /**
   * Clear spill metadata from a spill-source cell.
   * Sets spillSource to undefined (preserves mono-shape — never delete).
   */
  clearSpillSource(addr: Address): void {
    this.assertMutating('clearSpillSource');
    if (!this._inTransaction) {
      return this.runTransaction(() => this.clearSpillSource(addr));
    }

    const c = this.cells.get(addr.row, addr.col);
    if (c) {
      const before = c.spillSource;
      c.spillSource = undefined;
      this._emitOrBuffer({ type: 'spill-source-changed', address: addr, before, after: undefined });
    }
  }

  /**
   * Clear spilledFrom reference from a spilled cell.
   * Sets spilledFrom to undefined (preserves mono-shape — never delete).
   */
  clearSpilledFrom(addr: Address): void {
    this.assertMutating('clearSpilledFrom');
    if (!this._inTransaction) {
      return this.runTransaction(() => this.clearSpilledFrom(addr));
    }

    const c = this.cells.get(addr.row, addr.col);
    if (c) {
      const before = c.spilledFrom;
      c.spilledFrom = undefined;
      this._emitOrBuffer({ type: 'spill-from-changed', address: addr, before, after: undefined });
    }
  }

  /**
   * Batch spill operation — captures before state, applies mutations, emits one atomic event.
   * Used by SpillEngine to ensure spill updates are reversible as a single unit.
   */
  applySpillBatch(
    changes: Array<{
      addr: Address;
      spillSource?: Cell['spillSource'];
      spilledFrom?: Cell['spilledFrom'];
    }>
  ): void {
    if (!this._inTransaction) {
      return this.runTransaction(() => this.applySpillBatch(changes));
    }

    const eventChanges: Array<{
      address: Address;
      before: { spillSource?: Cell['spillSource']; spilledFrom?: Cell['spilledFrom'] };
      after: { spillSource?: Cell['spillSource']; spilledFrom?: Cell['spilledFrom'] };
    }> = [];

    // Capture before state for all cells
    for (const change of changes) {
      let cell = this.cells.get(change.addr.row, change.addr.col);
      const before = {
        spillSource: cell?.spillSource,
        spilledFrom: cell?.spilledFrom,
      };

      // Apply mutation - only create cell if setting metadata (not just clearing)
      const isSet = change.spillSource !== undefined || change.spilledFrom !== undefined;
      
      if (isSet || cell) {
        if (!cell && isSet) {
          cell = this.ensureCell(change.addr);
        }
        if (cell) {
          if ('spillSource' in change) {
            cell.spillSource = change.spillSource;
          }
          if ('spilledFrom' in change) {
            cell.spilledFrom = change.spilledFrom;
          }
        }
      }

      const after = {
        spillSource: cell?.spillSource,
        spilledFrom: cell?.spilledFrom,
      };

      eventChanges.push({ address: change.addr, before, after });
    }

    // Emit one batch event
    this._emitOrBuffer({ type: 'spill-batch-changed', changes: eventChanges });
  }

  getCellStyle(addr: Address): CellStyle | undefined {
    return this.getCell(addr)?.style;
  }

  setCellStyle(addr: Address, style: CellStyle | undefined): void {
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setCellStyle(addr, style));
    }

    const c = this.ensureCell(addr);

    // Auto-intern through workbook StyleCache (entropy-resistant boundary)
    // Protects against XLSX import, UI mutations, and spread operators
    let internedStyle = style;
    if (style && this.workbook?.getStyleCache) {
      internedStyle = this.workbook.getStyleCache().intern(style);
    }

    c.style = internedStyle; // Reference to canonical style (not a copy)
    this._emitOrBuffer({ type: 'style-changed', address: addr, style: internedStyle });
  }

  // ==================== Conditional Formatting ====================

  setConditionalFormattingRules(rules: ConditionalFormattingRule[]): void;
  setConditionalFormattingRules(
    range: Range,
    rules: ConditionalFormattingRule[],
    options?: { replace?: boolean; source?: 'preset' | 'manual' }
  ): void;
  setConditionalFormattingRules(
    rangeOrRules: Range | ConditionalFormattingRule[],
    rulesOrOptions?: ConditionalFormattingRule[] | { replace?: boolean; source?: 'preset' | 'manual' },
    options?: { replace?: boolean; source?: 'preset' | 'manual' }
  ): void {
    if (Array.isArray(rangeOrRules)) {
      this.conditionalRules = rangeOrRules.slice();
      this.events.emit({ type: 'sheet-mutated' });
      return;
    }

    const range = rangeOrRules;
    const rules = Array.isArray(rulesOrOptions) ? rulesOrOptions : [];
    const opts = Array.isArray(rulesOrOptions) ? options : rulesOrOptions;
    const replace = opts?.replace ?? true;

    const normalizedRange = this.normalizeRange(range);
    const updatedRules = rules.map(rule => ({
      ...rule,
      ranges: [normalizedRange],
    }));

    if (replace) {
      this.conditionalRules = updatedRules;
    } else {
      this.conditionalRules = [...this.conditionalRules, ...updatedRules];
    }

    this.events.emit({ type: 'sheet-mutated' });
  }

  addConditionalFormattingRule(rule: ConditionalFormattingRule): void {
    this.conditionalRules.push(rule);
    this.events.emit({ type: 'sheet-mutated' });
  }

  clearConditionalFormatting(): void {
    this.conditionalRules = [];
    this.events.emit({ type: 'sheet-mutated' });
  }

  getConditionalFormattingRules(): ConditionalFormattingRule[] {
    return this.conditionalRules.slice();
  }

  getHeaderFooter(): HeaderFooterSettings {
    return cloneHeaderFooterSettings(this.headerFooter);
  }

  setHeaderFooter(settings: HeaderFooterSettings): void {
    this.headerFooter = cloneHeaderFooterSettings(settings);
    this.events.emit({ type: 'header-footer-changed', settings: this.getHeaderFooter() });
  }

  // ── Data Validation ──────────────────────────────────────────────────────

  setDataValidation(addr: Address, rule: DataValidationRule): void {
    this.validationStore.set(`${addr.row}:${addr.col}`, rule);
    this.events.emit({ type: 'sheet-mutated' });
  }

  getDataValidation(addr: Address): DataValidationRule | undefined {
    return this.validationStore.get(`${addr.row}:${addr.col}`);
  }

  removeDataValidation(addr: Address): void {
    this.validationStore.delete(`${addr.row}:${addr.col}`);
    this.events.emit({ type: 'sheet-mutated' });
  }

  getValidationCells(): Address[] {
    const result: Address[] = [];
    for (const key of this.validationStore.keys()) {
      const [row, col] = key.split(':').map(Number);
      result.push({ row, col });
    }
    return result.sort(compareRowMajor);
  }

  // ── Sheet Protection ──────────────────────────────────────────────────────

  /**
   * Enable sheet protection with optional capability flags.
   *
   * When the sheet is protected, any cell whose `style.locked` is not
   * explicitly `false` (Excel default: locked = true) is considered protected.
   * The kernel stores the protection state for query by SDK-level guards.
   */
  protectSheet(options: SheetProtectionOptions = {}): void {
    const before = this.sheetProtection;
    this.sheetProtection = options;
    this.events.emit({ type: 'sheet-protection-changed', before, after: options });
  }

  /** Remove sheet protection. */
  unprotectSheet(): void {
    const before = this.sheetProtection;
    this.sheetProtection = null;
    this.events.emit({ type: 'sheet-protection-changed', before, after: null });
  }

  /** Return `true` if the sheet is currently protected. */
  isSheetProtected(): boolean {
    return this.sheetProtection !== null;
  }

  /** Return the current protection settings, or `null` if not protected. */
  getSheetProtection(): SheetProtectionOptions | null {
    return this.sheetProtection === null ? null : { ...this.sheetProtection };
  }

  // ── Freeze Panes ─────────────────────────────────────────────────────────

  /**
   * Freeze the top `rows` rows and left `cols` columns.
   * Both values must be non-negative integers; pass 0 for "no freeze" on an axis.
   * Setting both to 0 is equivalent to calling `clearFreezePanes()`.
   */
  setFreezePanes(rows: number, cols: number): void {
    if (rows < 0 || cols < 0 || !Number.isInteger(rows) || !Number.isInteger(cols)) {
      throw new RangeError(
        `setFreezePanes: rows and cols must be non-negative integers (got ${rows}, ${cols})`
      );
    }
    const before = this.freezeState;
    this.freezeState = (rows === 0 && cols === 0) ? null : { rows, cols };
    this.events.emit({ type: 'freeze-panes-changed', before, after: this.freezeState });
  }

  /** Remove all frozen panes. */
  clearFreezePanes(): void {
    const before = this.freezeState;
    this.freezeState = null;
    this.events.emit({ type: 'freeze-panes-changed', before, after: null });
  }

  /** Return the current freeze-pane state, or `null` if no panes are frozen. */
  getFreezePanes(): FreezeState | null {
    return this.freezeState === null ? null : { ...this.freezeState };
  }

  /**
   * Get the used range (smallest rectangle that includes all non-empty cells).
   */
  getUsedRange(): Range | null {
    if (this.cells.size === 0) return null;

    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;

    this.cells.forEach((row, col) => {
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    });

    if (!Number.isFinite(minRow) || !Number.isFinite(minCol)) return null;

    return {
      start: { row: minRow, col: minCol },
      end: { row: maxRow, col: maxCol },
    };
  }

  /**
   * Get contiguous block around an address (expands until empty rows/cols).
   */
  getContiguousRange(anchor: Address): Range | null {
    if (!this.getCell(anchor)) return null;

    let minRow = anchor.row;
    let maxRow = anchor.row;
    let minCol = anchor.col;
    let maxCol = anchor.col;

    const isRowNonEmpty = (row: number, startCol: number, endCol: number): boolean => {
      for (let col = startCol; col <= endCol; col++) {
        if (this.getCell({ row, col })) return true;
      }
      return false;
    };

    const isColNonEmpty = (col: number, startRow: number, endRow: number): boolean => {
      for (let row = startRow; row <= endRow; row++) {
        if (this.getCell({ row, col })) return true;
      }
      return false;
    };

    let expanded = true;
    while (expanded) {
      expanded = false;
      if (minRow > 0 && isRowNonEmpty(minRow - 1, minCol, maxCol)) {
        minRow -= 1;
        expanded = true;
      }
      if (maxRow < this.rowCount && isRowNonEmpty(maxRow + 1, minCol, maxCol)) {
        maxRow += 1;
        expanded = true;
      }
      if (minCol > 0 && isColNonEmpty(minCol - 1, minRow, maxRow)) {
        minCol -= 1;
        expanded = true;
      }
      if (maxCol < this.colCount && isColNonEmpty(maxCol + 1, minRow, maxRow)) {
        maxCol += 1;
        expanded = true;
      }
    }

    return {
      start: { row: minRow, col: minCol },
      end: { row: maxRow, col: maxCol },
    };
  }

  setColumnFilter(col: number, filter: ColumnFilter): void {
    const before = this.filters.get(col) ?? null;
    this.filters.set(col, filter);
    this.events.emit({ type: 'filter-changed', col, filter, before });
  }

  clearColumnFilter(col: number): void {
    const before = this.filters.get(col) ?? null;
    this.filters.delete(col);
    this.events.emit({ type: 'filter-changed', col, filter: null, before });
  }

  /** Clear all column filters at once. Emits one `filter-changed` event per cleared column. */
  clearAllFilters(): void {
    for (const [col, filter] of this.filters) {
      this.filters.delete(col);
      this.events.emit({ type: 'filter-changed', col, filter: null, before: filter });
    }
  }

  /** Return a snapshot of all active filters keyed by column number. */
  getAllFilters(): Map<number, ColumnFilter> {
    return new Map(this.filters);
  }

  /**
   * Get current filter for a column, if any.
   */
  getColumnFilter(col: number): ColumnFilter | undefined {
    return this.filters.get(col);
  }

  getVisibleRowIndices(range?: Range): number[] {
    const rows = range ? { start: range.start.row, end: range.end.row } : { start: 1, end: this.rowCount };
    const result: number[] = [];
    for (let r = rows.start; r <= rows.end; r++) {
      if (!this.rowMatchesFilters(r)) continue;
      result.push(r);
    }
    return result;
  }

  /**
   * Get row indices that pass all filters except the one on the provided column.
   * Used for building distinct value menus so a column's own filter doesn't hide its domain.
   */
  getVisibleRowIndicesExcluding(columnToExclude: number): number[] {
    const result: number[] = [];
    for (let r = 1; r <= this.rowCount; r++) {
      if (!this.rowMatchesFilters(r, columnToExclude)) continue;
      result.push(r);
    }
    return result;
  }

  /**
   * Return distinct cell values for a column, with counts.
   * @param col         1-based column index.
   * @param visibleOnly When true, only row indices that pass all OTHER column filters are
   *                    considered (Excel behaviour for filter dropdowns).
   * @returns           Array of `{ value, count }` sorted descending by count.
   */
  getDistinctValues(col: number, visibleOnly = false): { value: string; count: number }[] {
    const rows = visibleOnly
      ? this.getVisibleRowIndicesExcluding(col)
      : Array.from({ length: this.rowCount }, (_, i) => i + 1);
    const counts = new Map<string, number>();
    for (const row of rows) {
      const v = this.getCell({ row, col })?.value ?? null;
      const key = v === null ? '' : String(v);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return [...counts.entries()]
      .map(([value, count]) => ({ value, count }))
      .sort((a, b) => b.count - a.count);
  }

  // ── Auto-Filter Range ─────────────────────────────────────────────────────

  /**
   * Mark the auto-filter region (the row bearing the dropdown arrows and the
   * columns it spans).  Purely a UI marker — does not affect filter matching.
   */
  setAutoFilterRange(headerRow: number, startCol: number, endCol: number): void {
    const before = this.autoFilterRange;
    this.autoFilterRange = { headerRow, startCol, endCol };
    this.events.emit({ type: 'autofilter-range-changed', before, after: this.autoFilterRange });
  }

  /** Remove the auto-filter range marker. */
  clearAutoFilterRange(): void {
    const before = this.autoFilterRange;
    if (before === null) return;
    this.autoFilterRange = null;
    this.events.emit({ type: 'autofilter-range-changed', before, after: null });
  }

  /** Return the current auto-filter range, or `null` if not set. */
  getAutoFilterRange(): AutoFilterRange | null {
    return this.autoFilterRange === null ? null : { ...this.autoFilterRange };
  }

  // ── Range Sort ────────────────────────────────────────────────────────────

  /**
   * Sort a rectangular range in-place by one or more sort keys.
   *
   * - Rows within the range (inclusive) are reordered according to the keys.
   * - Cells OUTSIDE the range are not touched.
   * - Sort is stable (ties preserve original row order).
   * - Emits `sort-applied` after completion (for PatchRecorder / UI notification).
   *
   * @param range  The rectangular region to sort (1-based, inclusive).
   * @param keys   Primary → secondary → … sort keys; each key references a column
   *               that must fall within `range.start.col` … `range.end.col`.
   */
  sortRange(range: Range, keys: SortKey[]): void {
    const { start, end } = range;
    if (keys.length === 0) return;

    const numCols = end.col - start.col + 1;

    // Collect all data rows as value snapshots.
    // Use Cell['value'] (CellValue | undefined) from the store directly.
    const rows: Array<{ rowIndex: number; values: Array<Cell['value'] | null> }> = [];
    for (let r = start.row; r <= end.row; r++) {
      const values: Array<Cell['value'] | null> = [];
      for (let c = start.col; c <= end.col; c++) {
        values.push(this.cells.get(r, c)?.value ?? null);
      }
      rows.push({ rowIndex: r, values });
    }

    // Stable sort by keys (primary first). Nulls are always sorted last
    // regardless of direction, so they are tested before the dir inversion.
    rows.sort((a, b) => {
      for (const key of keys) {
        const ci = key.col - start.col;
        const av = a.values[ci] ?? null;
        const bv = b.values[ci] ?? null;
        const aBlank = av === null || (typeof av === 'string' && av.trim() === '');
        const bBlank = bv === null || (typeof bv === 'string' && bv.trim() === '');
        // Blanks last — independent of direction.
        if (aBlank && bBlank) continue;
        if (aBlank) return 1;
        if (bBlank) return -1;
        const cmp = this._compareValues(av, bv, key.type ?? 'text');
        if (cmp !== 0) return key.dir === 'asc' ? cmp : -cmp;
      }
      return 0;
    });

    // Write sorted rows back; use direct store ops to avoid per-cell events.
    const sortedRows = rows.map(r => r.values);
    for (let ri = 0; ri < sortedRows.length; ri++) {
      const destRow = start.row + ri;
      for (let ci = 0; ci < numCols; ci++) {
        const destCol = start.col + ci;
        const v = sortedRows[ri]![ci] ?? null;
        if (v === null) {
          this.cells.delete(destRow, destCol);
        } else {
          const existing = this.cells.get(destRow, destCol);
          this.cells.set(destRow, destCol, { ...(existing ?? { value: null }), value: v });
        }
      }
    }

    this.events.emit({
      type: 'sort-applied',
      startRow: start.row, startCol: start.col,
      endRow:   end.row,   endCol:   end.col,
      keys,
    });
  }

  /** Compare two cell values for sorting. Nulls sort last in all directions. */
  private _compareValues(a: Cell['value'] | null, b: Cell['value'] | null, type: 'text' | 'number' | 'date'): number {
    if (a === null && b === null) return 0;
    if (a === null) return 1;  // nulls last
    if (b === null) return -1;
    if (type === 'number') {
      const na = Number(a), nb = Number(b);
      return isNaN(na) ? (isNaN(nb) ? 0 : 1) : isNaN(nb) ? -1 : na - nb;
    }
    // text / date: coerce to string
    const sa = String(a).toLowerCase();
    const sb = String(b).toLowerCase();
    return sa < sb ? -1 : sa > sb ? 1 : 0;
  }

  private rowMatchesFilters(row: number, excludeCol?: number): boolean {
    for (const [col, filter] of this.filters) {
      if (excludeCol != null && col === excludeCol) continue;
      const v = this.getCell({ row, col })?.value ?? null;
      if (!this.matchesFilter(v, filter)) return false;
    }
    return true;
  }

  private matchesFilter(v: Cell['value'], f: ColumnFilter): boolean {
    if (f.type === 'empty')      return v === null || v === '';
    if (f.type === 'notEmpty')   return !(v === null || v === '');
    if (v === null) return false;
    if (f.type === 'in') {
      const arr = Array.isArray(f.value) ? (f.value as any[]) : [];
      return arr.some(x => String(x) === String(v));
    }
    if (f.type === 'equals')      return v === f.value;
    if (f.type === 'notEquals')   return v !== f.value;
    const sv = String(v).toLowerCase();
    const fv = String(f.value ?? '').toLowerCase();
    if (f.type === 'contains')    return sv.includes(fv);
    if (f.type === 'notContains') return !sv.includes(fv);
    if (f.type === 'startsWith')  return sv.startsWith(fv);
    if (f.type === 'endsWith')    return sv.endsWith(fv);
    if (typeof v !== 'number') return false;
    if (f.type === 'gt')  return v >  (f.value as number);
    if (f.type === 'gte') return v >= (f.value as number);
    if (f.type === 'lt')  return v <  (f.value as number);
    if (f.type === 'lte') return v <= (f.value as number);
    if (f.type === 'between') {
      const [a, b] = f.value as [number, number];
      return v >= Math.min(a, b) && v <= Math.max(a, b);
    }
    return true;
  }

  getColumnWidth(col: number): number { return this.colWidths.get(col) ?? 80; }
  setColumnWidth(col: number, px: number): void { this.colWidths.set(col, px); this.events.emit({ type: 'sheet-mutated' }); }
  getRowHeight(row: number): number { return this.rowHeights.get(row) ?? 20; }
  setRowHeight(row: number, px: number): void { this.rowHeights.set(row, px); this.events.emit({ type: 'sheet-mutated' }); }

  setFormulaEngine(engine?: IFormulaEngine) { this.formulaEngine = engine; }

  // ==================== Phase 28: Pivot Registry Integration ====================

  /**
   * Get the worksheet's unique identifier.
   * Used by pivot registry to associate pivots with worksheets.
   */
  getWorksheetId(): string {
    return this.worksheetId;
  }

  // ==================== DAG / RecalcCoordinator APIs (Phase 4) ====================

  /**
   * Register (or update) the dependency list for a formula cell.
   *
   * Call this whenever a formula is parsed/re-parsed. The coordinator will:
   *   1. Replace any existing edges for this cell with the new dep list.
   *   2. Mark the cell dirty so it participates in the next recalc().
   *
   * @param addr  The formula cell (the dependent).
   * @param deps  The cells this formula reads (the precedents).
   */
  registerDependencies(addr: Address, deps: Address[]): void {
    this.recalcCoordinator.registerFormula(addr.row, addr.col, deps);
  }

  /**
   * Remove a formula cell's dependency edges from the DAG (e.g. after clearing).
   */
  clearDependencies(addr: Address): void {
    this.recalcCoordinator.clearFormula(addr.row, addr.col);
  }

  /**
   * Mark a cell dirty without changing its value (e.g. volatile refresh).
   */
  notifyChanged(addr: Address): void {
    this.recalcCoordinator.notifyChanged(addr.row, addr.col);
  }

  /**
   * Number of cells in the dirty set — useful for testing and UI indicators.
   */
  get dirtyCount(): number {
    return this.recalcCoordinator.dirtyCount;
  }

  /** Return true when the given cell is currently awaiting recalculation. */
  isDirty(addr: Address): boolean {
    return this.recalcCoordinator.isDirty(packKey(addr.row, addr.col));
  }

  /**
   * Run a full recalculation pass in topological order.
   *
   * The caller supplies an `evaluate` function that is called for each dirty
   * cell in dependency order (precedents before dependents). After the pass:
   *   - All dirty cells are cleared.
   *   - Any detected cycles are emitted as a 'cycle-detected' event.
   *
   * @param evaluate  Called once per dirty cell (identified by packed NodeKey).
   *                  Use `unpackKey(key)` from CellStoreV1 to get row/col.
   * @returns         RecalcResult with `evaluated` count and `cycles` array.
   */
  recalc(evaluate: (key: number) => void): RecalcResult {
    const result = this.recalcCoordinator.recalc(evaluate);
    if (result.cycles.length > 0) {
      this.events.emit({ type: 'cycle-detected', cycles: result.cycles });
    }
    return result;
  }

  /**
   * Automatically recalculate all dirty formula cells using the FormulaEngine.
   * 
   * This is a convenience method that wraps recalc() with automatic formula evaluation.
   * It evaluates each dirty cell's formula in topological order and updates cell values.
   * 
   * Requires that a FormulaEngine was provided to the constructor (via the engine parameter).
   * 
   * @returns RecalcResult with evaluated count and any detected cycles
   * @throws Error if no FormulaEngine is available
   * 
   * @example
   * worksheet.setCellValue({row: 0, col: 0}, 10);
   * worksheet.setCellFormula({row: 0, col: 1}, '=A1*2');
   * worksheet.autoRecalculate(); // B1 now has value 20
   */
  autoRecalculate(): RecalcResult {
    if (!this.formulaEngine) {
      throw new Error('Cannot auto-recalculate: no FormulaEngine available. Pass an engine to the Worksheet constructor.');
    }

    return this.recalc((nodeKey) => {
      this.evaluateFormulaNode(nodeKey);
    });
  }

  /**
   * Lazily evaluate a dirty formula cell only when its value is requested.
   *
   * If the cell is clean, returns the cached value without touching the DAG.
   * If the cell is dirty, evaluates the dirty subset needed for this address
   * and leaves unrelated off-screen dirty cells untouched.
   */
  evaluateIfNeeded(addr: Address): Cell['value'] | undefined {
    if (!this.formulaEngine) return this.getCellValue(addr);
    const key = packKey(addr.row, addr.col);
    if (!this.recalcCoordinator.isDirty(key)) return this.getCellValue(addr);

    const result = this.recalcCoordinator.recalcSubset([key], nodeKey => {
      this.evaluateFormulaNode(nodeKey);
    });
    if (result.cycles.length > 0) {
      this.events.emit({ type: 'cycle-detected', cycles: result.cycles });
    }
    return this.getCellValue(addr);
  }

  /**
   * Lazily evaluate all requested dirty cells as a single topologically sorted
   * batch. Intended for renderer viewport reads and scroll updates.
   */
  evaluateBatch(addresses: Address[]): RecalcResult {
    if (!this.formulaEngine || addresses.length === 0) {
      return { evaluated: 0, cycles: [] };
    }

    const keys = addresses.map(addr => packKey(addr.row, addr.col));
    const result = this.recalcCoordinator.recalcSubset(keys, nodeKey => {
      this.evaluateFormulaNode(nodeKey);
    });
    if (result.cycles.length > 0) {
      this.events.emit({ type: 'cycle-detected', cycles: result.cycles });
    }
    return result;
  }

  /**
   * Apply formula results produced off-thread without firing cell mutation
   * events or marking the formula cells dirty again.
   */
  applyFormulaResults(addresses: Address[], values: Float64Array): void {
    const clearKeys: number[] = [];
    const count = Math.min(addresses.length, values.length);
    for (let i = 0; i < count; i++) {
      const value = values[i];
      const addr = addresses[i];
      const cell = this.cells.get(addr.row, addr.col);
      if (!cell?.formula) continue;
      if (!Number.isNaN(value)) {
        cell.value = value;
      }
      clearKeys.push(packKey(addr.row, addr.col));
    }
    this.recalcCoordinator.clearDirtyKeys(clearKeys);
  }

  private evaluateFormulaNode(nodeKey: number): CellValue {
    const { row, col } = unpackKey(nodeKey);
    const cell = this.cells.get(row, col);

    if (!cell?.formula) return cell?.value as CellValue ?? null;

    try {
      const result = (this.formulaEngine as any).evaluate(cell.formula, {
        worksheet: this as any,
        currentCell: { row, col },
      });

      if (result && typeof result === 'object' && 'message' in result && result instanceof Error) {
        cell.value = result.message;
      } else if (Array.isArray(result)) {
        cell.value = result as any;
      } else {
        cell.value = result as CellValue;
      }
    } catch (error) {
      cell.value = '#ERROR!';
      console.error(`Formula evaluation error at ${row}:${col}:`, error);
    }

    return cell.value as CellValue;
  }

  /**
   * DAG statistics for diagnostics and tests.
   */
  get dagStats(): { nodes: number; edges: number; dirty: number; volatiles: number } {
    return this.recalcCoordinator.stats;
  }

  // ==================== DAG / RecalcCoordinator APIs (Phase 5) ====================

  /**
   * Register a formula cell as **volatile**.
   *
   * Volatile cells are re-evaluated on every recalc tick regardless of whether
   * their dependencies changed. Typical volatile functions: NOW, TODAY, RAND,
   * RANDBETWEEN, INDIRECT, OFFSET.
   *
   * @param addr  The formula cell.
   * @param deps  Cells this formula reads (may be empty for NOW/RAND).
   */
  registerVolatile(addr: Address, deps: Address[]): void {
    this.recalcCoordinator.registerVolatile(addr.row, addr.col, deps);
  }

  /**
   * Unregister a cell's volatile flag and remove its dependency edges.
   * Call when a volatile formula is deleted or overwritten with a plain value.
   */
  clearVolatile(addr: Address): void {
    this.recalcCoordinator.clearVolatile(addr.row, addr.col);
  }

  /**
   * Re-seed all registered volatile cells as dirty and propagate.
   *
   * Call this at the start of each render frame (e.g., from a 1-second
   * `setInterval`) to refresh NOW, RAND, etc.
   *
   * @complexity O(V_volatile_reach + E_volatile_reach)
   */
  flushVolatiles(): void {
    this.recalcCoordinator.flushVolatiles();
  }

  /** Number of cells currently registered as volatile. */
  get volatileCount(): number {
    return this.recalcCoordinator.volatileCount;
  }

  /**
   * Run an **iterative** recalculation pass that resolves circular references.
   *
   * Unlike `recalc()` (which treats cycles as hard errors), `recalcIterative()`
   * runs up to `policy.maxIterations` passes over cycle nodes until the values
   * converge within `policy.tolerance`.  Mirrors Excel's iterative calculation
   * mode.
   *
   * @param evaluate  Called once per cell per iteration pass.  Must return the
   *                  new `CellValue` so convergence can be tracked. Non-numeric
   *                  values are treated as always-converged for that cell.
   *
   * @param policy    Optional override of `DEFAULT_RECALC_ITERATION_POLICY`
   *                  (maxIterations=100, tolerance=0.001, algorithm='gauss-seidel').
   *
   * @returns         `IterativeRecalcResult` — evaluated count, cycles, iteration
   *                  count, convergence flag, and final maxDelta.
   */
  recalcIterative(
    evaluate: (key: number) => CellValue,
    policy?: RecalcIterationPolicy,
  ): IterativeRecalcResult {
    const result = this.recalcCoordinator.recalcIterative(evaluate, policy);
    if (result.cycles.length > 0) {
      this.events.emit({ type: 'cycle-detected', cycles: result.cycles });
    }
    return result;
  }

  // ==================== Merge APIs (Phase 2) ====================

  /**
   * Merge a rectangular range of cells.
   *
   * **Behavior**:
   *   - Only the anchor (top-left) cell retains its value, formula, and style.
   *   - All non-anchor cells are cleared (their data is discarded).
   *   - Reads/writes to non-anchor cells are automatically redirected to anchor.
   *   - Overlapping an existing merge is a hard error (MergeConflictError).
   *   - A 1×1 range is rejected (already a single cell, no merge possible).
   *
   * @throws MergeConflictError if any cell in the range is already in a merge.
   * @throws RangeError if the range is a single cell.
   */
  mergeCells(range: Range): void {
    const norm = this.normalizeRange(range);
    const region: MergedRegion = {
      startRow: norm.start.row,
      startCol: norm.start.col,
      endRow:   norm.end.row,
      endCol:   norm.end.col,
    };

    // Throws MergeConflictError on any overlap — propagates to caller.
    this.mergeStore.add(region);

    // Clear non-anchor cells from ICellStore — their data is now inaccessible
    // and must not participate in searches or formula evaluation.
    for (let r = region.startRow; r <= region.endRow; r++) {
      for (let c = region.startCol; c <= region.endCol; c++) {
        if (r === region.startRow && c === region.startCol) continue; // keep anchor
        this.cells.delete(r, c);
      }
    }

    this.events.emit({ type: 'merge-added', region });
    this.events.emit({ type: 'sheet-mutated' });
  }

  /**
   * Remove the merge that overlaps with the given range.
   *
   * Removes ALL merges whose region intersects the provided range.
   * For the common case of removing a merged cell by clicking on it,
   * pass the range of the merged region (or just a single address within it).
   *
   * Non-merged ranges are silently ignored (no error).
   */
  cancelMerge(range: Range): void {
    const norm = this.normalizeRange(range);
    const removed = this.mergeStore.removeOverlapping(
      norm.start.row, norm.start.col, norm.end.row, norm.end.col
    );
    for (const region of removed) {
      this.events.emit({ type: 'merge-removed', region });
    }
    if (removed.length > 0) this.events.emit({ type: 'sheet-mutated' });
  }

  /**
   * Return the merged region containing [addr], or null if not merged.
   * O(1) via MergeStoreV1.getRegion().
   */
  getMergedRangeForCell(addr: Address): Range | null {
    const region = this.mergeStore.getRegion(addr.row, addr.col);
    if (!region) return null;
    return {
      start: { row: region.startRow, col: region.startCol },
      end:   { row: region.endRow,   col: region.endCol   },
    };
  }

  /** All merged regions as Range[]. O(n_merges). */
  getMergedRanges(): Range[] {
    return this.mergeStore.getAll().map(r => ({
      start: { row: r.startRow, col: r.startCol },
      end:   { row: r.endRow,   col: r.endCol   },
    }));
  }

  /**
   * True if [addr] is inside any merged region (anchor or non-anchor).
   * O(1). Used by SpillEngine to block spill over merged areas.
   */
  isInMerge(addr: Address): boolean {
    return this.mergeStore.isInMerge(addr.row, addr.col);
  }

  /**
   * Return the anchor address for a cell that’s inside a merged region,
   * or null if the cell is not in any merge.
   * O(1). Used by formula engine and rendering.
   */
  getMergeAnchorFor(addr: Address): Address | null {
    return this.mergeStore.getAnchor(addr.row, addr.col);
  }
  // ==================== Visibility API (Phase 3) ====================
  //
  // Behavioral contract:
  //   getCell      — always returns data regardless of hidden state
  //   find         — skips hidden cells (row or col hidden) unless includeHidden: true
  //   iteration    — controlled via findIterator's includeHidden option
  //   merge+hidden — orthogonal: merge is spatial, hidden is visual; no conflict
  //   spill        — NOT blocked by hidden cells (hidden is display-only)
  //   unhide       — Set.delete() → O(1) → traversal immediately restored

  /**
   * Hide a row. Idempotent. O(1).
   * Does NOT affect cell data — cells in hidden rows are still accessible via getCell().
   */
  hideRow(row: number): void {
    if (this.visibilityStore.isRowHidden(row)) return; // already hidden — no-op
    this.visibilityStore.hideRow(row);
    this.events.emit({ type: 'row-hidden', row });
    this.events.emit({ type: 'sheet-mutated' });
  }

  /**
   * Restore a hidden row to visible. Idempotent. O(1).
   * find() and findIterator() will immediately include it again.
   */
  showRow(row: number): void {
    if (!this.visibilityStore.isRowHidden(row)) return; // already visible — no-op
    this.visibilityStore.showRow(row);
    this.events.emit({ type: 'row-shown', row });
    this.events.emit({ type: 'sheet-mutated' });
  }

  /**
   * Hide a column. Idempotent. O(1).
   */
  hideCol(col: number): void {
    if (this.visibilityStore.isColHidden(col)) return; // already hidden — no-op
    this.visibilityStore.hideCol(col);
    this.events.emit({ type: 'col-hidden', col });
    this.events.emit({ type: 'sheet-mutated' });
  }

  /**
   * Restore a hidden column to visible. Idempotent. O(1).
   */
  showCol(col: number): void {
    if (!this.visibilityStore.isColHidden(col)) return; // already visible — no-op
    this.visibilityStore.showCol(col);
    this.events.emit({ type: 'col-shown', col });
    this.events.emit({ type: 'sheet-mutated' });
  }

  /** Returns true if the row is currently hidden. O(1). */
  isRowHidden(row: number): boolean {
    return this.visibilityStore.isRowHidden(row);
  }

  /** Returns true if the column is currently hidden. O(1). */
  isColHidden(col: number): boolean {
    return this.visibilityStore.isColHidden(col);
  }

  /**
   * Returns true if a cell's row OR column is hidden.
   * Used by the formula engine and renderers — O(1).
   */
  isCellHidden(addr: Address): boolean {
    return this.visibilityStore.isCellHidden(addr.row, addr.col);
  }

  /**
   * Returns the set of all currently hidden row indices.
   * O(1) — returns the live ReadonlySet reference.
   */
  getHiddenRows(): ReadonlySet<number> {
    return this.visibilityStore.getHiddenRows();
  }

  /**
   * Returns the set of all currently hidden column indices.
   * O(1) — returns the live ReadonlySet reference.
   */
  getHiddenCols(): ReadonlySet<number> {
    return this.visibilityStore.getHiddenCols();
  }
  /**
   * Delete a cell: clears all its data.
   *
   * If the cell is a **merge anchor**, the merge is also cancelled
   * (PM behavior: “Delete anchor unmerges”).
   *
   * If the cell is a **non-anchor** member of a merge, the call is redirected
   * to the anchor and the merge is cancelled.
   *
   * After cancellation, all cells in the former merge region become independent
   * empty cells.
   */
  /**
   * Iterate over all non-empty cells in sparse order.
   * 
   * This is O(n) where n = number of non-empty cells (NOT total grid size).
   * Critical for InsertColumn/DeleteColumn commands to avoid O(n²) scanning.
   * 
   * @param callback - Called for each non-empty cell with (row, col, cell)
   */
  forEachNonEmptyCell(callback: (row: number, col: number, cell: any) => void): void {
    this.cells.forEach((row, col, cell) => {
      callback(row, col, cell);
    });
  }

  deleteCell(addr: Address): void {
    // Resolve to anchor in case a non-anchor cell was passed.
    const resolved = this.resolveAnchor(addr);

    // INVARIANT 10: State Minimality - actually remove cell from storage
    const cell = this.cells.get(resolved.row, resolved.col);
    if (!cell) {
      return; // Already deleted
    }

    // If anchor of a merged region, cancel the merge first.
    if (this.mergeStore.isAnchor(resolved.row, resolved.col)) {
      const region = this.mergeStore.getRegion(resolved.row, resolved.col)!;
      this.mergeStore.removeByAnchor(resolved.row, resolved.col);
      this.events.emit({ type: 'merge-removed', region });
    }

    // Phase 3: Clear formula dependencies if cell had a formula
    if (cell.formula) {
      this.clearDependencies(resolved);
    }

    // Remove cell from storage (not just clear - actually delete)
    this.cells.delete(resolved.row, resolved.col);
    
    // Phase 3: Notify change to mark dependents dirty (if any cells reference this one)
    this.recalcCoordinator.notifyChanged(resolved.row, resolved.col);

    this.events.emit({ type: 'cell-changed', address: resolved, cell: { value: null } });
  }

  private normalizeRange(r: Range): Range {
    return {
      start: { row: Math.min(r.start.row, r.end.row), col: Math.min(r.start.col, r.end.col) },
      end:   { row: Math.max(r.start.row, r.end.row), col: Math.max(r.start.col, r.end.col) },
    };
  }

  // ==================== Hyperlink APIs ====================

  getHyperlink(addr: Address): CellHyperlink | undefined {
    const link = this.getCell(addr)?.hyperlink;
    return link ? { ...link } : undefined;
  }

  setHyperlink(addr: Address, hyperlink: CellHyperlink | undefined): void {
    if (!this._inTransaction) {
      return this.runTransaction(() => this.setHyperlink(addr, hyperlink));
    }

    const c = this.ensureCell(addr);
    if (hyperlink) {
      c.hyperlink = { ...hyperlink };
    } else {
      delete c.hyperlink;
    }

    this._emitOrBuffer({
      type: 'cell-changed',
      address: addr,
      cell: { ...c },
      previousValue: c.value ?? null,
    });
  }

  removeHyperlink(addr: Address): void {
    this.setHyperlink(addr, undefined);
  }

  // ==================== Comment APIs ====================

  /**
   * Add a comment to a cell (supports threading via parentId)
   */
  addComment(addr: Address, comment: Omit<CellComment, 'id' | 'createdAt'>): CellComment {
    const c = this.cells.getOrCreate(addr.row, addr.col);
    if (!c.comments) c.comments = [];

    const newComment: CellComment = {
      id: this.generateCommentId(),
      createdAt: new Date(),
      ...comment,
    };

    c.comments.push(newComment);
    this.events.emit({ type: 'comment-added', address: addr, comment: newComment });
    return newComment;
  }

  /**
   * Get all comments for a specific cell
   */
  getComments(addr: Address): CellComment[] {
    return this.getCell(addr)?.comments ?? [];
  }

  /**
   * Update an existing comment
   */
  updateComment(addr: Address, commentId: string, updates: Partial<Omit<CellComment, 'id' | 'createdAt'>>): boolean {
    const c = this.cells.get(addr.row, addr.col);
    if (!c?.comments) return false;

    const idx = c.comments.findIndex(cm => cm.id === commentId);
    if (idx === -1) return false;

    c.comments[idx] = {
      ...c.comments[idx],
      ...updates,
      editedAt: new Date(),
    };

    this.events.emit({ type: 'comment-updated', address: addr, commentId, comment: c.comments[idx] });
    return true;
  }

  /**
   * Delete a comment (and its threaded replies)
   */
  deleteComment(addr: Address, commentId: string, deleteReplies = true): boolean {
    const c = this.cells.get(addr.row, addr.col);
    if (!c?.comments) return false;

    const before = c.comments.length;

    if (deleteReplies) {
      // Remove comment and all its children
      const toDelete = new Set<string>([commentId]);
      let changed = true;
      while (changed) {
        changed = false;
        for (const cm of c.comments) {
          if (cm.parentId && toDelete.has(cm.parentId) && !toDelete.has(cm.id)) {
            toDelete.add(cm.id);
            changed = true;
          }
        }
      }
      c.comments = c.comments.filter(cm => !toDelete.has(cm.id));
    } else {
      c.comments = c.comments.filter(cm => cm.id !== commentId);
    }

    if (c.comments.length === before) return false;

    // Assign undefined instead of `delete c.comments` — preserves V8 mono-shape.
    // `delete obj.prop` changes the hidden class; `obj.prop = undefined` does not.
    if (c.comments.length === 0) c.comments = undefined;

    this.events.emit({ type: 'comment-deleted', address: addr, commentId });
    return true;
  }

  /**
   * Get all cells with comments (useful for navigation)
   */
  getAllComments(): Array<{ address: Address; comments: CellComment[] }> {
    const result: Array<{ address: Address; comments: CellComment[] }> = [];

    this.cells.forEach((row, col, cell) => {
      if (cell.comments && cell.comments.length > 0) {
        result.push({
          address: { row, col },
          comments: [...cell.comments],
        });
      }
    });

    // Sort by row, then col for consistent navigation.
    result.sort((a, b) => {
      if (a.address.row !== b.address.row) return a.address.row - b.address.row;
      return a.address.col - b.address.col;
    });

    return result;
  }

  /**
   * Find next/previous cell with comments (for navigation)
   */
  getNextCommentCell(fromAddr: Address, direction: 'next' | 'prev' = 'next'): Address | null {
    const all = this.getAllComments(); // already sorted by (row, col) numerically
    if (all.length === 0) return null;

    if (direction === 'next') {
      for (const item of all) {
        const { row, col } = item.address;
        // Strict numeric comparison — string key comparison was broken for row≥10
        if (row > fromAddr.row || (row === fromAddr.row && col > fromAddr.col)) {
          return item.address;
        }
      }
      return all[0].address; // wrap around
    } else {
      for (let i = all.length - 1; i >= 0; i--) {
        const { row, col } = all[i].address;
        if (row < fromAddr.row || (row === fromAddr.row && col < fromAddr.col)) {
          return all[i].address;
        }
      }
      return all[all.length - 1].address; // wrap around
    }
  }

  // ==================== Icon APIs ====================

  /**
   * Set an icon for a cell
   */
  setIcon(addr: Address, icon: CellIcon | undefined): void {
    const c = this.cells.getOrCreate(addr.row, addr.col);
    c.icon = icon;
    this.events.emit({ type: 'icon-changed', address: addr, icon });
  }

  /**
   * Get icon for a cell
   */
  getIcon(addr: Address): CellIcon | undefined {
    return this.getCell(addr)?.icon;
  }

  /**
   * Get all cells with icons
   */
  getAllIcons(): Array<{ address: Address; icon: CellIcon }> {
    const result: Array<{ address: Address; icon: CellIcon }> = [];

    this.cells.forEach((row, col, cell) => {
      if (cell.icon) result.push({ address: { row, col }, icon: cell.icon });
    });

    return result;
  }

  // ==================== Custom Cell Component APIs ====================

  setCellComponent(addr: Address, component: CustomCellComponent | undefined): void {
    const c = this.cells.getOrCreate(addr.row, addr.col);
    c.customComponent = component;
    this.events.emit({ type: 'cell-component-changed', address: addr, component });
  }

  getCellComponent(addr: Address): CustomCellComponent | undefined {
    return this.getCell(addr)?.customComponent;
  }

  clearCellComponent(addr: Address): void {
    this.setCellComponent(addr, undefined);
  }

  getAllCellComponents(): Array<{ address: Address; component: CustomCellComponent }> {
    const result: Array<{ address: Address; component: CustomCellComponent }> = [];

    this.cells.forEach((row, col, cell) => {
      if (cell.customComponent) {
        result.push({ address: { row, col }, component: cell.customComponent });
      }
    });

    result.sort((a, b) => {
      if (a.address.row !== b.address.row) return a.address.row - b.address.row;
      return a.address.col - b.address.col;
    });

    return result;
  }

  // ==================== General Search API (Phase 1) ====================

  /**
   * Return the addresses of every cell that has a formula string set
   * (`cell.formula != null`).
   *
   * Implementation: iterates the cell store (O(n_cells_in_store)) but skips
   * every cell without a formula, so the effective cost is O(f) where f is the
   * formula-cell count.  This covers all formula cells regardless of whether
   * they were registered with the dependency graph (cells with no refs, e.g.
   * `=NOW()`, or formula cells created without a live formula engine are still
   * found correctly).
   *
   * Used by the `./search` subpath (`findInFormulas`, `replaceInFormulas`).
   * Cells with plain values (no `.formula` field) are never included.
   *
   * Result order matches cell-store insertion order; callers must sort if needed.
   */
  getFormulaAddresses(): Address[] {
    const result: Address[] = [];
    this.cells.forEach((row, col, cell) => {
      if (cell.formula != null) result.push({ row, col });
    });
    return result;
  }

  /**
   * Return the addresses of cells this formula cell directly depends on
   * (its precedents — the cells it reads).
   *
   * If the cell has no formula registered in the dependency graph,
   * an empty array is returned (not an error).
   *
   * O(e) where e = number of edges from this node.
   */
  getPrecedents(addr: Address): Address[] {
    return this.recalcCoordinator.graph.getDependencies(addr.row, addr.col);
  }

  /**
   * Return the addresses of formula cells that directly depend on this cell
   * (its dependents — the cells that read it).
   *
   * O(e) where e = number of edges into this node.
   */
  getDependents(addr: Address): Address[] {
    return this.recalcCoordinator.graph.getDependents(addr.row, addr.col);
  }

  // ── Phase 18: Go To Special ──────────────────────────────────────────────

  /**
   * Go To Special — return all addresses matching the given special-cell
   * criteria, optionally constrained to `range`.
   *
   * Implements the Excel "Go To Special" dialog (Ctrl+G → Special) across
   * all 14 `SpecialCellType` variants defined in `search-types.ts`.
   *
   * ─── Supported types ────────────────────────────────────────────────────
   *   formulas       — cells with a formula string (+ optional value filter)
   *   constants      — non-formula cells with a non-null value (+ optional filter)
   *   blanks         — empty cells within the search range
   *   errors         — formula or value cells displaying an Excel error string
   *   visible        — non-hidden cells (row and column both visible)
   *   lastCell       — the bottom-right corner of the used range
   *   currentRegion  — contiguous block around `anchor` (Ctrl+*)
   *   currentArray   — the spill range anchored at `anchor`
   *   precedents     — direct precedents of `anchor`
   *   allPrecedents  — transitive closure of precedents
   *   dependents     — direct dependents of `anchor`
   *   allDependents  — transitive closure of dependents
   *   conditionalFormats — populated cells covered by ≥1 CF rule range
   *   dataValidation — always empty (not yet implemented in kernel)
   *
   * @param options  What to find; must include `type`.
   * @param range    Optional bounding box to restrict the search.
   * @returns        Sorted address array (row-major), possibly empty.
   */
  getSpecialCells(options: SpecialCellsOptions, range?: SearchRange): Address[] {
    const { type, value: valueFilter, anchor } = options;

    // ── helpers ─────────────────────────────────────────────────────────────

    /** True if `addr` falls within range `r` (inclusive). */
    const inRange = (addr: Address, r: { start: Address; end: Address }): boolean =>
      addr.row >= r.start.row && addr.row <= r.end.row &&
      addr.col >= r.start.col && addr.col <= r.end.col;

    /** True if `v` is an Excel error string like "#VALUE!" */
    const isExcelError = (v: unknown): boolean =>
      typeof v === 'string' && /^#[A-Z/0-9!?]+$/.test(v) && v.startsWith('#');

    /** True if `v` matches the optional SpecialCellValue filter. */
    const passesFilter = (v: unknown, f: SpecialCellValue | undefined): boolean => {
      if (!f) return true;
      switch (f) {
        case 'numbers':  return typeof v === 'number';
        case 'text':     return typeof v === 'string' && !isExcelError(v);
        case 'logicals': return typeof v === 'boolean';
        case 'errors':   return isExcelError(v);
      }
    };

    // ── dispatch ─────────────────────────────────────────────────────────────

    switch (type) {
      case 'formulas': {
        const result: Address[] = [];
        this.cells.forEach((row, col, cell) => {
          if (cell.formula == null) return;
          if (range && !inRange({ row, col }, range)) return;
          if (!passesFilter(cell.value, valueFilter)) return;
          result.push({ row, col });
        });
        return result.sort(compareRowMajor);
      }

      case 'constants': {
        const result: Address[] = [];
        this.cells.forEach((row, col, cell) => {
          if (cell.formula != null) return;
          if (cell.value === null || cell.value === undefined) return;
          if (range && !inRange({ row, col }, range)) return;
          if (!passesFilter(cell.value, valueFilter)) return;
          result.push({ row, col });
        });
        return result.sort(compareRowMajor);
      }

      case 'blanks': {
        // Blank = no cell entry, OR cell exists but value is null/'' and no formula.
        // Requires a bounded range; fall back to used range if none supplied.
        const searchRange = range ?? this.getUsedRange();
        if (!searchRange) return [];
        const { start, end } = searchRange;
        const result: Address[] = [];
        for (let r = start.row; r <= end.row; r++) {
          for (let c = start.col; c <= end.col; c++) {
            const cell = this.getCell({ row: r, col: c });
            if (!cell || (cell.formula == null && (cell.value === null || cell.value === ''))) {
              result.push({ row: r, col: c });
            }
          }
        }
        return result;
      }

      case 'visible': {
        const result: Address[] = [];
        this.cells.forEach((row, col, cell) => {
          if (this.isRowHidden(row) || this.isColHidden(col)) return;
          if (range && !inRange({ row, col }, range)) return;
          if (cell.value === null && cell.formula == null) return; // blank
          result.push({ row, col });
        });
        return result.sort(compareRowMajor);
      }

      case 'lastCell': {
        const ur = this.getUsedRange();
        return ur ? [ur.end] : [];
      }

      case 'currentRegion': {
        if (!anchor) return [];
        const region = this.getContiguousRange(anchor);
        if (!region) return this.getCell(anchor) ? [anchor] : [];
        const result: Address[] = [];
        this.cells.forEach((row, col) => {
          if (inRange({ row, col }, region)) result.push({ row, col });
        });
        return result.sort(compareRowMajor);
      }

      case 'currentArray': {
        if (!anchor) return [];
        const cell = this.getCell(anchor);
        if (!cell) return [];
        // Anchor is the spill source
        if (cell.spillSource) {
          const { endAddress } = cell.spillSource;
          const result: Address[] = [];
          for (let r = anchor.row; r <= endAddress.row; r++)
            for (let c = anchor.col; c <= endAddress.col; c++)
              result.push({ row: r, col: c });
          return result;
        }
        // Anchor is inside a spilled range — delegate to source
        if (cell.spilledFrom) {
          return this.getSpecialCells({ type: 'currentArray', anchor: cell.spilledFrom });
        }
        return [anchor];
      }

      case 'precedents':    return anchor ? this.getPrecedents(anchor)  : [];
      case 'dependents':    return anchor ? this.getDependents(anchor)  : [];

      case 'allPrecedents': {
        if (!anchor) return [];
        const visited = new Set<string>();
        const queue: Address[] = [anchor];
        const result: Address[] = [];
        while (queue.length) {
          const curr = queue.shift()!;
          for (const p of this.getPrecedents(curr)) {
            const k = `${p.row}:${p.col}`;
            if (visited.has(k)) continue;
            visited.add(k);
            result.push(p);
            queue.push(p);
          }
        }
        return result.sort(compareRowMajor);
      }

      case 'allDependents': {
        if (!anchor) return [];
        const visited = new Set<string>();
        const queue: Address[] = [anchor];
        const result: Address[] = [];
        while (queue.length) {
          const curr = queue.shift()!;
          for (const d of this.getDependents(curr)) {
            const k = `${d.row}:${d.col}`;
            if (visited.has(k)) continue;
            visited.add(k);
            result.push(d);
            queue.push(d);
          }
        }
        return result.sort(compareRowMajor);
      }

      case 'conditionalFormats': {
        const rules = this.getConditionalFormattingRules();
        if (rules.length === 0) return [];
        const seen = new Set<string>();
        const result: Address[] = [];
        this.cells.forEach((row, col) => {
          if (range && !inRange({ row, col }, range)) return;
          for (const rule of rules) {
            if (!rule.ranges || rule.ranges.length === 0) continue;
            for (const rng of rule.ranges) {
              if (inRange({ row, col }, rng)) {
                const k = `${row}:${col}`;
                if (!seen.has(k)) { seen.add(k); result.push({ row, col }); }
                return; // don't double-count this cell
              }
            }
          }
        });
        return result.sort(compareRowMajor);
      }

      case 'dataValidation': {
        const allValidation = this.getValidationCells();
        if (!range) return allValidation;
        return allValidation.filter(
          a => a.row >= range.start.row && a.row <= range.end.row &&
               a.col >= range.start.col && a.col <= range.end.col
        );
      }

      default:
        return [];
    }
  }

  /**
   * Find iterator - lazy search with generator pattern
   * 
   * **Excel Parity**: ✅ Range.Find() with findNext loop pattern
   * 
   * **Complexity**: O(n) where n = cells in range
   * **Memory**: O(1) - yields addresses one at a time
   * **Precision**: ±0 (exact text match)
   * **Error Strategy**: SKIP_ERRORS (silently skip malformed cells)
   * **Volatility**: Non-volatile (deterministic for given worksheet state)
   * 
   * **Phase 1 Implementation Status**: 🚧 STUB - Returns empty generator
   * 
   * @param options - Search configuration (what, lookIn, lookAt, matchCase, etc.)
   * @param range - Optional range to search (default: entire worksheet)
   * @yields Address of each matching cell
   * 
   * @example
   * ```ts
   * // Find all cells containing "Apple"
   * for (const addr of sheet.findIterator({ what: "Apple" })) {
   *   console.log(addr); // { row: 5, col: 2 }
   * }
   * 
   * // Case-sensitive search in formulas
   * for (const addr of sheet.findIterator({ 
   *   what: "SUM", 
   *   lookIn: "formulas", 
   *   matchCase: true 
   * })) {
   *   console.log(sheet.getCell(addr)?.formula);
   * }
   * ```
   */
  *findIterator(
    options: SearchOptions,
    range?: SearchRange
  ): Generator<Address, void, undefined> {
    // Empty pattern + no format query → matches nothing (Excel semantics).
    const hasTextQuery = options.what !== '';
    const hasFormatQuery = !!(options.searchFormat && Object.keys(options.searchFormat).length > 0);
    if (!hasTextQuery && !hasFormatQuery) return;

    const { lookIn = 'values', searchOrder = 'rows', includeHidden = false } = options;
    const matcher = hasTextQuery ? buildMatcher(options) : null;

    // ── 1. Collect addresses of all populated cells ──────────────────────────
    // forEach visits only cells that exist in the store; empty slots are skipped.
    // Non-anchor merged cells are never in the ICellStore (cleared at merge time),
    // but we add an explicit guard as a safety net for any edge case.
    const addresses: Address[] = [];
    this.cells.forEach((row, col) => {
      // Safety net: skip non-anchor merged cells.
      // In normal operation these are absent from ICellStore, but guard defensively.
      if (this.mergeStore.isNonAnchor(row, col)) return;

      // ── Phase 3: skip hidden cells ────────────────────────────────────────
      // Default: skip cells whose row OR column is hidden (Excel Ctrl+F parity).
      // Pass includeHidden: true to search the raw data model.
      if (!includeHidden && this.visibilityStore.isCellHidden(row, col)) return;

      // ── 2. Apply range filter ─────────────────────────────────────────────
      if (range) {
        if (row < range.start.row || row > range.end.row) return;
        if (col < range.start.col || col > range.end.col) return;
      }
      addresses.push({ row, col });
    });

    // ── 3. Sort by search order ─────────────────────────────────────────────
    addresses.sort(searchOrder === 'columns' ? compareColMajor : compareRowMajor);

    // ── 4. Yield matches ─────────────────────────────────────────────────────
    for (const addr of addresses) {
      const cell = this.cells.get(addr.row, addr.col);
      if (!cell) continue;

      let text: string | null = null;

      if (lookIn === 'values') {
        text = cellValueToString(cell.value);
      } else if (lookIn === 'formulas') {
        // Search the formula string if present; fall back to display value.
        text = cell.formula != null ? cell.formula : cellValueToString(cell.value);
      } else if (lookIn === 'comments') {
        if (cell.comments && cell.comments.length > 0) {
          text = cell.comments.map(c => c.text).join(' ');
        }
      }

      // ── Format filter ─────────────────────────────────────────────────
      if (hasFormatQuery && !styleMatchesFormat(cell.style, options.searchFormat!)) continue;

      // ── Text filter ───────────────────────────────────────────────────
      if (hasTextQuery) {
        if (text === null) continue;
        if (!matcher!(text)) continue;
      }

      yield addr;
    }
  }

  /**
   * Find single match (convenience wrapper)
   * 
   * **Excel Parity**: ✅ Range.Find(what, after)
   * 
   * **Complexity**: O(n) worst case, but returns first match
   * **Memory**: O(1)
   * **Precision**: ±0
   * **Error Strategy**: SKIP_ERRORS
   * **Volatility**: Non-volatile
   * 
   * **Phase 1 Implementation Status**: 🚧 STUB - Returns null
   * 
   * @param options - Search configuration
   * @param after - Start search after this address (default: top-left)
   * @param range - Optional range to search
   * @returns First matching address or null if no match
   * 
   * @example
   * ```ts
   * const addr = sheet.find({ what: "Apple", matchCase: false });
   * if (addr) {
   *   console.log(`Found at ${addr.row}, ${addr.col}`);
   * }
   * 
   * // Find next occurrence after A5
   * const nextAddr = sheet.find({ what: "Apple" }, { row: 4, col: 0 });
   * ```
   */
  find(
    options: SearchOptions,
    after?: Address,
    range?: SearchRange
  ): Address | null {
    const { searchDirection = 'next' } = options;

    // Eagerly collect all matches from the iterator (O(n) scan).
    // PM directive: correctness first, no premature optimisation.
    const allMatches = [...this.findIterator(options, range)];
    if (allMatches.length === 0) return null;

    if (searchDirection === 'next') {
      // ── Forward search ──────────────────────────────────────────────────
      if (!after) return allMatches[0];

      // Find first match strictly after `after` in iteration order.
      for (const addr of allMatches) {
        if (isStrictlyAfterRowMajor(addr, after)) return addr;
      }

      // Wrap-around: no match after `after` → return from the beginning.
      return allMatches[0];
    } else {
      // ── Backward search (previous) ───────────────────────────────────────
      // Phase 1: row-major backward only (column-major backward deferred).
      if (!after) return allMatches[allMatches.length - 1];

      // Scan in reverse; return first match strictly before `after`.
      for (let i = allMatches.length - 1; i >= 0; i--) {
        if (isStrictlyBeforeRowMajor(allMatches[i], after)) return allMatches[i];
      }

      // Wrap-around: no match before `after` → return from the end.
      return allMatches[allMatches.length - 1];
    }
  }

  /**
   * Find all matches (eager collection)
   * 
   * **Excel Parity**: ✅ Range.FindAll() pattern (loop + collect)
   * 
   * **Complexity**: O(n) where n = cells in range
   * **Memory**: O(m) where m = number of matches (eager allocation)
   * **Precision**: ±0
   * **Error Strategy**: SKIP_ERRORS
   * **Volatility**: Non-volatile
   * 
   * **Phase 1 Implementation Status**: 🚧 STUB - Returns empty array
   * 
   * **Warning**: Use sparingly for large worksheets. Prefer `findIterator()` for
   * memory-efficient streaming. This method allocates an array upfront.
   * 
   * @param options - Search configuration
   * @param range - Optional range to search
   * @returns Array of all matching addresses (empty if no matches)
   * 
   * @example
   * ```ts
   * // Find all "Apple" occurrences
   * const matches = sheet.findAll({ what: "Apple" });
   * console.log(`Found ${matches.length} matches`);
   * 
   * // Search in specific range
   * const rangeMatches = sheet.findAll(
   *   { what: "error", matchCase: false },
   *   { start: { row: 0, col: 0 }, end: { row: 100, col: 10 } }
   * );
   * ```
   */
  findAll(
    options: SearchOptions,
    range?: SearchRange
  ): Address[] {
    // findAll is a thin eager wrapper over the lazy iterator.
    // PM mandate: no independent scan loop here — must go through findIterator.
    return [...this.findIterator(options, range)];
  }

  // ==================== Snapshot API (Phase 7) ====================

  /**
   * Extract all serialisable Worksheet state into a plain object.
   *
   * Captures: cells, merges, visibility (hidden rows/cols), DAG dependency edges,
   * and volatile registrations.  Does NOT capture: column widths, row heights,
   * column filters, conditional formatting rules, or formula engine reference.
   *
   * Typical usage:
   * ```ts
   * import { SnapshotCodec } from './persistence/SnapshotCodec';
   * const buf = new SnapshotCodec().encode(ws.extractSnapshot());
   * ```
   *
   * @complexity O(V + E) where V = cells + merges + visibility items, E = DAG edges.
   */
  extractSnapshot(): WorksheetSnapshot {
    const cells: WorksheetSnapshot['cells'] = [];
    this.cells.forEach((row, col, cell) => cells.push({ row, col, cell }));

    const dagEdges: WorksheetSnapshot['dagEdges'] = [];
    this.recalcCoordinator.forEachFormula((row, col, deps) => {
      dagEdges.push({ row, col, deps });
    });

    return {
      version:    FORMAT_VERSION,
      cells,
      merges:     this.mergeStore.getAll(),
      hiddenRows: [...this.visibilityStore.getHiddenRows()],
      hiddenCols: [...this.visibilityStore.getHiddenCols()],
      dagEdges,
      volatiles:  this.recalcCoordinator.getVolatileAddresses(),
    };
  }

  /**
   * Replace all Worksheet state with the contents of a snapshot.
   *
   * Clears the current cell store, merge store, visibility store, and DAG,
   * then loads each section from the snapshot.  The formula engine reference,
   * column widths, row heights, column filters, and conditional formatting
   * rules are NOT touched — they survive the restore.
   *
   * @param snapshot  Plain WorksheetSnapshot produced by extractSnapshot() or
   *                  decoded by SnapshotCodec.decode().
   */
  applySnapshot(snapshot: WorksheetSnapshot): void {
    // ── Reset all stores ──────────────────────────────────────────────────
    this.cells          = new ColumnarCellStore();
    this.mergeStore     = new MergeStoreV1();
    this.visibilityStore = new VisibilityStoreV1();
    this.dag.clearAll();

    // ── Restore cells ────────────────────────────────────────────────────
    for (const { row, col, cell } of snapshot.cells) {
      this.cells.set(row, col, cell);
    }

    // ── Restore merges ───────────────────────────────────────────────────
    for (const region of snapshot.merges) {
      this.mergeStore.add(region);
    }

    // ── Restore visibility ───────────────────────────────────────────────
    for (const row of snapshot.hiddenRows) this.visibilityStore.hideRow(row);
    for (const col of snapshot.hiddenCols) this.visibilityStore.hideCol(col);

    // ── Restore DAG ──────────────────────────────────────────────────────
    for (const { row, col, deps } of snapshot.dagEdges) {
      this.recalcCoordinator.registerFormula(row, col, deps);
    }

    // ── Restore volatiles ────────────────────────────────────────────────
    for (const { row, col } of snapshot.volatiles) {
      // Re-register with empty local deps list; the DAG edges carry structural deps.
      this.recalcCoordinator.registerVolatile(row, col, []);
    }
  }

  // ==================== Private Helpers ====================

  private generateCommentId(): string {
    return `comment_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
}
