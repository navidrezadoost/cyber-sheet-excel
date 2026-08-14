/**
 * SpreadsheetSDK — stable public API facade for @cyber-sheet/core.
 *
 * Design principles:
 *  - Framework-agnostic: no React/Vue/Angular imports
 *  - Synchronous: no Promises in the hot path
 *  - Minimal surface: only the operations that belong on a public contract
 *  - Typed errors: every failure mode is represented with a concrete class
 *  - Tree-shakeable: this module has no side effects
 *
 * Usage:
 *  ```ts
 *  import { createSpreadsheet } from '@cyber-sheet/core/sdk';
 *
 *  const sheet = createSpreadsheet('Sheet1');
 *  sheet.setCell(1, 1, 'Hello');
 *  sheet.setCell(1, 2, 42);
 *  const snap = sheet.snapshot();
 *  sheet.dispose();
 *  ```
 */

import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Worksheet } from '../worksheet';
import { snapshotCodec } from '../persistence/SnapshotCodec';
import type { WorksheetSnapshot } from '../persistence/SnapshotCodec';
import type { Cell, CellStyle, ExtendedCellValue, Range, Address, DataValidationRule, SheetProtectionOptions, FreezeState, ColumnFilter, SortKey, AutoFilterRange } from '../types';
import type { Disposable } from '../events';
import type { WorksheetPatch } from '../patch/WorksheetPatch';
import { PatchOps } from '../patch/WorksheetPatch';
import { SyncUndoStack } from './SyncUndoStack';
import { recordingApplyPatch } from '../patch/PatchRecorder';
import { buildPivot, pivotGridToValues } from './pivot';
import type { PivotDefinition, PivotGrid } from './pivot';

// ---------------------------------------------------------------------------
// Typed errors — defined in errors.ts, re-exported from here for back-compat.
// ---------------------------------------------------------------------------

import {
  SdkError,
  DisposedError,
  BoundsError,
  SnapshotError,
  MergeError,
  PatchError,
  ProtectedCellError,
  ProtectedSheetOperationError,
  ValidationError,
  PatchRecorderError,
  UndoError,
  PivotSourceError,
  PivotFieldError,
  EmptyPivotSourceError,
} from './errors';

export {
  SdkError,
  DisposedError,
  BoundsError,
  SnapshotError,
  MergeError,
  PatchError,
  ProtectedCellError,
  ProtectedSheetOperationError,
  ValidationError,
  PatchRecorderError,
  UndoError,
  // Phase 25 — pivot errors
  PivotSourceError,
  PivotFieldError,
  EmptyPivotSourceError,
} from './errors';
// Phase 25 — pivot types re-exported on the public surface
export type { PivotDefinition, PivotValueSpec, PivotGrid, PivotGridRow, PivotAggregator } from './pivot';
// ---------------------------------------------------------------------------

/** All event types emitted by SpreadsheetSDK. */
export type SdkEventType =
  | 'cell-changed'
  | 'style-changed'
  | 'structure-changed'
  | 'filter-changed'
  | 'sort-applied'
  | 'cycle-detected';

/**
 * Payload emitted for SDK events.
 *
 * `structure-changed` covers merges, visibility changes, and `sheet-mutated`.
 * UI adapters should re-render their viewport on any `structure-changed`.
 */
export type SdkEvent =
  | { type: 'cell-changed';    row: number; col: number }
  | { type: 'style-changed';   row: number; col: number }
  | { type: 'structure-changed' }
  | { type: 'filter-changed';  col: number }
  | { type: 'sort-applied';    startRow: number; startCol: number; endRow: number; endCol: number }
  | { type: 'cycle-detected' };

export type SdkEventListener = (event: SdkEvent) => void;

// ---------------------------------------------------------------------------
// Mutation trace hook — Phase 24
// ---------------------------------------------------------------------------

/**
 * Diagnostic payload emitted after every tracked write operation.
 *
 * `operation` — name of the SDK method (e.g. `'setCell'`, `'undo'`).
 * `durationMs` — wall-clock time spent inside the operation, in milliseconds.
 * `timestamp`  — Unix epoch in milliseconds at the start of the operation.
 */
export type MutationTraceEvent = {
  operation: string;
  durationMs: number;
  timestamp: number;
};

/**
 * Callback invoked synchronously after every tracked write operation.
 * The hook must not throw — exceptions are swallowed to protect callers.
 */
export type MutationTraceHook = (event: MutationTraceEvent) => void;

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

export type SpreadsheetOptions = {
  /**
   * Maximum number of entries in the undo/redo stack.
   * @default 100
   */
  maxUndoHistory?: number;
  /**
   * Initial row count.
   * @default 5000
   */
  rows?: number;
  /**
   * Initial column count.
   * @default 16384
   */
  cols?: number;
};

// ---------------------------------------------------------------------------
// Public interface
// ---------------------------------------------------------------------------

/**
 * The stable public contract for a spreadsheet sheet.
 * All implementations must satisfy this interface exactly.
 */
export interface SpreadsheetSDK {
  // ── Metadata ─────────────────────────────────────────────────────────────
  /** The name of this sheet (set at creation time). */
  readonly name: string;
  /** Total number of rows in this sheet (1-based max row index). */
  readonly rowCount: number;
  /** Total number of columns in this sheet (1-based max col index). */
  readonly colCount: number;

  // ── Cell access (1-based row/col) ─────────────────────────────────────────
  /**
   * Set the value of a cell.
   * Row and col are 1-based integers.
   */
  setCell(row: number, col: number, value: ExtendedCellValue): void;

  /**
   * Return the raw `Cell` record at (row, col), or `undefined` if the cell
   * has never been written.
   */
  getCell(row: number, col: number): Cell | undefined;

  /**
   * Return only the cell's value, or `null` if the cell is empty.
   */
  getCellValue(row: number, col: number): ExtendedCellValue;

  // ── Patch API ─────────────────────────────────────────────────────────────
  /**
   * Apply a `WorksheetPatch` and return the **inverse** patch.
   * The inverse can be passed back to `applyPatch` to undo the operation.
   *
   * This call **does not** push to the undo stack — use `setCell` / the
   * mutation helpers for undo-tracked mutations.
   */
  applyPatch(patch: WorksheetPatch): WorksheetPatch;

  // ── Snapshot ──────────────────────────────────────────────────────────────
  /**
   * Extract a plain-object snapshot of the entire sheet state.
   * Safe to JSON-serialize; does not reference internal mutable structures.
   */
  snapshot(): WorksheetSnapshot;

  /**
   * Replace all sheet state with the contents of a previously-captured
   * snapshot. Clears the undo/redo stack.
   * @throws SnapshotError if the snapshot is structurally invalid.
   */
  restore(snapshot: WorksheetSnapshot): void;

  /**
   * Encode the current sheet state to a compact binary buffer.
   * Format: CSEX magic bytes + u16 version + LZ4-compressed JSON.
   */
  encodeSnapshot(): Uint8Array;

  /**
   * Decode a binary snapshot buffer and restore the sheet state.
   * @throws SnapshotError if decoding fails or the version is unknown.
   */
  decodeAndRestore(bytes: Uint8Array): void;

  // ── Undo / Redo ───────────────────────────────────────────────────────────
  /** Undo the last tracked mutation. Returns `true` if an entry was undone. */
  undo(): boolean;
  /** Redo the last undone mutation. Returns `true` if an entry was redone. */
  redo(): boolean;
  /** Whether there is anything to undo. */
  readonly canUndo: boolean;
  /** Whether there is anything to redo. */
  readonly canRedo: boolean;

  // ── Merge ─────────────────────────────────────────────────────────────────
  /** Merge the rectangular region from (startRow, startCol) to (endRow, endCol). */
  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): void;
  /** Remove the merge covering the given anchor address. */
  cancelMerge(startRow: number, startCol: number, endRow: number, endCol: number): void;
  /** Return all currently merged regions as Range objects ({ start, end } addresses). */
  getMergedRanges(): Range[];
  /** Return `true` if (row, col) is within any merged region. */
  isInMerge(row: number, col: number): boolean;

  // ── Row / Column visibility ───────────────────────────────────────────────
  /** Hide a row. Idempotent. */
  hideRow(row: number): void;
  /** Show a hidden row. Idempotent. */
  showRow(row: number): void;
  /** Hide a column. Idempotent. */
  hideCol(col: number): void;
  /** Show a hidden column. Idempotent. */
  showCol(col: number): void;
  /** Return `true` if the row is hidden. */
  isRowHidden(row: number): boolean;
  /** Return `true` if the column is hidden. */
  isColHidden(col: number): boolean;

  // ── Data Validation ───────────────────────────────────────────────────────
  /** Attach a validation rule to the cell at (row, col). */
  setDataValidation(row: number, col: number, rule: DataValidationRule): void;
  /** Return the validation rule for the cell, or `undefined` if none. */
  getDataValidation(row: number, col: number): DataValidationRule | undefined;
  /** Remove any validation rule from the cell at (row, col). */
  removeDataValidation(row: number, col: number): void;
  /** Return addresses of all cells that have a data-validation rule (row-major order). */
  getValidationCells(): Address[];

  // ── Sheet Protection ───────────────────────────────────────────────────────
  /**
   * Enable sheet protection. While the sheet is protected, any cell whose
   * `style.locked` is not explicitly `false` (Excel default: locked = true)
   * will reject `setCell` / `lockCell` mutations with `ProtectedCellError`.
   *
   * @param options  Optional capability overrides (e.g. `allowFormatCells`).
   */
  setSheetProtection(options?: SheetProtectionOptions): void;
  /** Remove sheet protection. This is a no-op if the sheet is not protected. */
  removeSheetProtection(): void;
  /** Return `true` if the sheet is currently protected. */
  isSheetProtected(): boolean;
  /** Return the current protection settings, or `null` if not protected. */
  getSheetProtection(): SheetProtectionOptions | null;
  /**
   * Return `true` if the cell at (row, col) would currently block mutation.
   *
   * A cell is protected when: `isSheetProtected()` is true AND
   * `cell.style?.locked !== false` (Excel default: all cells are locked).
   */
  isCellProtected(row: number, col: number): boolean;
  /**
   * Explicitly lock a single cell (sets `style.locked = true`).
   * This change is tracked on the undo stack.
   */
  lockCell(row: number, col: number): void;
  /**
   * Explicitly unlock a single cell (sets `style.locked = false`).
   * This change is tracked on the undo stack.
   */
  unlockCell(row: number, col: number): void;

  /**
   * Return `true` if the cell at (row, col) has `style.locked !== false`,
   * regardless of whether the sheet is currently protected.
   *
   * Use `isCellProtected()` to check whether a mutation would be blocked.
   */
  isCellLocked(row: number, col: number): boolean;

  /**
   * Lock all cells in `range` (sets `style.locked = true` on each).
   * Recorded as a single undo entry.
   */
  lockCells(range: Range): void;

  /**
   * Unlock all cells in `range` (sets `style.locked = false` on each).
   * Recorded as a single undo entry.
   */
  unlockCells(range: Range): void;

  /**
   * Return the formula string for the cell at (row, col), or `null` if the
   * cell has no formula, has a plain value, or the formula is hidden
   * (sheet is protected and `cell.style.hidden === true`).
   */
  getFormula(row: number, col: number): string | null;

  // ── Freeze Panes ──────────────────────────────────────────────────────────
  /**
   * Freeze the top `rows` rows and the left `cols` columns.
   * Passing `rows=0, cols=0` is equivalent to `clearFreezePanes()`.
   * This change is tracked on the undo stack.
   */
  setFreezePanes(rows: number, cols: number): void;
  /** Remove all frozen panes. No-op if nothing is frozen. */
  clearFreezePanes(): void;
  /** Return the current freeze-pane state, or `null` if no panes are frozen. */
  getFreezePanes(): FreezeState | null;
  // ── Column Filters ─────────────────────────────────────────────────
  /**
   * Apply a filter to a column.  This change is tracked on the undo stack.
   *
   * Supported filter types: `equals`, `notEquals`, `contains`, `notContains`,
   * `startsWith`, `endsWith`, `gt`, `gte`, `lt`, `lte`, `between`,
   * `empty`, `notEmpty`, `in`.
   */
  setFilter(col: number, filter: ColumnFilter): void;
  /** Remove the filter on a column.  No-op if no filter is active.  Undoable. */
  clearFilter(col: number): void;
  /**
   * Remove all column filters at once.  No-op if no filters are active.  Undoable
   * as a single undo entry (restores all cleared filters simultaneously).
   */
  clearAllFilters(): void;
  /** Return the current filter for a column, or `undefined` if none. */
  getFilter(col: number): ColumnFilter | undefined;
  /**
   * Return the row indices (1-based) that pass ALL currently active filters.
   * Hidden rows are included — filter visibility is independent of row hide/show.
   *
   * @param range Optional sub-range; when omitted the full sheet row range is used.
   */
  getVisibleRows(range?: Range): number[];
  /**
   * Return distinct cell values for a column, together with their occurrence counts.
   * Useful for building a filter dropdown.
   *
   * @param col         1-based column index.
   * @param visibleOnly When `true`, only rows that pass all OTHER column filters
   *                    are considered (Excel’s behaviour for filter dropdowns).
   *                    Defaults to `false`.
   */
  getDistinctValues(col: number, visibleOnly?: boolean): { value: string; count: number }[];

  // ── Auto-Filter Range ─────────────────────────────────────────────
  /**
   * Mark the row and columns that carry the filter dropdown arrows.
   * This is purely a UI marker — it does not affect filter matching.
   * The change is tracked on the undo stack.
   */
  setAutoFilterRange(headerRow: number, startCol: number, endCol: number): void;
  /** Remove the auto-filter range marker.  No-op if none is set.  Undoable. */
  clearAutoFilterRange(): void;
  /** Return the current auto-filter range, or `null` if none is set. */
  getAutoFilterRange(): AutoFilterRange | null;

  // ── Range Sort ───────────────────────────────────────────────────────
  /**
   * Sort a rectangular range by one or more sort keys.
   *
   * - Sort is stable; equal rows preserve their original order.
   * - Cells outside the range are unaffected.
   * - The operation is fully undoable (pre-sort snapshot is captured).
   * - Emits a `sort-applied` SDK event.
   *
   * @param range  1-based inclusive rectangle to sort.
   * @param keys   Primary → secondary → … sort keys.
   */
  sortRange(range: Range, keys: SortKey[]): void;

  // ── Pivot ─────────────────────────────────────────────────────────────────
  /**
   * Build and write a pivot table to the worksheet starting at `target`.
   *
   * The first row of `definition.source` is treated as the header row.
   * Output is written as a plain value grid (header row + one row per group).
   * The entire operation is a single undo/redo entry.
   *
   * @param definition  Serializable pivot specification.
   * @param target      Top-left cell of the output region (1-based).
   * @returns           The computed `PivotGrid` (pure output — also returned for
   *                    inspection without requiring a second `buildPivot` call).
   *
   * @throws {PivotSourceError}      if the source range is invalid.
   * @throws {PivotFieldError}       if a field name is absent from the source headers.
   * @throws {EmptyPivotSourceError} if the source has no data rows.
   * @throws {ProtectedCellError}    if a target cell is locked and the sheet is protected.
   * @throws {DisposedError}         if called after `dispose()`.
   */
  createPivot(definition: PivotDefinition, target: Address): PivotGrid;

  // ── Events ────────────────────────────────────────────────────────────────
  /**
   * Subscribe to SDK events.
   * Returns a `Disposable` — call `dispose()` to unsubscribe.
   *
   * Events are dispatched **synchronously** during the mutation that caused
   * them. Listeners must not throw; errors in listeners are caught and logged.
   */
  on(event: SdkEventType, listener: SdkEventListener): Disposable;

  // ── Lifecycle ─────────────────────────────────────────────────────────────
  /**
   * Release all internal resources (event listeners, stores).
   * After calling `dispose()`, all further method calls throw `DisposedError`.
   */
  dispose(): void;

  // ── Diagnostics ──────────────────────────────────────────────────────────
  /**
   * Attach (or detach) a mutation trace hook.
   *
   * When set, the hook is called **synchronously** after every tracked write
   * (`setCell`, `applyPatch`, `undo`, `redo`, `sortRange`, …).
   * Pass `null` to remove the current hook.
   *
   * Useful for performance profiling and test assertions.
   */
  setMutationTraceHook(hook: MutationTraceHook | null): void;
}

// ---------------------------------------------------------------------------
// Internal implementation
// ---------------------------------------------------------------------------

/** Internal listener map type. */
type ListenerMap = Map<SdkEventType, Set<SdkEventListener>>;

class SpreadsheetV1 implements SpreadsheetSDK {
  /** Exposed only for testing infra — not part of the public interface. */
  readonly _ws: Worksheet;
  private readonly _undo: SyncUndoStack;
  private readonly _listeners: ListenerMap = new Map();
  private _disposed = false;
  private _traceHook: MutationTraceHook | null = null;

  get rowCount(): number { return this._ws.rowCount; }
  get colCount(): number { return this._ws.colCount; }
  private readonly _unsubscribe: Disposable;

  constructor(name: string, options: Required<SpreadsheetOptions>) {
    this._ws = new Worksheet(name, options.rows, options.cols);
    this._undo = new SyncUndoStack(options.maxUndoHistory);

    // Bridge internal SheetEvents → SdkEvents
    this._unsubscribe = this._ws.on((e) => {
      switch (e.type) {
        case 'cell-changed':
          this._emit('cell-changed', { type: 'cell-changed', row: e.address.row, col: e.address.col });
          this._emit('structure-changed', { type: 'structure-changed' });
          break;
        case 'style-changed':
          this._emit('style-changed', { type: 'style-changed', row: e.address.row, col: e.address.col });
          break;
        case 'sheet-mutated':
        case 'merge-added':
        case 'merge-removed':
        case 'row-hidden':
        case 'row-shown':
        case 'col-hidden':
        case 'col-shown':
        case 'sheet-protection-changed':
        case 'freeze-panes-changed':
        case 'autofilter-range-changed':
          this._emit('structure-changed', { type: 'structure-changed' });
          break;
        case 'filter-changed':
          this._emit('filter-changed', { type: 'filter-changed', col: e.col });
          this._emit('structure-changed', { type: 'structure-changed' });
          break;
        case 'sort-applied':
          this._emit('sort-applied', { type: 'sort-applied', startRow: e.startRow, startCol: e.startCol, endRow: e.endRow, endCol: e.endCol });
          this._emit('structure-changed', { type: 'structure-changed' });
          break;
        case 'cycle-detected':
          this._emit('cycle-detected', { type: 'cycle-detected' });
          break;
      }
    });
  }

  private _guard(method: string): void {
    if (this._disposed) throw new DisposedError(method);
  }

  private _checkBounds(row: number, col: number): void {
    if (!Number.isInteger(row) || row < 1 || row > this._ws.rowCount) {
      throw new BoundsError(`row ${row} out of range 1..${this._ws.rowCount}`);
    }
    if (!Number.isInteger(col) || col < 1 || col > this._ws.colCount) {
      throw new BoundsError(`col ${col} out of range 1..${this._ws.colCount}`);
    }
  }

  private _guardCell(row: number, col: number, method: string): void {
    if (this._ws.isSheetProtected()) {
      const cell = this._ws.getCell({ row, col });
      // Excel default: cells are locked unless explicitly set to false.
      if (cell?.style?.locked !== false) {
        throw new ProtectedCellError(row, col);
      }
    }
  }

  /**
   * Throws `ProtectedSheetOperationError` when the sheet is protected and
   * the operation-level flag is not enabled.
   *
   * @param flag     The `SheetProtectionOptions` key that permits this op.
   * @param method   Human-readable operation name for the error message.
   */
  private _guardSheetOp(flag: keyof SheetProtectionOptions, method: string): void {
    if (!this._ws.isSheetProtected()) return;
    const opts = this._ws.getSheetProtection() as SheetProtectionOptions;
    if (!opts[flag]) {
      throw new ProtectedSheetOperationError(method, flag);
    }
  }

  private _emit(type: SdkEventType, event: SdkEvent): void {
    const set = this._listeners.get(type);
    if (!set) return;
    for (const listener of set) {
      try { listener(event); } catch (err) {
        // Listeners must not throw. Log and continue so one bad listener
        // doesn't break the others.
        console.error('[SpreadsheetSDK] Uncaught error in event listener:', err);
      }
    }
  }

  /**
   * Execute `fn` and re-wrap any non-SdkError throw as a typed `SdkError`.
   * This ensures no internal implementation detail escapes the public surface.
   *
   * @param wrap  Factory that converts the raw error into a typed SdkError.
   */
  private _wrapMutation<T>(
    fn: () => T,
    wrap: (err: unknown) => SdkError,
  ): T {
    try {
      return fn();
    } catch (err) {
      if (err instanceof SdkError) throw err;
      throw wrap(err);
    }
  }

  /**
   * Execute `fn` and fire the mutation trace hook (if set) with timing info.
   * Guards and bounds checks should be performed BEFORE calling this.
   */
  private _tracedWrite<T>(operation: string, fn: () => T): T {
    if (!this._traceHook) return fn();
    const timestamp = Date.now();
    const t0 = performance.now();
    const result = fn();
    const durationMs = performance.now() - t0;
    try {
      this._traceHook({ operation, durationMs, timestamp });
    } catch {
      // Trace hooks must not crash the SDK.
    }
    return result;
  }

  // ── Diagnostics ──────────────────────────────────────────────────────────

  setMutationTraceHook(hook: MutationTraceHook | null): void {
    this._guard('setMutationTraceHook');
    this._traceHook = hook;
  }

  // ── Metadata ─────────────────────────────────────────────────────────────

  get name(): string { return this._ws.name; }

  // ── Cell access ───────────────────────────────────────────────────────────

  setCell(row: number, col: number, value: ExtendedCellValue): void {
    this._guard('setCell');
    this._checkBounds(row, col);
    this._guardCell(row, col, 'setCell');
    this._tracedWrite('setCell', () => this._wrapMutation(() => {
      const patch: WorksheetPatch = {
        seq: 0,
        ops: [{
          op: 'setCellValue' as const,
          row,
          col,
          before: this._ws.getCellValue({ row, col }),
          after: value,
        }],
      };
      this._undo.applyAndRecord(this._ws, patch);
    }, (err) => new PatchError(`setCell(${row},${col}) failed: ${(err as Error).message ?? err}`, err)));
  }

  getCell(row: number, col: number): Cell | undefined {
    this._guard('getCell');
    this._checkBounds(row, col);
    return this._ws.getCell({ row, col });
  }

  getCellValue(row: number, col: number): ExtendedCellValue {
    this._guard('getCellValue');
    this._checkBounds(row, col);
    return this._ws.getCellValue({ row, col });
  }

  // ── Patch API ─────────────────────────────────────────────────────────────

  applyPatch(patch: WorksheetPatch): WorksheetPatch {
    this._guard('applyPatch');
    return this._tracedWrite('applyPatch', () => this._wrapMutation(
      () => recordingApplyPatch(this._ws, patch),
      (err) => new PatchError(`applyPatch failed: ${(err as Error).message ?? err}`, err),
    ));
  }

  // ── Snapshot ──────────────────────────────────────────────────────────────

  snapshot(): WorksheetSnapshot {
    this._guard('snapshot');
    // Deep-copy the snapshot so mutations on the Worksheet after this call
    // do not retroactively affect the captured state.
    // WorksheetSnapshot is a plain JSON-serializable object by contract,
    // so a JSON round-trip is the safest cross-environment deep-clone.
    return JSON.parse(JSON.stringify(this._ws.extractSnapshot())) as WorksheetSnapshot;
  }

  restore(snap: WorksheetSnapshot): void {
    this._guard('restore');
    try {
      this._ws.applySnapshot(snap);
    } catch (err) {
      throw new SnapshotError('applySnapshot failed', err);
    }
    this._undo.clear();
  }

  encodeSnapshot(): Uint8Array {
    this._guard('encodeSnapshot');
    return this._wrapMutation(
      () => snapshotCodec.encode(this._ws.extractSnapshot()),
      (err) => new SnapshotError(`encode failed: ${(err as Error).message ?? err}`, err),
    );
  }

  decodeAndRestore(bytes: Uint8Array): void {
    this._guard('decodeAndRestore');
    let snap: WorksheetSnapshot;
    try {
      snap = snapshotCodec.decode(bytes);
    } catch (err) {
      throw new SnapshotError('decode failed', err);
    }
    this.restore(snap);
  }

  // ── Undo / Redo ───────────────────────────────────────────────────────────

  undo(): boolean {
    this._guard('undo');
    return this._tracedWrite('undo', () => this._wrapMutation(
      () => this._undo.undo(this._ws),
      (err) => new PatchError(`undo failed: ${(err as Error).message ?? err}`, err),
    ));
  }

  redo(): boolean {
    this._guard('redo');
    return this._tracedWrite('redo', () => this._wrapMutation(
      () => this._undo.redo(this._ws),
      (err) => new PatchError(`redo failed: ${(err as Error).message ?? err}`, err),
    ));
  }

  get canUndo(): boolean { return this._undo.canUndo; }
  get canRedo(): boolean { return this._undo.canRedo; }

  // ── Merge ─────────────────────────────────────────────────────────────────

  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): void {
    this._guard('mergeCells');
    this._guardSheetOp('allowFormatCells', 'mergeCells');
    this._wrapMutation(() => {
      const patch: WorksheetPatch = {
        seq: 0,
        ops: [{ op: 'mergeCells', startRow, startCol, endRow, endCol }],
      };
      this._undo.applyAndRecord(this._ws, patch);
    }, (err) => new MergeError(
      `mergeCells(${startRow},${startCol},${endRow},${endCol}) failed: ${(err as Error).message ?? err}`,
      err,
    ));
  }

  cancelMerge(startRow: number, startCol: number, endRow: number, endCol: number): void {
    this._guard('cancelMerge');
    this._wrapMutation(() => {
      const patch: WorksheetPatch = {
        seq: 0,
        ops: [{ op: 'cancelMerge', startRow, startCol, endRow, endCol }],
      };
      this._undo.applyAndRecord(this._ws, patch);
    }, (err) => new MergeError(
      `cancelMerge(${startRow},${startCol},${endRow},${endCol}) failed: ${(err as Error).message ?? err}`,
      err,
    ));
  }

  getMergedRanges(): Range[] {
    this._guard('getMergedRanges');
    return this._ws.getMergedRanges();
  }

  isInMerge(row: number, col: number): boolean {
    this._guard('isInMerge');
    return this._ws.isInMerge({ row, col });
  }

  // ── Row / Column visibility ───────────────────────────────────────────────

  hideRow(row: number): void {
    this._guard('hideRow');
    this._guardSheetOp('allowFormatRows', 'hideRow');
    const patch: WorksheetPatch = { seq: 0, ops: [{ op: 'hideRow', row }] };
    this._undo.applyAndRecord(this._ws, patch);
  }

  showRow(row: number): void {
    this._guard('showRow');
    this._guardSheetOp('allowFormatRows', 'showRow');
    const patch: WorksheetPatch = { seq: 0, ops: [{ op: 'showRow', row }] };
    this._undo.applyAndRecord(this._ws, patch);
  }

  hideCol(col: number): void {
    this._guard('hideCol');
    this._guardSheetOp('allowFormatColumns', 'hideCol');
    const patch: WorksheetPatch = { seq: 0, ops: [{ op: 'hideCol', col }] };
    this._undo.applyAndRecord(this._ws, patch);
  }

  showCol(col: number): void {
    this._guard('showCol');
    this._guardSheetOp('allowFormatColumns', 'showCol');
    const patch: WorksheetPatch = { seq: 0, ops: [{ op: 'showCol', col }] };
    this._undo.applyAndRecord(this._ws, patch);
  }

  isRowHidden(row: number): boolean {
    this._guard('isRowHidden');
    return this._ws.isRowHidden(row);
  }

  isColHidden(col: number): boolean {
    this._guard('isColHidden');
    return this._ws.isColHidden(col);
  }

  // ── Data Validation ───────────────────────────────────────────────────────

  setDataValidation(row: number, col: number, rule: DataValidationRule): void {
    this._guard('setDataValidation');
    this._guardSheetOp('allowFormatCells', 'setDataValidation');
    this._ws.setDataValidation({ row, col }, rule);
  }

  getDataValidation(row: number, col: number): DataValidationRule | undefined {
    this._guard('getDataValidation');
    return this._ws.getDataValidation({ row, col });
  }

  removeDataValidation(row: number, col: number): void {
    this._guard('removeDataValidation');
    this._ws.removeDataValidation({ row, col });
  }

  getValidationCells(): Address[] {
    this._guard('getValidationCells');
    return this._ws.getValidationCells();
  }

  // ── Sheet Protection ───────────────────────────────────────────────────────

  setSheetProtection(options: SheetProtectionOptions = {}): void {
    this._guard('setSheetProtection');
    const before = this._ws.getSheetProtection();
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setSheetProtection' as const, before, after: options }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  removeSheetProtection(): void {
    this._guard('removeSheetProtection');
    const before = this._ws.getSheetProtection();
    if (before === null) return;
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setSheetProtection' as const, before, after: null }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  isSheetProtected(): boolean {
    this._guard('isSheetProtected');
    return this._ws.isSheetProtected();
  }

  getSheetProtection(): SheetProtectionOptions | null {
    this._guard('getSheetProtection');
    return this._ws.getSheetProtection();
  }

  isCellProtected(row: number, col: number): boolean {
    this._guard('isCellProtected');
    this._checkBounds(row, col);
    if (!this._ws.isSheetProtected()) return false;
    const cell = this._ws.getCell({ row, col });
    return cell?.style?.locked !== false;
  }

  lockCell(row: number, col: number): void {
    this._guard('lockCell');
    this._checkBounds(row, col);
    this._guardSheetOp('allowFormatCells', 'lockCell');
    const before = this._ws.getDirectCellStyle({ row, col });
    const after: CellStyle = { ...before, locked: true };
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setCellStyle' as const, row, col, before, after }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  unlockCell(row: number, col: number): void {
    this._guard('unlockCell');
    this._checkBounds(row, col);
    this._guardSheetOp('allowFormatCells', 'unlockCell');
    const before = this._ws.getDirectCellStyle({ row, col });
    const after: CellStyle = { ...before, locked: false };
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setCellStyle' as const, row, col, before, after }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  isCellLocked(row: number, col: number): boolean {
    this._guard('isCellLocked');
    this._checkBounds(row, col);
    const cell = this._ws.getCell({ row, col });
    // Excel default: a cell is locked unless style.locked is explicitly false.
    return cell?.style?.locked !== false;
  }

  lockCells(range: Range): void {
    this._guard('lockCells');
    this._guardSheetOp('allowFormatCells', 'lockCells');
    const ops: WorksheetPatch['ops'] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        this._checkBounds(r, c);
        const before = this._ws.getDirectCellStyle({ row: r, col: c });
        const after: CellStyle = { ...before, locked: true };
        ops.push({ op: 'setCellStyle' as const, row: r, col: c, before, after });
      }
    }
    if (ops.length > 0) this._undo.applyAndRecord(this._ws, { seq: 0, ops });
  }

  unlockCells(range: Range): void {
    this._guard('unlockCells');
    this._guardSheetOp('allowFormatCells', 'unlockCells');
    const ops: WorksheetPatch['ops'] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        this._checkBounds(r, c);
        const before = this._ws.getDirectCellStyle({ row: r, col: c });
        const after: CellStyle = { ...before, locked: false };
        ops.push({ op: 'setCellStyle' as const, row: r, col: c, before, after });
      }
    }
    if (ops.length > 0) this._undo.applyAndRecord(this._ws, { seq: 0, ops });
  }

  getFormula(row: number, col: number): string | null {
    this._guard('getFormula');
    this._checkBounds(row, col);
    const cell = this._ws.getCell({ row, col });
    if (cell == null || cell.formula == null) return null;
    // When the sheet is protected and the cell has style.hidden=true, the
    // formula is hidden (Excel "Hidden" protection attribute).
    if (this._ws.isSheetProtected() && cell.style?.hidden === true) return null;
    return cell.formula;
  }

  // ── Freeze Panes ──────────────────────────────────────────────────────────

  setFreezePanes(rows: number, cols: number): void {
    this._guard('setFreezePanes');
    const before = this._ws.getFreezePanes();
    const after: FreezeState | null = (rows === 0 && cols === 0) ? null : { rows, cols };
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setFreezePanes' as const, before, after }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  clearFreezePanes(): void {
    this._guard('clearFreezePanes');
    const before = this._ws.getFreezePanes();
    if (before === null) return;
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setFreezePanes' as const, before, after: null }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  getFreezePanes(): FreezeState | null {
    this._guard('getFreezePanes');
    return this._ws.getFreezePanes();
  }

  // ── Column Filters ────────────────────────────────────────────────────────

  setFilter(col: number, filter: ColumnFilter): void {
    this._guard('setFilter');
    this._guardSheetOp('allowFilter', 'setFilter');
    const before = this._ws.getColumnFilter(col) ?? null;
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setColumnFilter' as const, col, before, after: filter }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  clearFilter(col: number): void {
    this._guard('clearFilter');
    this._guardSheetOp('allowFilter', 'clearFilter');
    const before = this._ws.getColumnFilter(col) ?? null;
    if (before === null) return;
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setColumnFilter' as const, col, before, after: null }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  clearAllFilters(): void {
    this._guard('clearAllFilters');
    this._guardSheetOp('allowFilter', 'clearAllFilters');
    const all = this._ws.getAllFilters();
    if (all.size === 0) return;
    const ops = [...all.entries()].map(([col, before]) => ({
      op: 'setColumnFilter' as const, col, before, after: null,
    }));
    const patch: WorksheetPatch = { seq: 0, ops };
    this._undo.applyAndRecord(this._ws, patch);
  }

  getFilter(col: number): ColumnFilter | undefined {
    this._guard('getFilter');
    return this._ws.getColumnFilter(col);
  }

  getVisibleRows(range?: Range): number[] {
    this._guard('getVisibleRows');
    return this._ws.getVisibleRowIndices(range);
  }

  getDistinctValues(col: number, visibleOnly = false): { value: string; count: number }[] {
    this._guard('getDistinctValues');
    return this._ws.getDistinctValues(col, visibleOnly);
  }

  // ── Auto-Filter Range ─────────────────────────────────────────────────────

  setAutoFilterRange(headerRow: number, startCol: number, endCol: number): void {
    this._guard('setAutoFilterRange');
    this._guardSheetOp('allowFilter', 'setAutoFilterRange');
    const before = this._ws.getAutoFilterRange();
    const after: AutoFilterRange = { headerRow, startCol, endCol };
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setAutoFilterRange' as const, before, after }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  clearAutoFilterRange(): void {
    this._guard('clearAutoFilterRange');
    this._guardSheetOp('allowFilter', 'clearAutoFilterRange');
    const before = this._ws.getAutoFilterRange();
    if (before === null) return;
    const patch: WorksheetPatch = {
      seq: 0,
      ops: [{ op: 'setAutoFilterRange' as const, before, after: null }],
    };
    this._undo.applyAndRecord(this._ws, patch);
  }

  getAutoFilterRange(): AutoFilterRange | null {
    this._guard('getAutoFilterRange');
    return this._ws.getAutoFilterRange();
  }

  // ── Range Sort ────────────────────────────────────────────────────────────

  sortRange(range: Range, keys: SortKey[]): void {
    this._guard('sortRange');
    this._guardSheetOp('allowSort', 'sortRange');
    if (keys.length === 0) return;

    this._tracedWrite('sortRange', () => {
    const { start, end } = range;
    const numRows = end.row - start.row + 1;
    const numCols = end.col - start.col + 1;

    // Capture pre-sort snapshot for undo.
    const snapshot: ExtendedCellValue[][] = [];
    for (let ri = 0; ri < numRows; ri++) {
      const row: ExtendedCellValue[] = [];
      for (let ci = 0; ci < numCols; ci++) {
        row.push(this._ws.getCellValue({ row: start.row + ri, col: start.col + ci }));
      }
      snapshot.push(row);
    }

    // Apply sort directly on worksheet.
    this._ws.sortRange(range, keys);

    // Capture post-sort values for redo forward patch.
    const afterSnapshot: ExtendedCellValue[][] = [];
    for (let ri = 0; ri < numRows; ri++) {
      const row: ExtendedCellValue[] = [];
      for (let ci = 0; ci < numCols; ci++) {
        row.push(this._ws.getCellValue({ row: start.row + ri, col: start.col + ci }));
      }
      afterSnapshot.push(row);
    }

    // Build forward + inverse patches from snapshots.
    const makeCellOps = (from: ExtendedCellValue[][], to: ExtendedCellValue[][]): WorksheetPatch['ops'] =>
      from.flatMap((rowVals, ri) =>
        rowVals.map((before, ci) => ({
          op: 'setCellValue' as const,
          row: start.row + ri,
          col: start.col + ci,
          before,
          after: to[ri]![ci] ?? null,
        }))
      );

    const forwardOps = makeCellOps(snapshot, afterSnapshot);
    const inverseOps = makeCellOps(afterSnapshot, snapshot);

    this._undo.recordPreBuilt(
      { seq: 0, ops: forwardOps },
      { seq: 0, ops: inverseOps },
    );
    });
  }

  // ── Pivot ─────────────────────────────────────────────────────────────────

  createPivot(definition: PivotDefinition, target: Address): PivotGrid {
    this._guard('createPivot');

    // ── 1. Validate source bounds ─────────────────────────────────────────
    const { source } = definition;
    const { start: srcStart, end: srcEnd } = source;
    if (
      srcStart.row < 1 || srcStart.col < 1 ||
      srcEnd.row < srcStart.row || srcEnd.col < srcStart.col ||
      srcEnd.row > this._ws.rowCount || srcEnd.col > this._ws.colCount
    ) {
      throw new PivotSourceError(
        `source range (${srcStart.row},${srcStart.col}):(${srcEnd.row},${srcEnd.col}) is out of bounds`,
      );
    }

    // ── 2. Extract source grid (0-indexed) ───────────────────────────────
    const rawGrid: ExtendedCellValue[][] = [];
    for (let r = srcStart.row; r <= srcEnd.row; r++) {
      const row: ExtendedCellValue[] = [];
      for (let c = srcStart.col; c <= srcEnd.col; c++) {
        row.push(this._ws.getCellValue({ row: r, col: c }));
      }
      rawGrid.push(row);
    }

    // ── 3. Compute pivot (pure — may throw PivotSourceError / PivotFieldError / EmptyPivotSourceError) ──
    const pivotGrid = buildPivot(rawGrid, definition);
    const afterValues = pivotGridToValues(pivotGrid);

    // ── 4. (Pivot creation is a structured operation — not blocked by sheet protection.) ──
    const isProtected = this._ws.isSheetProtected();

    // ── 5. Capture before-state ──────────────────────────────────────────
    const beforeValues: ExtendedCellValue[][] = [];
    const beforeStyles: (CellStyle | undefined)[][] = [];
    for (let r = 0; r < afterValues.length; r++) {
      const vRow: ExtendedCellValue[] = [];
      const sRow: (CellStyle | undefined)[] = [];
      for (let c = 0; c < (afterValues[r]?.length ?? 0); c++) {
        const row = target.row + r;
        const col = target.col + c;
        vRow.push(this._ws.getCellValue({ row, col }));
        sRow.push(this._ws.getCell({ row, col })?.style);
      }
      beforeValues.push(vRow);
      beforeStyles.push(sRow);
    }

    return this._tracedWrite('createPivot', () => {
      // ── 6. Write values to worksheet ────────────────────────────────────
      for (let r = 0; r < afterValues.length; r++) {
        for (let c = 0; c < (afterValues[r]?.length ?? 0); c++) {
          const row = target.row + r;
          const col = target.col + c;
          this._ws.setCellValue({ row, col }, afterValues[r]![c] ?? null);
        }
      }

      // ── 7. Auto-lock written cells if sheet is protected ─────────────────
      const afterStyles: (CellStyle | undefined)[][] = [];
      for (let r = 0; r < afterValues.length; r++) {
        const sRow: (CellStyle | undefined)[] = [];
        for (let c = 0; c < (afterValues[r]?.length ?? 0); c++) {
          const row = target.row + r;
          const col = target.col + c;
          if (isProtected) {
            const existing = this._ws.getCell({ row, col })?.style ?? {};
            const locked: CellStyle = { ...existing, locked: true };
            this._ws.setCellStyle({ row, col }, locked);
            sRow.push(locked);
          } else {
            sRow.push(this._ws.getCell({ row, col })?.style);
          }
        }
        afterStyles.push(sRow);
      }

      // ── 8. Build patch entry (single undo entry for the whole pivot) ─────
      const forwardOp = PatchOps.createPivot(
        definition as {
          source: { start: { row: number; col: number }; end: { row: number; col: number } };
          rows: string[];
          values: Array<{ field: string; aggregator: 'sum' | 'count' | 'avg'; label?: string }>;
        },
        target.row, target.col,
        beforeValues, afterValues,
        beforeStyles, afterStyles,
      );

      // Inverse: restore before-state (setCellValue + setCellStyle per cell, reversed)
      const inverseOps: WorksheetPatch['ops'] = [];
      for (let r = afterValues.length - 1; r >= 0; r--) {
        for (let c = (afterValues[r]?.length ?? 0) - 1; c >= 0; c--) {
          const row = target.row + r;
          const col = target.col + c;
          inverseOps.push({
            op:     'setCellStyle',
            row, col,
            before: afterStyles[r]?.[c],
            after:  beforeStyles[r]?.[c],
          });
          inverseOps.push({
            op:     'setCellValue',
            row, col,
            before: afterValues[r]![c] ?? null,
            after:  beforeValues[r]?.[c] ?? null,
          });
        }
      }

      this._undo.recordPreBuilt(
        { seq: 0, ops: [forwardOp] },
        { seq: 0, ops: inverseOps },
      );

      return pivotGrid;
    });
  }

  // ── Events ────────────────────────────────────────────────────────────────

  on(eventType: SdkEventType, listener: SdkEventListener): Disposable {
    this._guard('on');
    if (!this._listeners.has(eventType)) {
      this._listeners.set(eventType, new Set());
    }
    this._listeners.get(eventType)!.add(listener);
    return {
      dispose: () => {
        this._listeners.get(eventType)?.delete(listener);
      },
    };
  }

  // ── Lifecycle ─────────────────────────────────────────────────────────────

  dispose(): void {
    if (this._disposed) return;
    this._disposed = true;
    this._unsubscribe.dispose();
    this._listeners.clear();
    this._undo.clear();
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

const DEFAULT_OPTIONS: Required<SpreadsheetOptions> = {
  maxUndoHistory: 100,
  rows: DEFAULT_WORKSHEET_ROWS,
  cols: DEFAULT_WORKSHEET_COLS,
};

/**
 * Create a new `SpreadsheetSDK` instance backed by a single `Worksheet`.
 *
 * @param name    Worksheet name (defaults to `'Sheet1'`).
 * @param options Configuration overrides.
 */
export function createSpreadsheet(
  name: string = 'Sheet1',
  options: SpreadsheetOptions = {},
): SpreadsheetSDK {
  const resolved: Required<SpreadsheetOptions> = {
    maxUndoHistory: options.maxUndoHistory ?? DEFAULT_OPTIONS.maxUndoHistory,
    rows:           options.rows           ?? DEFAULT_OPTIONS.rows,
    cols:           options.cols           ?? DEFAULT_OPTIONS.cols,
  };
  return new SpreadsheetV1(name, resolved);
}
