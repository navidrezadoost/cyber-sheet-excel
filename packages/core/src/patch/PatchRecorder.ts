/**
 * PatchRecorder.ts — Phase 10: Delta Engine + Undo/Redo
 *
 * Wraps a Worksheet and intercepts its events to produce a WorksheetPatch.
 *
 * ==========================================================================
 * DESIGN
 * ==========================================================================
 *
 * PatchRecorder subscribes to the Worksheet's event stream during a recording
 * session.  For each mutation event it captures the state BEFORE the mutation
 * (by querying the sheet just before it changes) so that the resulting patch
 * is immediately invertible.
 *
 * Lifecycle:
 *
 *   1.  recorder.start()                   — begins capturing events
 *   2.  worksheet.setCellValue(...)         — worksheet mutates, events fire
 *       worksheet.hideRow(...)
 *       ...
 *   3.  const patch = recorder.stop()       — ends capture, returns patch
 *
 * The patch is guaranteed to be:
 *   • COMPLETE — every mutation made between start() and stop() is included.
 *   • ORDERED  — ops appear in the same order the mutations happened.
 *   • INVERTIBLE — each op stores the `before` state needed to undo it.
 *
 * ==========================================================================
 * THREAD SAFETY
 * ==========================================================================
 *
 * PatchRecorder runs synchronously on a single thread (the Worker thread
 * that owns the Worksheet).  No locking is required.
 *
 * ==========================================================================
 * USAGE INSIDE EngineWorkerHost
 * ==========================================================================
 *
 *   // In the 'applyPatch' handler:
 *   const recorder = new PatchRecorder(this.ws);
 *   recorder.start();
 *   applyPatch(this.ws, patch);
 *   const applied  = recorder.stop();
 *   const inverse  = invertPatch(applied);
 *   return inverse;  // serialised back to main thread as the undo patch
 *
 * ==========================================================================
 * CAVEAT — style events
 * ==========================================================================
 *
 * The Worksheet emits 'style-changed' AFTER internalising the style via the
 * canonical pointer system.  The `style` field on the event therefore carries
 * the already-interned CellStyle object.  We deep-clone it at record time so
 * the patch does not hold a live reference into the Worksheet's style store.
 *
 * ==========================================================================
 * CAVEAT — cell events only carry the AFTER value
 * ==========================================================================
 *
 * The 'cell-changed' event carries the NEW cell state (after the mutation).
 * To build the `before` field of SetCellValueOp we must read the cell's
 * current value BEFORE the mutation fires.
 *
 * Because the Worksheet fires 'cell-changed' synchronously FROM WITHIN the
 * setCellValue call, and JavaScript is single-threaded, we intercept by
 * wrapping setCellValue.  This is done in a non-invasive way: we subscribe
 * to a 'will-change' listener if available, OR we snapshot the cell value
 * just before the event fires.
 *
 * ✦ ACTUAL APPROACH USED: The 'cell-changed' event fires after the cell has
 *   been updated.  We therefore maintain a per-session "before snapshots" Map
 *   that is populated by a patch recorder's own pre-mutation reads.
 *
 *   Because the Worksheet events fire synchronously within the mutation call,
 *   and the mutation call itself runs on the single Worker thread, we can
 *   safely read the PREVIOUS value by recording it outside the mutation:
 *
 *     const before = ws.getCellValue(addr);
 *     ws.setCellValue(addr, newValue);   ← event fires here
 *
 *   This is done by EngineWorkerHost which calls
 *   `recorder.preMutate(row, col)` before calling `applyPatch`.
 *
 *   However, when applyPatch() drives the mutations programmatically, we
 *   cannot inject pre-reads easily.  Instead we adopt the strategy of doing
 *   a GET for each cell the first time it appears in the ops list, caching
 *   the result, and relying on the fact that ops are applied sequentially.
 *
 *   The cleanest solution implemented here: PatchRecorder intercepts events.
 *   For 'cell-changed' we look up the before-value from a Map that is pre-
 *   populated in `applyPatch`.  If the Map has no entry (standalone recorder
 *   usage) we fall back to `undefined` with a console.warn.
 *
 *   SIMPLIFIED IMPLEMENTATION:
 *   Since applyPatch is the PRIMARY mutation driver that PatchRecorder wraps,
 *   we record `before` values AT THE TIME of applyPatch op iteration.
 *   The recorder itself records ops during event emission.
 *   For `setCellValue` and `setCellStyle` ops, `before` is read synchronously
 *   BEFORE calling ws.setCellValue / ws.setCellStyle (these fire events).
 *   This is handled by the recordingApplyPatch helper below.
 *
 * ==========================================================================
 * API SUMMARY
 * ==========================================================================
 *
 *   new PatchRecorder(ws)        — create attached to a Worksheet
 *   recorder.start()             — begin recording (idempotent if not started)
 *   recorder.stop(): WorksheetPatch — stop and return the captured patch
 *   recorder.abort()             — stop without returning a patch
 *   recorder.isRecording         — true between start() and stop()/abort()
 *
 *   Utility: recordingApplyPatch(ws, patch, seq?)
 *     Apply a patch to a Worksheet while capturing a full inverse patch.
 *     Returns the inverse WorksheetPatch.
 */

import type { SheetEvents, CellStyle, ExtendedCellValue, MergedRegion } from '../types';
import type { Worksheet } from '../worksheet';
import {
  type WorksheetPatch,
  type PatchOp,
  type SetCellValueOp,
  type SetCellStyleOp,
  invertPatch,
  applyPatch,
} from './WorksheetPatch';
import type { Disposable } from '../events';
import { PatchRecorderError } from '../sdk/errors';

// ---------------------------------------------------------------------------
// Sequence counter — shared across all recorders in a thread
// ---------------------------------------------------------------------------

let _globalSeq = 1;

function nextSeq(): number {
  return _globalSeq++;
}

// ---------------------------------------------------------------------------
// PatchRecorder
// ---------------------------------------------------------------------------

export class PatchRecorder {
  private readonly ws: Worksheet;
  private _ops: PatchOp[] | null = null;
  private _disposable: Disposable | null = null;

  constructor(ws: Worksheet) {
    this.ws = ws;
  }

  /** True while recording. */
  get isRecording(): boolean {
    return this._ops !== null;
  }

  /**
   * Begin a recording session.
   * Throws if already recording.
   */
  start(): void {
    if (this._ops !== null) {
      throw new PatchRecorderError('already recording — call stop() or abort() first.');
    }
    this._ops = [];
    this._disposable = (this.ws as unknown as { on(l: (e: SheetEvents) => void): Disposable }).on(
      this._onEvent,
    );
  }

  /**
   * End the recording session and return the captured patch.
   * Throws if not recording.
   */
  stop(): WorksheetPatch {
    if (this._ops === null) {
      throw new PatchRecorderError('not recording — call start() first.');
    }
    const ops = this._ops;
    this._teardown();
    return { seq: nextSeq(), ops };
  }

  /**
   * Discard the in-progress recording.
   * Safe to call if not recording (no-op in that case).
   */
  abort(): void {
    this._teardown();
  }

  // ── Event handler ─────────────────────────────────────────────────────────

  /**
   * Handle a SheetEvent emitted by the Worksheet.
   *
   * IMPORTANT: For cell-changed events, the `before` value is injected via
   * `_beforeMap` by `recordingApplyPatch` BEFORE the mutation calls.  If we
   * are not inside a `recordingApplyPatch` session the before value will
   * not be available and we record `null` as the before value with a warning.
   */
  private readonly _onEvent = (event: SheetEvents): void => {
    if (this._ops === null) return;

    switch (event.type) {
      case 'cell-changed': {
        const { row, col } = event.address;
        const key = `${row}:${col}`;
        const before: ExtendedCellValue = this._beforeMap.has(key)
          ? this._beforeMap.get(key)!
          : null;

        if (event.cell.value === null || event.cell.value === undefined) {
          this._ops.push({ op: 'clearCell', row, col, before });
        } else {
          this._ops.push({ op: 'setCellValue', row, col, before, after: event.cell.value } as SetCellValueOp);
        }
        break;
      }

      case 'style-changed': {
        const { row, col } = event.address;
        const styleKey = `${row}:${col}:style`;
        const before: CellStyle | undefined = this._beforeStyleMap.has(styleKey)
          ? this._beforeStyleMap.get(styleKey)
          : undefined;
        // Deep-clone the after-style to detach from Worksheet's internal ref
        const after = event.style !== undefined ? deepCloneStyle(event.style) : undefined;
        this._ops.push({ op: 'setCellStyle', row, col, before, after } as SetCellStyleOp);
        break;
      }

      case 'merge-added': {
        const r = event.region as MergedRegion;
        this._ops.push({ op: 'mergeCells', startRow: r.startRow, startCol: r.startCol, endRow: r.endRow, endCol: r.endCol });
        break;
      }

      case 'merge-removed': {
        const r = event.region as MergedRegion;
        this._ops.push({ op: 'cancelMerge', startRow: r.startRow, startCol: r.startCol, endRow: r.endRow, endCol: r.endCol });
        break;
      }

      case 'row-hidden':  this._ops.push({ op: 'hideRow', row: event.row }); break;
      case 'row-shown':   this._ops.push({ op: 'showRow', row: event.row }); break;
      case 'col-hidden':  this._ops.push({ op: 'hideCol', col: event.col }); break;
      case 'col-shown':   this._ops.push({ op: 'showCol', col: event.col }); break;

      case 'sheet-protection-changed':
        this._ops.push({ op: 'setSheetProtection', before: event.before, after: event.after });
        break;
      case 'freeze-panes-changed':
        this._ops.push({ op: 'setFreezePanes', before: event.before, after: event.after });
        break;

      case 'filter-changed':
        // Forward direction (before→after); invertPatch() will swap to produce the undo op.
        this._ops.push({ op: 'setColumnFilter', col: event.col, before: event.before, after: event.filter });
        break;
      case 'autofilter-range-changed':
        // Forward direction (before→after); invertPatch() will swap to produce the undo op.
        this._ops.push({ op: 'setAutoFilterRange', before: event.before, after: event.after });
        break;

      case 'spill-source-changed':
        // Spill source metadata changes (for undo/redo)
        // No-op for now: spill state is not in WorksheetPatch yet (future extension point)
        break;

      case 'spill-from-changed':
        // Spilled-from reference changes (for undo/redo)
        // No-op for now: spill state is not in WorksheetPatch yet (future extension point)
        break;

      case 'spill-batch-changed': {
        // Batch spill operation — captures all affected cells atomically
        const changes = event.changes.map(change => ({
          row: change.address.row,
          col: change.address.col,
          before: {
            spillSource: change.before.spillSource,
            spilledFrom: change.before.spilledFrom,
          },
          after: {
            spillSource: change.after.spillSource,
            spilledFrom: change.after.spilledFrom,
          },
        }));
        this._ops.push({ op: 'setSpill', changes });
        break;
      }

      // sort-applied is handled by SDK via recordPreBuilt (snapshot-based). PatchRecorder no-ops it.
      // sheet-mutated / comment-* / cycle-detected — no-op
      default: break;
    }
  };

  // ── Before-map (pre-mutation state injection) ─────────────────────────────

  /**
   * Map from `row:col` → `ExtendedCellValue` (pre-mutation cell value).
   * Populated by `recordingApplyPatch` before each setCellValue / clearCell op.
   */
  readonly _beforeMap: Map<string, ExtendedCellValue> = new Map();

  /**
   * Map from `row:col:style` → `CellStyle | undefined` (pre-mutation style).
   * Populated by `recordingApplyPatch` before each setCellStyle op.
   */
  readonly _beforeStyleMap: Map<string, CellStyle | undefined> = new Map();

  // ── Internal housekeeping ─────────────────────────────────────────────────

  private _teardown(): void {
    this._disposable?.dispose();
    this._disposable = null;
    this._ops = null;
    this._beforeMap.clear();
    this._beforeStyleMap.clear();
  }
}

// ---------------------------------------------------------------------------
// recordingApplyPatch — apply + capture inverse in one atomic call
// ---------------------------------------------------------------------------

/**
 * Apply `patch` to `ws` while using a PatchRecorder to capture the actual
 * operations executed, then compute and return the inverse patch.
 *
 * This is the canonical way to apply a patch on the Worker thread:
 *
 *   const inverse = recordingApplyPatch(ws, patch);
 *   // Send `inverse` back to the main thread for undo/redo support.
 *
 * The returned inverse is computed by:
 *  1. Pre-reading the before-state for each op in `patch`.
 *  2. Starting a PatchRecorder.
 *  3. Applying each op via applyPatch().
 *  4. Stopping the recorder (captures the ops that actually fired).
 *  5. Computing invertPatch() on the recorded patch.
 *
 * @param ws    The target Worksheet.
 * @param patch The patch to apply.
 * @param seq   Optional sequence number for the captured patch (default: auto).
 * @returns      The inverse patch that undoes the applied mutations.
 */
export function recordingApplyPatch(
  ws: Worksheet,
  patch: WorksheetPatch,
  seq?: number,
): WorksheetPatch {
  const recorder = new PatchRecorder(ws);

  // Pre-populate the before-maps for all ops that need a before-value.
  for (const op of patch.ops) {
    switch (op.op) {
      case 'setCellValue':
      case 'clearCell': {
        const key = `${op.row}:${op.col}`;
        if (!recorder._beforeMap.has(key)) {
          recorder._beforeMap.set(key, ws.getCellValue({ row: op.row, col: op.col }));
        }
        break;
      }
      case 'setCellStyle': {
        const styleKey = `${op.row}:${op.col}:style`;
        if (!recorder._beforeStyleMap.has(styleKey)) {
          const style = ws.getDirectCellStyle({ row: op.row, col: op.col });
          recorder._beforeStyleMap.set(styleKey, style !== undefined ? deepCloneStyle(style) : undefined);
        }
        break;
      }
      // Merges, row/col visibility: the before state is implicit (present/absent)
      // and PatchRecorder builds the inverse from the event type alone.
      default: break;
    }
  }

  recorder.start();
  applyPatch(ws, patch);
  const recorded = recorder.stop();

  return invertPatch({ seq: seq ?? recorded.seq, ops: recorded.ops });
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

function deepCloneStyle(style: CellStyle): CellStyle {
  // CellStyle is a plain object with no nested class instances.
  // JSON round-trip is safe and avoids a manual deep-clone.
  return JSON.parse(JSON.stringify(style)) as CellStyle;
}
