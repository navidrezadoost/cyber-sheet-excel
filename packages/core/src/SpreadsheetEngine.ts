/**
 * SpreadsheetEngine.ts — Execution Kernel (Orchestrator Layer)
 *
 * ==========================================================================
 * WHY THIS EXISTS
 * ==========================================================================
 *
 * We have 3 independently-correct subsystems:
 *   1. TransactionSystem — guarantees state(t+1) = apply(state(t), ops)
 *   2. Patch/Spill system  — guarantees invertibility (group algebra)
 *   3. RecomputeScheduler  — guarantees deterministic evaluation order
 *
 * Composition tests proved they integrate correctly.
 *
 * BUT: it's still possible to misuse them:
 *   ❌ ws.setCellValue(A1, '10'); scheduler.flush(); // manual scheduling
 *   ❌ scheduler.flush(); ws.setCellValue(A1, '20'); // mid-flight mutation
 *   ❌ ws.setCellValue(...); eventBus.emit(...);    // events before recompute
 *
 * This layer makes incorrect usage **structurally impossible**.
 *
 * ==========================================================================
 * DESIGN PRINCIPLE
 * ==========================================================================
 *
 * The SpreadsheetEngine is the **only legal way to mutate state**.
 *
 * It enforces the atomic execution pipeline:
 *
 *   run(callback)
 *   ↓
 *   1. beginTransaction()
 *   2. callback() executes (mutations allowed ONLY here)
 *   3. commitTransaction() → returns dirty set
 *   4. scheduler.schedule(dirty)
 *   5. scheduler.flush() → evaluates formulas
 *   6. eventBus.emit() → observers notified
 *
 * External callers CANNOT:
 *   • Call scheduler.flush() directly (private)
 *   • Mutate during flush (execution state enforced)
 *   • Emit events manually (engine controls timing)
 *   • Begin/commit transactions directly (encapsulated)
 *
 * ==========================================================================
 * EXECUTION STATES
 * ==========================================================================
 *
 * The engine is a state machine with 3 states:
 *
 *   IDLE       → ready for new run()
 *   MUTATING   → inside callback, mutations allowed
 *   COMPUTING  → scheduler running, mutations forbidden
 *
 * State transitions:
 *
 *   IDLE → [run() called] → MUTATING
 *   MUTATING → [callback returns] → COMPUTING
 *   COMPUTING → [scheduler drains] → IDLE
 *
 * Illegal transitions throw ExecutionError.
 *
 * ==========================================================================
 * CORE INVARIANTS
 * ==========================================================================
 *
 * E1. Single execution thread
 *     Only one run() can be active at a time. Concurrent calls throw.
 *
 * E2. Mutation isolation
 *     Mutations are only legal inside the run() callback. Direct
 *     Worksheet calls outside run() throw.
 *
 * E3. Scheduler encapsulation
 *     scheduler.flush() is private. Only the engine can trigger evaluation.
 *
 * E4. Event ordering
 *     Events are emitted AFTER scheduler completes, never during mutations.
 *
 * E5. Deterministic execution
 *     Same mutations + same topological order → same result. Always.
 *
 * E6. Re-entrancy safety
 *     Event handlers cannot call run() (state is COMPUTING during emission).
 *     Attempts throw ExecutionError. Use setTimeout for sequential runs.
 *
 * ==========================================================================
 * API SURFACE
 * ==========================================================================
 *
 * ```ts
 * const engine = new SpreadsheetEngine();
 *
 * // Mutation API (the ONLY legal way to change state)
 * await engine.run(async (ws) => {
 *   ws.setCellValue({ row: 1, col: 1 }, 10);
 *   ws.setCellValue({ row: 2, col: 1 }, '=A1+1');
 * });
 *
 * // Read API (safe to call anytime)
 * const value = engine.getCellValue({ row: 2, col: 1 }); // 11
 *
 * // Event subscription
 * engine.on('cellsChanged', (event) => {
 *   console.log('Changed:', event.addresses);
 * });
 * ```
 *
 * ==========================================================================
 * IMPLEMENTATION NOTES
 * ==========================================================================
 *
 * The engine holds:
 *   • Worksheet instance (private)
 *   • TransactionContext (created per run())
 *   • RecomputeScheduler (private, never exposed)
 *   • EventEmitter (private, controlled emission)
 *   • ExecutionState (IDLE | MUTATING | COMPUTING)
 *
 * During run():
 *   1. Assert state == IDLE (else throw ExecutionError)
 *   2. Set state = MUTATING
 *   3. Begin transaction
 *   4. Execute callback(worksheet)
 *   5. Commit → get dirty set
 *   6. Set state = COMPUTING
 *   7. Schedule dirty cells
 *   8. Flush scheduler (may be async)
 *   9. Emit events (cellsChanged, formulasEvaluated, etc.)
 *  10. Set state = IDLE
 *
 * Error handling:
 *   • If callback throws → rollback + set state = IDLE + rethrow
 *   • If scheduler throws → state = IDLE + rethrow (state is corrupt)
 */

import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Worksheet } from './worksheet';
import { FormulaEngine } from './FormulaEngine';
import {
  RecomputeScheduler,
  TaskPriority,
  type ScheduledTask,
  type ViewportRect,
} from './RecomputeScheduler';
import { TransactionContext, TransactionError, type CommitResult } from './transaction';
import { PatchRecorder } from './patch/PatchRecorder';
import { type WorksheetPatch } from './patch/WorksheetPatch';
import type { Address, CellValue, ExtendedCellValue } from './types';

// ---------------------------------------------------------------------------
// Execution state machine
// ---------------------------------------------------------------------------

const enum ExecutionState {
  /** Ready for new run(). No active transaction or scheduler work. */
  IDLE      = 'IDLE',
  /** Inside run() callback. Mutations allowed. */
  MUTATING  = 'MUTATING',
  /** Scheduler running. Mutations forbidden. */
  COMPUTING = 'COMPUTING',
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

export class ExecutionError extends Error {
  constructor(
    public readonly code: 'CONCURRENT_RUN' | 'MUTATION_OUTSIDE_RUN' | 'ILLEGAL_STATE',
    message: string
  ) {
    super(message);
    this.name = 'ExecutionError';
  }
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

export interface CellsChangedEvent {
  type: 'cellsChanged';
  addresses: Address[];
  patch: WorksheetPatch;
}

export interface FormulasEvaluatedEvent {
  type: 'formulasEvaluated';
  count: number;
  hasCycles: boolean;
}

export type EngineEvent = CellsChangedEvent | FormulasEvaluatedEvent;

type EventHandler = (event: EngineEvent) => void;

// ---------------------------------------------------------------------------
// SpreadsheetEngine (Orchestrator)
// ---------------------------------------------------------------------------

/**
 * Execution kernel for the spreadsheet runtime.
 *
 * Enforces the atomic execution pipeline:
 *   mutation → commit → schedule → evaluate → notify
 *
 * This is the ONLY legal way to mutate spreadsheet state.
 */
export class SpreadsheetEngine {
  private readonly _ws: Worksheet;
  private readonly _formulaEngine: FormulaEngine;
  private readonly _scheduler: RecomputeScheduler;
  private readonly _eventHandlers: Map<string, Set<EventHandler>> = new Map();

  private _state: ExecutionState = ExecutionState.IDLE;
  private _txn: TransactionContext | null = null;

  constructor(name = 'Sheet1') {
    this._ws = new Worksheet(name, DEFAULT_WORKSHEET_ROWS, DEFAULT_WORKSHEET_COLS, undefined, undefined, this);
    this._formulaEngine = new FormulaEngine();

    // Scheduler is PRIVATE — external callers cannot flush() directly
    this._scheduler = new RecomputeScheduler({
      timeSliceMs: 8,
      evaluator: (task: ScheduledTask) => {
        const cell = this._ws.getCell(task);
        if (!cell || !cell.formula) return null;

        return this._formulaEngine.evaluate(cell.formula, {
          worksheet: this._ws,
          currentCell: task,
        });
      },
      onComplete: (task: ScheduledTask, value: unknown) => {
        this._ws.setCellValue(task, value as CellValue);
      },
    });
  }

  // ── Public API ───────────────────────────────────────────────────────────

  /**
   * Execute a mutation callback inside a transaction.
   *
   * The execution pipeline:
   *   1. Begin transaction
   *   2. Execute callback (mutations allowed ONLY here)
   *   3. Commit transaction → dirty set
   *   4. Schedule dirty cells
   *   5. Evaluate formulas (async or sync based on mode)
   *   6. Emit events
   *
   * Throws:
   *   • ExecutionError(CONCURRENT_RUN) if already running
   *   • ExecutionError(ILLEGAL_STATE) if callback is async
   *   • TransactionError if transaction fails
   *   • Any error thrown by callback (after rollback)
   */
  async run(callback: (ws: Worksheet) => void): Promise<void> {
    // E1: Single execution thread
    if (this._state !== ExecutionState.IDLE) {
      throw new ExecutionError(
        'CONCURRENT_RUN',
        `Cannot run() while state=${this._state}. Wait for previous run() to complete.`
      );
    }

    this._state = ExecutionState.MUTATING;

    try {
      // Begin transaction
      this._txn = new TransactionContext(this._ws);

      // Execute callback (mutations allowed)
      // CRITICAL: Reject async callbacks to prevent mutations after state transition
      const result = callback(this._ws);
      if (result && typeof result === 'object' && 'then' in result) {
        throw new ExecutionError(
          'ILLEGAL_STATE',
          'engine.run() callback must be synchronous. Async callbacks create '
          + 'ambiguous execution windows where mutations could occur after MUTATING '
          + 'state ends. Use synchronous mutations only.'
        );
      }

      // Commit → get dirty set
      const { patch, inverse } = this._txn.commit();
      this._txn = null;

      // Transition to COMPUTING (mutations now forbidden)
      this._state = ExecutionState.COMPUTING;

      // Extract dirty addresses from patch
      const dirtyAddresses = this._extractDirtyAddresses(patch);

      // Schedule dirty cells (topological order from DependencyGraph would go here)
      for (const addr of dirtyAddresses) {
        this._scheduler.schedule(addr, TaskPriority.High);
      }

      // Flush scheduler (evaluate all formulas)
      await this._scheduler.flush();

      // Emit events (AFTER recompute completes — E4)
      this._emit({
        type: 'cellsChanged',
        addresses: dirtyAddresses,
        patch,
      });

      this._emit({
        type: 'formulasEvaluated',
        count: this._scheduler.metrics.completed,
        hasCycles: false, // TODO: wire from DependencyGraph
      });

      // Return to IDLE
      this._state = ExecutionState.IDLE;

    } catch (err) {
      // Rollback on error
      if (this._txn) {
        this._txn.rollback();
        this._txn = null;
      }

      // Reset scheduler (cancel stale tasks)
      this._scheduler.invalidate();

      // Return to IDLE
      this._state = ExecutionState.IDLE;

      throw err;
    }
  }

  /**
   * Check if the engine is currently in MUTATING state.
   * Used by Worksheet to enforce E2 invariant (mutations only inside run()).
   */
  isMutating(): boolean {
    return this._state === ExecutionState.MUTATING;
  }

  /**
   * Read a cell value (safe to call anytime).
   */
  getCellValue(addr: Address): ExtendedCellValue {
    return this._ws.getCellValue(addr);
  }

  /**
   * Set viewport (influences scheduler priority).
   */
  setViewport(viewport: ViewportRect): void {
    this._scheduler.setViewport(viewport);
  }

  /**
   * Subscribe to engine events.
   */
  on(eventType: 'cellsChanged' | 'formulasEvaluated', handler: EventHandler): () => void {
    let handlers = this._eventHandlers.get(eventType);
    if (!handlers) {
      handlers = new Set();
      this._eventHandlers.set(eventType, handlers);
    }
    handlers.add(handler);

    // Return unsubscribe function
    return () => {
      handlers?.delete(handler);
    };
  }

  /**
   * Get current execution state (for debugging/testing).
   */
  get executionState(): string {
    return this._state;
  }

  /**
   * Get scheduler metrics (for debugging/testing).
   */
  get metrics() {
    return this._scheduler.metrics;
  }

  /**
   * Reset scheduler metrics.
   */
  resetMetrics(): void {
    this._scheduler.resetMetrics();
  }

  // ── Private helpers ──────────────────────────────────────────────────────

  private _extractDirtyAddresses(patch: WorksheetPatch): Address[] {
    const addresses: Address[] = [];
    const seen = new Set<string>();

    for (const op of patch.ops) {
      if (op.op === 'setCellValue' || op.op === 'clearCell') {
        const key = `${op.row},${op.col}`;
        if (!seen.has(key)) {
          seen.add(key);
          addresses.push({ row: op.row, col: op.col });
        }
      }
      // TODO: handle setSpill ops (extract all affected cells)
    }

    return addresses;
  }

  private _emit(event: EngineEvent): void {
    const handlers = this._eventHandlers.get(event.type);
    if (!handlers) return;

    for (const handler of handlers) {
      try {
        handler(event);
      } catch (err) {
        // Swallow handler errors to prevent cascading failures
        console.error(`Event handler error for ${event.type}:`, err);
      }
    }
  }
}
