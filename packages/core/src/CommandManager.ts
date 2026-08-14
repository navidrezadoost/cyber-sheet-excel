/**
 * CommandManager: Temporal Identity Stability
 * 
 * Undo/Redo with Canonical Pointer Discipline
 * 
 * Rules:
 * 1. Store canonical style pointers only (never clone)
 * 2. Replay via direct pointer reassignment
 * 3. No reconstruction, no spreading, no delta patching
 * 4. Validate: styleA === styleA' after undo cycle
 * 
 * This is not UX plumbing. This is the final identity stress test.
 */

import type { Address, CellValue, ExtendedCellValue, CellStyle } from './types';
import type { Worksheet } from './worksheet';
import type { EventBus } from './EventBus';
import { getCommandEventType } from './eventBridge';
import { GraphInvariantValidator } from './dag/GraphInvariantValidator';
import { GraphTransformationValidator } from './dag/GraphTransformationValidator';
import type { AddressTransform } from './dag/AddressTransform';

/**
 * Command interface: Pure pointer operations
 */
export interface Command {
  /**
   * Execute the operation
   */
  execute(): void;
  
  /**
   * Undo the operation (restore previous pointers)
   */
  undo(): void;
  
  /**
   * Optional description for debugging
   */
  description?: string;
}

/**
 * TransformCommand: Commands that perform coordinate space transformations
 * 
 * Used for Insert/Delete Row/Column operations.
 * 
 * =============================================================================
 * TRANSFORMATION SEMANTICS (CRITICAL CONTRACT)
 * =============================================================================
 * 
 * 1. getTransform() returns f: Address → Address ∪ {null}
 *    - Mapping evaluated in CURRENT coordinate space
 *    - NOT a snapshot of original state
 * 
 * 2. getUndoTransform() returns f⁻¹
 *    - Inverse recomputed from stored metadata
 *    - Also evaluated in CURRENT space at undo time
 * 
 * 3. Validator enforces: graph(f(S)) == f(graph(S))
 *    - Runs after execute(), undo(), redo()
 *    - Catches coordinate space desync immediately
 * 
 * =============================================================================
 * COMPOSITION INVARIANT ENFORCEMENT
 * =============================================================================
 * 
 * CommandManager guarantees:
 *   - Each command evaluated in state produced by previous commands
 *   - Validator runs after EVERY command (no batching)
 *   - Coordinate indices always resolved in CURRENT space
 * 
 * Identity Property (enforced, not just tested):
 *   deleteColumn(k) ∘ insertColumn(k) = identity
 * 
 * This is verified structurally by:
 *   1. Transform evaluated in current state
 *   2. Inverse transform recomputed from metadata
 *   3. Validator checks structural equivalence
 * 
 * =============================================================================
 * EXAMPLE
 * =============================================================================
 * 
 * ```ts
 * class InsertColumnCommand implements TransformCommand {
 *   constructor(private columnIndex: number, private worksheet: Worksheet) {}
 *   
 *   getTransform(): AddressTransform {
 *     // Evaluated at execute() time in CURRENT state
 *     return new InsertColumnTransform(this.columnIndex);
 *   }
 *   
 *   getUndoTransform(): AddressTransform {
 *     // Evaluated at undo() time in CURRENT state (post-insert)
 *     return new DeleteColumnTransform(this.columnIndex);
 *   }
 *   
 *   execute(): void {
 *     const transform = this.getTransform();
 *     // Apply to cells, formulas, DAG, merges, selection
 *     // Validator runs automatically after this
 *   }
 * }
 * ```
 */
export interface TransformCommand extends Command {
  /**
   * Get the forward transformation (for execute)
   * 
   * CRITICAL: Evaluated in CURRENT coordinate space at execute() time,
   * not in original snapshot space.
   */
  getTransform(): AddressTransform;
  
  /**
   * Get the inverse transformation (for undo)
   * 
   * CRITICAL: Recomputed from stored metadata at undo() time,
   * evaluated in CURRENT coordinate space (post-transform state).
   */
  getUndoTransform(): AddressTransform;
}

/**
 * Cell state snapshot: Canonical pointers only
 * 
 * NEVER store:
 * - { ...style } (spreading)
 * - JSON.parse(JSON.stringify(style)) (serialization)
 * - Patch deltas ({ bold: true })
 * 
 * ONLY store:
 * - Direct reference to canonical frozen style
 */
interface CellSnapshot {
  addr: Address;
  value: ExtendedCellValue;
  style: CellStyle | undefined;  // Canonical pointer (frozen, interned)
}

/**
 * SetValueCommand: Change cell value
 */
export class SetValueCommand implements Command {
  private worksheet: Worksheet;
  private addr: Address;
  private previousValue: ExtendedCellValue;
  private newValue: ExtendedCellValue;
  
  constructor(worksheet: Worksheet, addr: Address, newValue: ExtendedCellValue) {
    this.worksheet = worksheet;
    this.addr = addr;
    this.previousValue = worksheet.getCellValue(addr);
    this.newValue = newValue;
  }
  
  execute(): void {
    this.worksheet.setCellValue(this.addr, this.newValue);
  }
  
  undo(): void {
    this.worksheet.setCellValue(this.addr, this.previousValue);
  }
}

/**
 * SetStyleCommand: Change cell style
 * 
 * CRITICAL: Stores canonical pointers, not clones.
 * Undo replays by reassigning the previous canonical reference.
 * 
 * If this passes:
 *   styleA === getCellStyle(addr) after undo
 * 
 * Then temporal stability is proven.
 */
export class SetStyleCommand implements Command {
  private worksheet: Worksheet;
  private addr: Address;
  private previousStyle: CellStyle | undefined;  // Canonical pointer
  private newStyle: CellStyle | undefined;       // Canonical pointer
  
  constructor(worksheet: Worksheet, addr: Address, newStyle: CellStyle | undefined) {
    this.worksheet = worksheet;
    this.addr = addr;
    
    // Capture current canonical style
    this.previousStyle = worksheet.getDirectCellStyle(addr);
    this.newStyle = newStyle;
  }
  
  execute(): void {
    // setCellStyle auto-interns if needed, but prefer passing canonical
    this.worksheet.setCellStyle(this.addr, this.newStyle);
  }
  
  undo(): void {
    // Restore previous canonical pointer
    // This must result in strict pointer equality: styleA === styleA'
    this.worksheet.setCellStyle(this.addr, this.previousStyle);
  }
}

/**
 * BatchCommand: Multiple operations executed atomically
 * 
 * Used for range operations (e.g., format 100 cells)
 */
export class BatchCommand implements Command {
  private commands: Command[];
  public description?: string;
  
  constructor(commands: Command[], description?: string) {
    this.commands = commands;
    this.description = description;
  }
  
  execute(): void {
    for (const cmd of this.commands) {
      cmd.execute();
    }
  }
  
  undo(): void {
    // Undo in reverse order
    for (let i = this.commands.length - 1; i >= 0; i--) {
      this.commands[i].undo();
    }
  }
}

/**
 * SetRangeStyleCommand: Apply style to range
 * 
 * Stores canonical pointers for each cell in range.
 * No reconstruction during replay.
 */
export class SetRangeStyleCommand implements Command {
  private worksheet: Worksheet;
  private snapshots: CellSnapshot[] = [];
  private range: { start: Address; end: Address };
  private newStyle: CellStyle | undefined;
  
  constructor(
    worksheet: Worksheet,
    range: { start: Address; end: Address },
    newStyle: CellStyle | undefined
  ) {
    this.worksheet = worksheet;
    this.range = range;
    this.newStyle = newStyle;
    
    // Capture previous state (canonical pointers)
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        const addr = { row, col };
        this.snapshots.push({
          addr,
          value: worksheet.getCellValue(addr),
          style: worksheet.getDirectCellStyle(addr)  // Canonical pointer
        });
      }
    }
  }
  
  execute(): void {
    for (let row = this.range.start.row; row <= this.range.end.row; row++) {
      for (let col = this.range.start.col; col <= this.range.end.col; col++) {
        this.worksheet.setCellStyle({ row, col }, this.newStyle);
      }
    }
  }
  
  undo(): void {
    // Restore each cell's previous canonical pointer
    for (const snapshot of this.snapshots) {
      this.worksheet.setCellStyle(snapshot.addr, snapshot.style);
    }
  }
}

/**
 * CommandManager: Undo/Redo Stack with Pointer Discipline
 * 
 * Final Identity Stress Test:
 * - Temporal stability (pointer cycling)
 * - No reconstruction overhead
 * - No GC pressure from cloning
 * - Strict === equality after undo
 * 
 * With optional DAG invariant validation (DEV/TEST only):
 * - Guards against DAG corruption
 * - Fails fast on invariant violations
 * - Zero cost in production (process.env.NODE_ENV check)
 */
export class CommandManager {
  private undoStack: Command[] = [];
  private redoStack: Command[] = [];
  private maxHistorySize: number;
  private changeListeners = new Set<() => void>();
  private eventBus?: EventBus;
  
  /** Optional worksheet reference for DAG validation (DEV/TEST only) */
  private worksheet?: Worksheet;
  
  constructor(maxHistorySize: number = 100, worksheet?: Worksheet, eventBus?: EventBus) {
    this.maxHistorySize = maxHistorySize;
    this.worksheet = worksheet;
    this.eventBus = eventBus;
  }

  setEventBus(eventBus: EventBus | undefined): void {
    this.eventBus = eventBus;
  }
  
  /**
   * Execute command and add to history
   * 
   * In DEV/TEST mode: Validates DAG invariants after execution
   * For TransformCommands: Also validates transformation preservation
   */
  execute(command: Command): void {
    command.execute();
    
    // DEV/TEST only: Validate DAG invariants after command execution
    if (process.env.NODE_ENV !== 'production' && this.worksheet) {
      // Access private DAG through worksheet's internal structure
      // This is safe because validators only read state
      const dag = (this.worksheet as any).dag;
      if (dag) {
        // For transformation commands: Use GraphTransformationValidator
        if (this.isTransformCommand(command)) {
          const transform = command.getTransform();
          GraphTransformationValidator.validateAfterTransform(
            transform,
            dag,
            this.worksheet
          );
        } else {
          // For regular commands: Use GraphInvariantValidator
          GraphInvariantValidator.validateAll(dag, this.worksheet);
        }
      }
    }
    
    this.undoStack.push(command);
    if (this.undoStack.length > this.maxHistorySize) {
      this.undoStack.shift();
    }
    
    // Clear redo stack when new command executed
    this.redoStack = [];
    this.notifyChange();
    this.eventBus?.emit('command-execute', {
      commandType: getCommandEventType(command),
      description: command.description,
      command,
    });
  }

  /**
   * Subscribe to command history mutations (execute, undo, redo, clear).
   * Returns an unsubscribe function.
   */
  subscribe(listener: () => void): () => void {
    this.changeListeners.add(listener);
    return () => {
      this.changeListeners.delete(listener);
    };
  }

  private notifyChange(): void {
    for (const listener of this.changeListeners) {
      listener();
    }
  }
  
  /**
   * Undo last command
   * 
   * Returns true if undo was successful
   * In DEV/TEST mode: Validates DAG invariants after undo
   * For TransformCommands: Uses undo transformation for validation
   */
  undo(): boolean {
    const command = this.undoStack.pop();
    if (!command) return false;
    
    command.undo();
    this.redoStack.push(command);
    
    // DEV/TEST only: Validate DAG invariants after undo
    if (process.env.NODE_ENV !== 'production' && this.worksheet) {
      const dag = (this.worksheet as any).dag;
      if (dag) {
        // For transformation commands: Use undo transform for validation
        if (this.isTransformCommand(command)) {
          const undoTransform = command.getUndoTransform();
          GraphTransformationValidator.validateAfterTransform(
            undoTransform,
            dag,
            this.worksheet
          );
        } else {
          GraphInvariantValidator.validateAll(dag, this.worksheet);
        }
      }
    }

    this.notifyChange();
    this.eventBus?.emit('command-undo', {
      commandType: getCommandEventType(command),
      description: command.description,
      command,
    });
    return true;
  }
  
  /**
   * Redo last undone command
   * 
   * Returns true if redo was successful
   * In DEV/TEST mode: Validates DAG invariants after redo
   * For TransformCommands: Uses forward transformation for validation
   */
  redo(): boolean {
    const command = this.redoStack.pop();
    if (!command) return false;
    
    command.execute();
    this.undoStack.push(command);
    
    // DEV/TEST only: Validate DAG invariants after redo
    if (process.env.NODE_ENV !== 'production' && this.worksheet) {
      const dag = (this.worksheet as any).dag;
      if (dag) {
        // For transformation commands: Use forward transform for validation
        if (this.isTransformCommand(command)) {
          const transform = command.getTransform();
          GraphTransformationValidator.validateAfterTransform(
            transform,
            dag,
            this.worksheet
          );
        } else {
          GraphInvariantValidator.validateAll(dag, this.worksheet);
        }
      }
    }

    this.notifyChange();
    this.eventBus?.emit('command-redo', {
      commandType: getCommandEventType(command),
      description: command.description,
      command,
    });
    return true;
  }
  
  /**
   * Check if undo is available
   */
  canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  
  /**
   * Check if redo is available
   */
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }
  
  /**
   * Clear all history
   */
  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.notifyChange();
  }
  
  /**
   * Get history size
   */
  getHistorySize(): { undoCount: number; redoCount: number } {
    return {
      undoCount: this.undoStack.length,
      redoCount: this.redoStack.length
    };
  }
  
  /**
   * Get undo stack description (for debugging)
   */
  getUndoHistory(): string[] {
    return this.undoStack
      .map(cmd => cmd.description || 'Unnamed command')
      .filter(Boolean);
  }
  
  /**
   * Type guard: Check if command is a TransformCommand
   */
  private isTransformCommand(command: Command): command is TransformCommand {
    return 'getTransform' in command && 'getUndoTransform' in command;
  }
}

/**
 * Factory functions for common commands
 */
export class Commands {
  /**
   * Create SetValue command
   */
  static setValue(worksheet: Worksheet, addr: Address, value: CellValue): Command {
    return new SetValueCommand(worksheet, addr, value);
  }
  
  /**
   * Create SetStyle command
   */
  static setStyle(worksheet: Worksheet, addr: Address, style: CellStyle | undefined): Command {
    return new SetStyleCommand(worksheet, addr, style);
  }
  
  /**
   * Create SetRangeStyle command
   */
  static setRangeStyle(
    worksheet: Worksheet,
    range: { start: Address; end: Address },
    style: CellStyle | undefined
  ): Command {
    return new SetRangeStyleCommand(worksheet, range, style);
  }
  
  /**
   * Create batch command
   */
  static batch(commands: Command[], description?: string): Command {
    return new BatchCommand(commands, description);
  }
}
