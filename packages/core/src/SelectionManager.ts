/**
 * SelectionManager.ts
 * 
 * Manages cell/range selection state and provides selection-based operations.
 * Central authority for tracking active cell, selection ranges, and selection metadata.
 * 
 * Architecture: Single Source of Truth for Selection
 * - Emits events on selection changes
 * - Provides selection state queries (mixed styles, formats, etc.)
 * - Handles selection transformations (expand, shrink, navigate)
 * - Coordinates with renderer for visual feedback
 */

import type { Address, Range, CellStyle } from './types';
import type { Worksheet } from './worksheet';
import { Emitter } from './events';

export interface SelectionState {
  /** Primary active cell (cursor position, where typing goes) */
  activeCell: Address;
  
  /** Selected ranges (can be multiple for Ctrl+click selections) */
  ranges: Range[];
  
  /** Primary range (the last selected range, typically ranges[ranges.length - 1]) */
  primaryRange: Range;
  
  /** Whether the selection is in edit mode (typing replaces content) */
  isEditing: boolean;
}

export interface SelectionEvent {
  type: 'selection-changed';
  selection: SelectionState;
  previousSelection: SelectionState | null;
}

/**
 * Style state for multi-cell selections
 * Indicates whether a style property is consistent across selection
 */
export interface StyleState<T> {
  /** The value (if consistent) or undefined (if mixed) */
  value: T | undefined;
  
  /** Whether the value is mixed across selection */
  isMixed: boolean;
}

export interface SelectionStyleSummary {
  // Font properties
  fontFamily?: StyleState<string | undefined>;
  fontSize?: StyleState<number | undefined>;
  bold?: StyleState<boolean | undefined>;
  italic?: StyleState<boolean | undefined>;
  underline?: StyleState<boolean | string | undefined>;
  strikethrough?: StyleState<boolean | undefined>;
  color?: StyleState<any>; // Can be string or ExcelColorSpec
  
  // Alignment properties
  align?: StyleState<string | undefined>;
  valign?: StyleState<string | undefined>;
  wrap?: StyleState<boolean | undefined>;
  rotation?: StyleState<number | undefined>;
  
  // Fill and border
  fill?: StyleState<any>;
  border?: StyleState<any>;
  
  // Number format
  numberFormat?: StyleState<string | undefined>;
}

/**
 * SelectionManager - Central selection state management
 * 
 * Responsibilities:
 * - Track active cell and selected ranges
 * - Emit selection change events
 * - Provide selection-based style queries (mixed state detection)
 * - Handle keyboard navigation (arrow keys, Home, End, Ctrl+arrow)
 * - Manage edit mode state
 * - Coordinate with undo/redo for selection restoration
 * 
 * @architecture State management layer
 * @invariant activeCell is always within worksheet bounds
 * @invariant ranges array is never empty
 */
export class SelectionManager extends Emitter<SelectionEvent> {
  private worksheet: Worksheet;
  private state: SelectionState;
  
  constructor(worksheet: Worksheet, initialCell: Address = { row: 1, col: 1 }) {
    super();
    this.worksheet = worksheet;
    
    // Initialize with single cell selection
    this.state = {
      activeCell: initialCell,
      ranges: [{ start: initialCell, end: initialCell }],
      primaryRange: { start: initialCell, end: initialCell },
      isEditing: false,
    };
  }
  
  /**
   * Get current selection state (readonly)
   */
  getState(): Readonly<SelectionState> {
    return { ...this.state };
  }
  
  /**
   * Get active cell address
   */
  getActiveCell(): Address {
    return { ...this.state.activeCell };
  }
  
  /**
   * Select a single cell
   */
  selectCell(address: Address): void {
    const previousState = { ...this.state };
    
    this.state = {
      activeCell: address,
      ranges: [{ start: address, end: address }],
      primaryRange: { start: address, end: address },
      isEditing: false,
    };
    
    this.emit({
      type: 'selection-changed',
      selection: this.state,
      previousSelection: previousState,
    });
  }
  
  /**
   * Select a range of cells
   */
  selectRange(range: Range, append: boolean = false): void {
    const previousState = { ...this.state };
    
    if (append) {
      // Add to existing selection (Ctrl+click behavior)
      this.state = {
        ...this.state,
        ranges: [...this.state.ranges, range],
        primaryRange: range,
        activeCell: range.start,
        isEditing: false,
      };
    } else {
      // Replace selection
      this.state = {
        activeCell: range.start,
        ranges: [range],
        primaryRange: range,
        isEditing: false,
      };
    }
    
    this.emit({
      type: 'selection-changed',
      selection: this.state,
      previousSelection: previousState,
    });
  }
  
  /**
   * Extend selection from active cell to target address (Shift+click)
   */
  extendSelection(targetAddress: Address): void {
    const previousState = { ...this.state };
    
    const newRange: Range = {
      start: this.state.activeCell,
      end: targetAddress,
    };
    
    this.state = {
      ...this.state,
      ranges: [newRange],
      primaryRange: newRange,
      isEditing: false,
    };
    
    this.emit({
      type: 'selection-changed',
      selection: this.state,
      previousSelection: previousState,
    });
  }
  
  /**
   * Move active cell by offset (arrow key navigation)
   */
  moveActiveCell(rowOffset: number, colOffset: number, extend: boolean = false): void {
    const currentRow = this.state.activeCell.row;
    const currentCol = this.state.activeCell.col;
    
    const newRow = Math.max(1, currentRow + rowOffset);
    const newCol = Math.max(1, currentCol + colOffset);
    
    const newAddress: Address = { row: newRow, col: newCol };
    
    if (extend) {
      this.extendSelection(newAddress);
    } else {
      this.selectCell(newAddress);
    }
  }
  
  /**
   * Select the cells that contain data in the worksheet.
   */
  selectAll(): void {
    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxRow = 0;
    let maxCol = 0;

    this.worksheet.forEachNonEmptyCell((row, col, cell) => {
      const hasValue = cell?.value !== undefined && cell.value !== null && cell.value !== '';
      if (!hasValue && !cell?.formula) return;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    });

    const allRange: Range = {
      start: maxRow === 0 ? this.state.activeCell : { row: minRow, col: minCol },
      end: maxRow === 0 ? this.state.activeCell : { row: maxRow, col: maxCol },
    };
    
    this.selectRange(allRange);
  }
  
  /**
   * Enter edit mode (F2 or double-click)
   */
  enterEditMode(): void {
    if (!this.state.isEditing) {
      this.state = { ...this.state, isEditing: true };
      this.emit({
        type: 'selection-changed',
        selection: this.state,
        previousSelection: null,
      });
    }
  }
  
  /**
   * Exit edit mode (Enter, Tab, or Esc)
   */
  exitEditMode(): void {
    if (this.state.isEditing) {
      this.state = { ...this.state, isEditing: false };
      this.emit({
        type: 'selection-changed',
        selection: this.state,
        previousSelection: null,
      });
    }
  }
  
  /**
   * Get all cells in current selection
   * Returns array of addresses covering all selected ranges
   */
  getSelectedCells(): Address[] {
    const cells: Address[] = [];
    
    for (const range of this.state.ranges) {
      const startRow = Math.min(range.start.row, range.end.row);
      const endRow = Math.max(range.start.row, range.end.row);
      const startCol = Math.min(range.start.col, range.end.col);
      const endCol = Math.max(range.start.col, range.end.col);
      
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          cells.push({ row, col });
        }
      }
    }
    
    return cells;
  }
  
  /**
   * Get style summary for current selection
   * Detects mixed states for toolbar UI (indeterminate checkboxes, etc.)
   */
  getStyleSummary(): SelectionStyleSummary {
    const cells = this.getSelectedCells();
    
    if (cells.length === 0) {
      return {};
    }
    
    // Collect styles from all cells
    const styles: (CellStyle | undefined)[] = cells.map(addr => 
      this.worksheet.getCellStyle(addr)
    );
    
    // Helper to check if property is consistent
    const getStyleState = <K extends keyof CellStyle>(
      prop: K
    ): StyleState<CellStyle[K]> => {
      const values = styles.map(s => s?.[prop]).filter(v => v !== undefined);
      
      if (values.length === 0) {
        return { value: undefined, isMixed: false };
      }
      
      const firstValue = values[0];
      const allSame = values.every(v => v === firstValue);
      
      return {
        value: allSame ? firstValue : undefined,
        isMixed: !allSame,
      };
    };
    
    return {
      fontFamily: getStyleState('fontFamily'),
      fontSize: getStyleState('fontSize'),
      bold: getStyleState('bold'),
      italic: getStyleState('italic'),
      underline: getStyleState('underline'),
      strikethrough: getStyleState('strikethrough'),
      color: getStyleState('color'),
      align: getStyleState('align'),
      valign: getStyleState('valign'),
      wrap: getStyleState('wrap'),
      rotation: getStyleState('rotation'),
      fill: getStyleState('fill'),
      border: getStyleState('border'),
      numberFormat: getStyleState('numberFormat'),
    };
  }
  
  /**
   * Check if current  selection is a single cell
   */
  isSingleCell(): boolean {
    return this.state.ranges.length === 1 &&
           this.state.primaryRange.start.row === this.state.primaryRange.end.row &&
           this.state.primaryRange.start.col === this.state.primaryRange.end.col;
  }
  
  /**
   * Get selection dimensions
   */
  getSelectionDimensions(): { rows: number; cols: number } {
    const range = this.state.primaryRange;
    
    return {
      rows: Math.abs(range.end.row - range.start.row) + 1,
      cols: Math.abs(range.end.col - range.start.col) + 1,
    };
  }
  
  /**
   * Check if address is within current selection
   */
  isSelected(address: Address): boolean {
    return this.state.ranges.some(range => {
      const startRow = Math.min(range.start.row, range.end.row);
      const endRow = Math.max(range.start.row, range.end.row);
      const startCol = Math.min(range.start.col, range.end.col);
      const endCol = Math.max(range.start.col, range.end.col);
      
      return address.row >= startRow && address.row <= endRow &&
             address.col >= startCol && address.col <= endCol;
    });
  }
}
