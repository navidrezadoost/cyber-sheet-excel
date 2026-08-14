/**
 * FormattingController.ts
 * 
 * High-level controller for cell formatting operations.
 * Provides methods for applying styles to selections with undo/redo support.
 * 
 * Architecture: Formatting Command Layer
 * - Integrates with CommandManager for undo/redo
 * - Handles batch formatting across multiple cells
 * - Provides format painter functionality
 * - Manages style cache for performance
 * 
 * Phase: UI Toolkit V1 - Formatting Operations
 */

import type { Address, Range, CellStyle } from './types';
import type { Worksheet } from './worksheet';
import type { Command, CommandManager } from './CommandManager';
import { BatchCommand } from './CommandManager';
import type { ClipboardService, ClipboardPayload } from './ClipboardService';

/**
 * Format painter state
 * Stores copied formatting to apply to target cells
 */
interface FormatPainterState {
  sourceStyles: Map<string, CellStyle>;
  isActive: boolean;
  isPersistent: boolean; // true if double-clicked (multi-apply mode)
}

/**
 * SetStyleCommand - Apply style to single cell with undo support
 */
class SetStyleCommand implements Command {
  description = 'Set cell style';
  
  private worksheet: Worksheet;
  private address: Address;
  private previousStyle: CellStyle | undefined;
  private newStyle: CellStyle | undefined;
  
  constructor(worksheet: Worksheet, address: Address, newStyle: CellStyle | undefined) {
    this.worksheet = worksheet;
    this.address = address;
    this.previousStyle = worksheet.getDirectCellStyle(address);
    this.newStyle = newStyle;
  }
  
  execute(): void {
    if (this.newStyle) {
      this.worksheet.setCellStyle(this.address, this.newStyle);
    }
  }
  
  undo(): void {
    if (this.previousStyle) {
      this.worksheet.setCellStyle(this.address, this.previousStyle);
    } else {
      this.worksheet.setCellStyle(this.address, undefined);
    }
  }
}

/**
 * BatchSetStyleCommand - Apply style to multiple cells atomically
 */
type StyleUpdater = CellStyle | ((prevStyle: CellStyle | undefined, addr: Address) => CellStyle);

class BatchSetStyleCommand implements Command {
  description = 'Set style for multiple cells';
  
  private commands: SetStyleCommand[] = [];
  
  constructor(
    worksheet: Worksheet,
    addresses: Address[],
    styleOrFn: StyleUpdater
  ) {
    for (const addr of addresses) {
      let newStyle: CellStyle;
      
      if (typeof styleOrFn === 'function') {
        const prevStyle = worksheet.getDirectCellStyle(addr);
        newStyle = styleOrFn(prevStyle, addr);
      } else {
        newStyle = styleOrFn;
      }
      
      this.commands.push(new SetStyleCommand(worksheet, addr, newStyle));
    }
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
 * MergeCellsCommand - Merge a rectangular range of cells
 */
class MergeCellsCommand implements Command {
  description = 'Merge cells';
  
  private worksheet: Worksheet;
  private range: Range;
  private removedCells: Map<string, { value: any; style: CellStyle | undefined }> = new Map();
  
  constructor(worksheet: Worksheet, range: Range) {
    this.worksheet = worksheet;
    this.range = range;
    
    // Store all non-anchor cell values and styles for undo
    const norm = this.normalizeRange(range);
    for (let r = norm.start.row; r <= norm.end.row; r++) {
      for (let c = norm.start.col; c <= norm.end.col; c++) {
        if (r === norm.start.row && c === norm.start.col) continue; // Skip anchor
        const key = `${r},${c}`;
        const addr = { row: r, col: c };
        this.removedCells.set(key, {
          value: worksheet.getCellValue(addr),
          style: worksheet.getDirectCellStyle(addr),
        });
      }
    }
  }
  
  private normalizeRange(range: Range): { start: Address; end: Address } {
    return {
      start: {
        row: Math.min(range.start.row, range.end.row),
        col: Math.min(range.start.col, range.end.col),
      },
      end: {
        row: Math.max(range.start.row, range.end.row),
        col: Math.max(range.start.col, range.end.col),
      },
    };
  }
  
  execute(): void {
    this.worksheet.mergeCells(this.range);
  }
  
  undo(): void {
    // First, cancel the merge
    this.worksheet.cancelMerge(this.range);
    
    // Then restore the removed cells' data
    for (const [key, data] of this.removedCells.entries()) {
      const [rowStr, colStr] = key.split(',');
      const addr = { row: parseInt(rowStr), col: parseInt(colStr) };
      
      if (data.value !== undefined) {
        this.worksheet.setCellValue(addr, data.value);
      }
      if (data.style) {
        this.worksheet.setCellStyle(addr, data.style);
      }
    }
  }
}

/**
 * UnmergeCellsCommand - Unmerge a range (remove merge overlapping with range)
 */
class UnmergeCellsCommand implements Command {
  description = 'Unmerge cells';
  
  private worksheet: Worksheet;
  private range: Range;
  private removedMerges: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }> = [];
  
  constructor(worksheet: Worksheet, range: Range) {
    this.worksheet = worksheet;
    this.range = range;
  }
  
  execute(): void {
    // Store removed merges for undo before canceling
    // This is a simplified implementation - full version would query mergeStore first
    // For now, cancelMerge handles removal
    this.worksheet.cancelMerge(this.range);
  }
  
  undo(): void {
    // Restore all merges that were removed
    for (const merge of this.removedMerges) {
      const mergeRange: Range = {
        start: { row: merge.startRow, col: merge.startCol },
        end: { row: merge.endRow, col: merge.endCol },
      };
      this.worksheet.mergeCells(mergeRange);
    }
  }
}

/**
 * FormattingController - High-level formatting operations
 * 
 * Responsibilities:
 * - Apply font formatting (family, size, bold, italic, underline, color)
 * - Apply alignment (horizontal, vertical, wrap text, merge cells)
 * - Apply fill/background colors
 * - Apply borders
 * - Apply number formats
 * - Format painter (copy/apply formatting)
 * - Batch operations with single undo entry
 * 
 * @architecture Command-based formatting layer
 * @invariant All operations create undoable commands
 * @invariant Worksheet modifications only through commands
 */
export class FormattingController {
  private worksheet: Worksheet;
  private commandManager: CommandManager;
  private formatPainterState: FormatPainterState = {
    sourceStyles: new Map(),
    isActive: false,
    isPersistent: false,
  };
  
  constructor(worksheet: Worksheet, commandManager: CommandManager) {
    this.worksheet = worksheet;
    this.commandManager = commandManager;
  }
  
  // ============================================================================
  // FONT FORMATTING
  // ============================================================================
  
  /**
   * Apply font family to selection
   */
  setFontFamily(addresses: Address[], fontFamily: string): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      fontFamily,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Apply font size to selection
   */
  setFontSize(addresses: Address[], fontSize: number): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      fontSize,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Toggle bold formatting across the selection.
   * Excel behavior: if every cell is bold, remove bold; otherwise apply bold to all.
   */
  toggleBold(addresses: Address[]): void {
    if (addresses.length === 0) return;
    const allBold = addresses.every((addr) => this.worksheet.getCellStyle(addr)?.bold === true);
    this.setBold(addresses, !allBold);
  }
  
  /**
   * Set bold formatting explicitly
   */
  setBold(addresses: Address[], bold: boolean): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      bold,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Toggle italic formatting across the selection uniformly.
   */
  toggleItalic(addresses: Address[]): void {
    if (addresses.length === 0) return;
    const allItalic = addresses.every((addr) => this.worksheet.getCellStyle(addr)?.italic === true);
    this.setItalic(addresses, !allItalic);
  }
  
  /**
   * Set italic formatting explicitly
   */
  setItalic(addresses: Address[], italic: boolean): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      italic,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Toggle underline formatting across the selection uniformly.
   */
  toggleUnderline(addresses: Address[]): void {
    if (addresses.length === 0) return;
    const allUnderline = addresses.every((addr) => this.worksheet.getCellStyle(addr)?.underline === true);
    this.setUnderline(addresses, !allUnderline);
  }
  
  /**
   * Set underline formatting explicitly
   */
  setUnderline(addresses: Address[], underline: boolean): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      underline,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Toggle strikethrough formatting
   */
  toggleStrikethrough(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      strikethrough: !prevStyle?.strikethrough,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Set font color
   */
  setFontColor(addresses: Address[], color: string): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      color,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  // ============================================================================
  // ALIGNMENT
  // ============================================================================
  
  /**
   * Set horizontal alignment
   */
  setHorizontalAlign(
    addresses: Address[],
    align: 'left' | 'center' | 'right' | 'fill' | 'justify'
  ): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      align,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Set vertical alignment
   */
  setVerticalAlign(
    addresses: Address[],
    valign: 'top' | 'middle' | 'bottom'
  ): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      valign,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Toggle wrap text
   */
  toggleWrapText(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      wrap: !prevStyle?.wrap,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Set text rotation (degrees)
   */
  setRotation(addresses: Address[], rotation: number): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      rotation,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Set text indentation
   */
  setIndent(addresses: Address[], indent: number): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      indent,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Increase indent level
   */
  increaseIndent(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const currentIndent = prevStyle?.indent || 0;
      return {
        ...(prevStyle || {}),
        indent: Math.min(currentIndent + 1, 250), // Excel max  indent is 250
      };
    });
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Decrease indent level
   */
  decreaseIndent(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const currentIndent = prevStyle?.indent || 0;
      return {
        ...(prevStyle || {}),
        indent: Math.max(currentIndent - 1, 0),
      };
    });
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Merge cells in a range
   * Only keeps the top-left cell's content
   * 
   * @param range - Range to merge
   * @throws MergeConflictError if any cell in range is already merged
   * @throws RangeError if range is single cell
   */
  mergeCells(range: Range): void {
    const cmd = new MergeCellsCommand(this.worksheet, range);
    this.commandManager.execute(cmd);
  }
  
  /**
   * Merge cells and center content
   * Convenience method that merges AND applies center alignment
   */
  mergeAndCenter(range: Range): void {
    // First merge
    this.mergeCells(range);
    
    // Then center the anchor cell
    const addresses = [range.start];
    this.setHorizontalAlign(addresses, 'center');
    this.setVerticalAlign(addresses, 'middle');
  }
  
  /**
   * Unmerge cells (cancel merge overlapping with range)
   */
  unmergeCells(range: Range): void {
    const cmd = new UnmergeCellsCommand(this.worksheet, range);
    this.commandManager.execute(cmd);
  }
  
  // ============================================================================
  // FILL / BACKGROUND
  // ============================================================================
  
  /**
   * Set cell background fill color
   */
  setFill(addresses: Address[], fill: string): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      fill,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Remove fill (make transparent)
   */
  removeFill(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const newStyle = { ...(prevStyle || {}) };
      delete newStyle.fill;
      return newStyle;
    });
    
    this.commandManager.execute(cmd);
  }
  
  // ============================================================================
  // BORDERS
  // ============================================================================
  
  /**
   * Set border for cells
   */
  setBorder(
    addresses: Address[],
    border: CellStyle['border']
  ): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      border,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Set all borders (outline + inside)
   */
  setAllBorders(addresses: Address[], color: string = '#000000'): void {
    const border = {
      top: color,
      right: color,
      bottom: color,
      left: color,
    };
    
    this.setBorder(addresses, border);
  }
  
  /**
   * Set outer border only
   */
  setOuterBorder(addresses: Address[], range: Range, color: string = '#000000'): void {
    // For outer border, we need to selectively apply to edge cells only
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);
    
    for (const addr of addresses) {
      const isTop = addr.row === startRow;
      const isBottom = addr.row === endRow;
      const isLeft = addr.col === startCol;
      const isRight = addr.col === endCol;
      
      const border: any = {};
      if (isTop) border.top = color;
      if (isBottom) border.bottom = color;
      if (isLeft) border.left = color;
      if (isRight) border.right = color;
      
      if (Object.keys(border).length > 0) {
        const cmd = new SetStyleCommand(this.worksheet, addr, {
          ...this.worksheet.getDirectCellStyle(addr),
          border,
        });
        this.commandManager.execute(cmd);
      }
    }
  }
  
  /**
   * Remove all borders
   */
  removeBorders(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const newStyle = { ...(prevStyle || {}) };
      delete newStyle.border;
      return newStyle;
    });
    
    this.commandManager.execute(cmd);
  }
  
  // ============================================================================
  // NUMBER FORMATS
  // ============================================================================
  
  /**
   * Set number format string
   */
  setNumberFormat(addresses: Address[], numberFormat: string): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      numberFormat,
    }));
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Quick format presets
   */
  applyNumberFormatPreset(
    addresses: Address[],
    preset: 'general' | 'number' | 'currency' | 'accounting' | 'percentage' | 'date' | 'time' | 'scientific' | 'fraction' | 'text'
  ): void {
    const formatMap: Record<string, string> = {
      general: 'General',
      number: '#,##0.00',
      currency: '$#,##0.00',
      accounting: '$#,##0.00;($#,##0.00)',
      percentage: '0.00%',
      date: 'M/D/YYYY',
      time: 'h:mm:ss AM/PM',
      scientific: '0.00E+00',
      fraction: '# ?/?',
      text: '@',
    };
    
    const format = formatMap[preset] || 'General';
    this.setNumberFormat(addresses, format);
  }
  
  /**
   * Increase decimal places
   */
  increaseDecimalPlaces(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const currentFormat = prevStyle?.numberFormat || '0';
      
      // Simple implementation: add one decimal place
      const match = currentFormat.match(/0\.(0*)/);
      if (match) {
        const decimals = match[1].length;
        const newFormat = currentFormat.replace(/0\.(0*)/, `0.${'0'.repeat(decimals + 1)}`);
        return { ...(prevStyle || {}), numberFormat: newFormat };
      }
      
      // Default: add .0
      return { ...(prevStyle || {}), numberFormat: currentFormat + '.0' };
    });
    
    this.commandManager.execute(cmd);
  }
  
  /**
   * Decrease decimal places
   */
  decreaseDecimalPlaces(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      const currentFormat = prevStyle?.numberFormat || '0';
      
      const match = currentFormat.match(/0\.(0+)/);
      if (match && match[1].length > 0) {
        const decimals = match[1].length;
        const newFormat = currentFormat.replace(/0\.0+/, `0.${'0'.repeat(Math.max(0, decimals - 1))}`);
        return { ...(prevStyle || {}), numberFormat: newFormat };
      }
      
      return prevStyle || {};
    });
    
    this.commandManager.execute(cmd);
  }
  
  // ============================================================================
  // FORMAT PAINTER
  // ============================================================================
  
  /**
   * Copy formatting from source cells
   * @param persistent If true, format painter stays active for multiple applications (double-click behavior)
   */
  copyFormat(addresses: Address[], persistent: boolean = false): void {
    this.formatPainterState.sourceStyles.clear();
    
    for (const addr of addresses) {
      const style = this.worksheet.getCellStyle(addr);
      const key = `${addr.row},${addr.col}`;
      this.formatPainterState.sourceStyles.set(key, style ? { ...style } : {});
    }
    
    this.formatPainterState.isActive = true;
    this.formatPainterState.isPersistent = persistent;
  }
  
  /**
   * Apply copied formatting to target cells
   */
  applyFormat(targetAddresses: Address[]): void {
    if (!this.formatPainterState.isActive || this.formatPainterState.sourceStyles.size === 0) {
      return;
    }
    
    // If we only have one source style, apply it to all targets
    const sourceStyles = Array.from(this.formatPainterState.sourceStyles.values());
    const sourceStyle = sourceStyles[0];
    
    if (sourceStyles.length === 1) {
      // Single source: apply same style to all targets
      const cmd = new BatchSetStyleCommand(this.worksheet, targetAddresses, sourceStyle);
      this.commandManager.execute(cmd);
    } else {
      // Multiple sources: apply pattern (tile source styles over target)
      const commands = targetAddresses.map((addr, index) => {
        const style = sourceStyles[index % sourceStyles.length];
        return new SetStyleCommand(this.worksheet, addr, style);
      });
      const cmd = new BatchCommand(commands, 'Format Painter');
      this.commandManager.execute(cmd);
    }
    
    // Clear format painter if not persistent
    if (!this.formatPainterState.isPersistent) {
      this.clearFormatPainter();
    }
  }
  
  /**
   * Cancel format painter
   */
  clearFormatPainter(): void {
    this.formatPainterState = {
      sourceStyles: new Map(),
      isActive: false,
      isPersistent: false,
    };
  }
  
  /**
   * Check if format painter is active
   */
  isFormatPainterActive(): boolean {
    return this.formatPainterState.isActive;
  }
  
  // ============================================================================
  // CLEAR FORMATTING
  // ============================================================================
  
  /**
   * Clear all formatting from cells (keep values)
   */
  clearFormat(addresses: Address[]): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, () => ({}));
    this.commandManager.execute(cmd);
  }
  
  /**
   * Clear specific format property
   */
  clearFormatProperty<K extends keyof CellStyle>(addresses: Address[], property: K): void {
    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => {
      if (!prevStyle) return {};
      
      const newStyle = { ...prevStyle };
      delete newStyle[property];
      return newStyle;
    });
    
    this.commandManager.execute(cmd);
  }

  /**
   * Apply a full cell style preset (Cell Styles gallery) in one undo step.
   */
  applyCellStylePreset(addresses: Address[], preset: CellStyle): void {
    if (addresses.length === 0) return;

    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle) => ({
      ...(prevStyle || {}),
      ...preset,
    }));

    this.commandManager.execute(cmd);
  }

  /**
   * Build a single undo step for Format as Table styling.
   */
  createTableStyleCommand(
    range: Range,
    options: {
      headerRowColor: string;
      firstRowStripedColor: string;
      secondRowStripedColor: string;
      borderColor?: string;
    }
  ): Command {
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);

    const addresses: Address[] = [];
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        addresses.push({ row, col });
      }
    }

    const borderColor = options.borderColor ?? '#BFBFBF';

    const cmd = new BatchSetStyleCommand(this.worksheet, addresses, (prevStyle, addr) => {
      const isHeader = addr.row === startRow;
      const isTop = addr.row === startRow;
      const isBottom = addr.row === endRow;
      const isLeft = addr.col === startCol;
      const isRight = addr.col === endCol;

      const nextStyle: CellStyle = { ...(prevStyle || {}) };

      if (isHeader) {
        nextStyle.fill = options.headerRowColor;
        nextStyle.bold = true;
        nextStyle.color = '#FFFFFF';
      } else {
        const dataRowIndex = addr.row - startRow - 1;
        nextStyle.fill = dataRowIndex % 2 === 0
          ? options.firstRowStripedColor
          : options.secondRowStripedColor;
      }

      if (isTop || isBottom || isLeft || isRight) {
        const border: NonNullable<CellStyle['border']> = {};
        if (isTop) border.top = borderColor;
        if (isBottom) border.bottom = borderColor;
        if (isLeft) border.left = borderColor;
        if (isRight) border.right = borderColor;
        nextStyle.border = border;
      }

      return nextStyle;
    });

    cmd.description = 'Format as table';
    return cmd;
  }

  /**
   * Apply Format as Table styling to a range in a single undo step.
   */
  applyTableStyle(
    range: Range,
    options: {
      headerRowColor: string;
      firstRowStripedColor: string;
      secondRowStripedColor: string;
      borderColor?: string;
    }
  ): void {
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);

    const addresses: Address[] = [];
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        addresses.push({ row, col });
      }
    }

    if (addresses.length === 0) return;

    this.commandManager.execute(this.createTableStyleCommand(range, options));
  }
}
