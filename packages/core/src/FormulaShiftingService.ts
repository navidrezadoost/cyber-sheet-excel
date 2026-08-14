/**
 * FormulaShiftingService.ts
 * 
 * Token-based formula reference transformation for copy/paste operations.
 * 
 * Architecture: Token-based (NOT regex-based)
 * - Tokenize → Transform → Rebuild pipeline
 * - O(n) single-pass lexer
 * - Handles nested functions, string literals, ranges
 * - Extensible for future structured references
 * 
 * Phase: Clipboard Architecture V1 - Phase 0.1
 * Baseline: BASELINE_STABLE_V1
 */

import type { Address } from './types';

/**
 * Excel limits (0-indexed internally)
 * - MAX_COL: 16383 (Column XFD)
 * - MAX_ROW: 1048575 (Row 1048576 in 1-indexed Excel)
 */
const MAX_COL = 16383;
const MAX_ROW = 1048575;
const MAX_RANGE_REMAP_SCAN = 10000;

/**
 * Token types for formula parsing
 * 
 * @extensibility Future tokens for:
 * - STRUCTURED_REF: Table[@Column] syntax
 * - NAMED_RANGE: User-defined names
 * - DYNAMIC_ARRAY: Spill range operators (#)
 */
export type Token =
  | CellRefToken
  | RangeToken
  | SymbolToken
  | NumberToken
  | TextToken
  | FunctionToken
  | SheetRefToken;

export interface CellRefToken {
  type: 'CELL_REF';
  value: string;      // Display string (e.g., "$A$1", "B2")
  row: number;        // 0-indexed row
  col: number;        // 0-indexed column
  rowAbs: boolean;    // Row is absolute ($)
  colAbs: boolean;    // Column is absolute ($)
}

export interface RangeToken {
  type: 'RANGE';
  start: CellRefToken;
  end: CellRefToken;
}

export interface SymbolToken {
  type: 'SYMBOL';
  value: string;      // Operators (+, -, *, /, etc.), delimiters, whitespace, or error tokens (#REF!)
}

export interface NumberToken {
  type: 'NUMBER';
  value: string;      // e.g., "123", "45.67"
}

export interface TextToken {
  type: 'TEXT';
  value: string;      // String literal with quotes: "Hello"
}

export interface FunctionToken {
  type: 'FUNCTION';
  value: string;      // Function name: SUM, IF, etc.
}

export interface SheetRefToken {
  type: 'SHEET_REF';
  sheet: string;      // Sheet name
  ref: CellRefToken | RangeToken;
}

type ReferenceMapper = (addr: Address) => Address | null;

/**
 * FormulaShiftingService - Token-based reference transformation
 * 
 * Critical Design Constraint:
 * - MUST NOT use regex for shifting (breaks on nested functions)
 * - MUST use token-based transformation (O(n) predictable)
 */
export class FormulaShiftingService {
  /**
   * Shift all cell references in a formula
   * 
   * Pipeline:
   * 1. Tokenize formula into structured tokens
   * 2. Transform CELL_REF and RANGE tokens
   * 3. Rebuild formula string from tokens
   * 
   * @param formula - Original formula string (with leading =)
   * @param sourceAddr - Original cell address
   * @param targetAddr - New cell address
   * @returns Shifted formula string
   * 
   * @example
   * shift("=A1+B2", {row: 0, col: 0}, {row: 1, col: 1})
   * // Returns: "=B2+C3"
   * 
   * @example
   * shift("=$A$1+B2", {row: 0, col: 0}, {row: 1, col: 1})
   * // Returns: "=$A$1+C3"
   */
  static shift(
    formula: string,
    sourceAddr: Address,
    targetAddr: Address
  ): string {
    const rowOffset = targetAddr.row - sourceAddr.row;
    const colOffset = targetAddr.col - sourceAddr.col;

    // Step 1: Tokenize
    const tokens = this.tokenize(formula);

    // Step 2: Transform references
    const transformedTokens = tokens.map(token =>
      this.transformToken(token, rowOffset, colOffset)
    );

    // Step 3: Rebuild formula
    return this.rebuild(transformedTokens);
  }

  /**
   * Adjust formula references after inserting rows in the worksheet coordinate
   * space. Absolute references still move structurally, but keep their $ flags.
   */
  static adjustForRowInsertion(formula: string, rowIndex: number, count = 1): string {
    if (count <= 0) return formula;
    return this.remapReferences(formula, addr => (
      addr.row >= rowIndex
        ? { row: addr.row + count, col: addr.col }
        : addr
    ));
  }

  /**
   * Adjust formula references after inserting columns in the worksheet coordinate
   * space. Absolute references still move structurally, but keep their $ flags.
   */
  static adjustForColumnInsertion(formula: string, colIndex: number, count = 1): string {
    if (count <= 0) return formula;
    return this.remapReferences(formula, addr => (
      addr.col >= colIndex
        ? { row: addr.row, col: addr.col + count }
        : addr
    ));
  }

  /**
   * Adjust references after moving a contiguous block of rows. `toIndex` is the
   * final start index of the moved block after reordering.
   */
  static adjustForRowReorder(
    formula: string,
    fromIndex: number,
    toIndex: number,
    count = 1
  ): string {
    if (count <= 0 || fromIndex === toIndex) return formula;
    return this.remapReferences(formula, addr => ({
      row: this.mapReorderedIndex(addr.row, fromIndex, toIndex, count),
      col: addr.col,
    }));
  }

  /**
   * Adjust references after moving a contiguous block of columns. `toIndex` is
   * the final start index of the moved block after reordering.
   */
  static adjustForColumnReorder(
    formula: string,
    fromIndex: number,
    toIndex: number,
    count = 1
  ): string {
    if (count <= 0 || fromIndex === toIndex) return formula;
    return this.remapReferences(formula, addr => ({
      row: addr.row,
      col: this.mapReorderedIndex(addr.col, fromIndex, toIndex, count),
    }));
  }

  /**
   * Map an index through a block move. The destination is the final start index
   * of the moved block.
   */
  static mapReorderedIndex(index: number, fromIndex: number, toIndex: number, count = 1): number {
    if (count <= 0 || fromIndex === toIndex) return index;
    const fromEnd = fromIndex + count;
    if (toIndex >= fromIndex && toIndex < fromEnd) return index;

    if (index >= fromIndex && index < fromEnd) {
      return toIndex + (index - fromIndex);
    }

    if (toIndex < fromIndex && index >= toIndex && index < fromIndex) {
      return index + count;
    }

    if (toIndex > fromIndex && index >= fromEnd && index < toIndex + count) {
      return index - count;
    }

    return index;
  }

  /**
   * Tokenize formula into structured tokens
   * O(n) single-pass lexer
   * 
   * Made public for dependency extraction in PasteCommand.
   * 
   * @param formula - Formula string (with or without leading =)
   * @returns Array of tokens
   */
  static tokenize(formula: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;

    // Skip leading '='
    if (formula[0] === '=') i = 1;

    while (i < formula.length) {
      const char = formula[i];

      // Whitespace - preserve as symbol for formatting stability
      if (char === ' ' || char === '\t' || char === '\n') {
        let ws = '';
        while (i < formula.length && (formula[i] === ' ' || formula[i] === '\t' || formula[i] === '\n')) {
          ws += formula[i];
          i++;
        }
        tokens.push({ type: 'SYMBOL', value: ws });
        continue;
      }

      // Cell reference pattern: $?[A-Z]+$?[0-9]+
      if (this.isCellRefStart(formula, i)) {
        const ref = this.parseCellRef(formula, i);
        
        // Check for range (next non-space char is ':')
        const nextNonSpace = this.skipWhitespace(formula, ref.endIndex);
        if (nextNonSpace < formula.length && formula[nextNonSpace] === ':') {
          // This is a range start
          const colonPos = nextNonSpace;
          const endRefPos = this.skipWhitespace(formula, colonPos + 1);
          
          if (this.isCellRefStart(formula, endRefPos)) {
            const endRef = this.parseCellRef(formula, endRefPos);
            tokens.push({
              type: 'RANGE',
              start: ref.token,
              end: endRef.token
            });
            i = endRef.endIndex;
            continue;
          }
        }
        
        tokens.push(ref.token);
        i = ref.endIndex;
        continue;
      }

      // Number: [0-9]+\.?[0-9]*
      if (this.isDigit(char)) {
        const num = this.parseNumber(formula, i);
        tokens.push({ type: 'NUMBER', value: num.value });
        i = num.endIndex;
        continue;
      }

      // Text literal: "..."
      if (char === '"') {
        const text = this.parseText(formula, i);
        tokens.push({ type: 'TEXT', value: text.value });
        i = text.endIndex;
        continue;
      }

      // Function name: [A-Z]+\(
      if (this.isUpperAlpha(char)) {
        const lookahead = this.peekFunction(formula, i);
        if (lookahead.isFunction) {
          tokens.push({ type: 'FUNCTION', value: lookahead.name });
          i = lookahead.endIndex;
          continue;
        }
      }

      // Symbol (operator, parenthesis, comma, etc.)
      tokens.push({ type: 'SYMBOL', value: char });
      i++;
    }

    return tokens;
  }

  /**
   * Transform token by shifting references
   */
  private static transformToken(
    token: Token,
    rowOffset: number,
    colOffset: number
  ): Token {
    if (token.type === 'CELL_REF') {
      return this.shiftCellRef(token, rowOffset, colOffset);
    }

    if (token.type === 'RANGE') {
      const shiftedStart = this.shiftCellRef(token.start, rowOffset, colOffset);
      const shiftedEnd = this.shiftCellRef(token.end, rowOffset, colOffset);
      
      // If either ref became #REF!, entire range is #REF!
      if (shiftedStart.type === 'SYMBOL' || shiftedEnd.type === 'SYMBOL') {
        return { type: 'SYMBOL', value: '#REF!' };
      }
      
      return {
        type: 'RANGE',
        start: shiftedStart as CellRefToken,
        end: shiftedEnd as CellRefToken
      };
    }

    // All other tokens pass through unchanged
    return token;
  }

  private static remapReferences(formula: string, mapper: ReferenceMapper): string {
    try {
      const tokens = this.tokenize(formula);
      return this.rebuild(tokens.map(token => this.remapToken(token, mapper)));
    } catch {
      return formula;
    }
  }

  private static remapToken(token: Token, mapper: ReferenceMapper): Token {
    if (token.type === 'CELL_REF') {
      return this.remapCellRef(token, mapper);
    }

    if (token.type === 'RANGE') {
      return this.remapRange(token, mapper);
    }

    if (token.type === 'SHEET_REF') {
      const ref = token.ref.type === 'RANGE'
        ? this.remapRange(token.ref, mapper)
        : this.remapCellRef(token.ref, mapper);

      if (ref.type === 'SYMBOL') return ref;
      return { ...token, ref };
    }

    return token;
  }

  private static remapRange(token: RangeToken, mapper: ReferenceMapper): RangeToken | SymbolToken {
    const start = this.remapCellRef(token.start, mapper);
    const end = this.remapCellRef(token.end, mapper);
    if (start.type === 'SYMBOL' || end.type === 'SYMBOL') {
      return { type: 'SYMBOL', value: '#REF!' };
    }

    const rowBounds = this.remapAxisBounds(token.start.row, token.end.row, row => {
      const mapped = mapper({ row, col: token.start.col });
      return mapped?.row ?? null;
    });
    const colBounds = this.remapAxisBounds(token.start.col, token.end.col, col => {
      const mapped = mapper({ row: token.start.row, col });
      return mapped?.col ?? null;
    });

    if (!rowBounds || !colBounds) return { type: 'SYMBOL', value: '#REF!' };

    const nextStart = this.rebuildCellRef(token.start, rowBounds.min, colBounds.min);
    const nextEnd = this.rebuildCellRef(token.end, rowBounds.max, colBounds.max);
    if (nextStart.type === 'SYMBOL' || nextEnd.type === 'SYMBOL') {
      return { type: 'SYMBOL', value: '#REF!' };
    }

    return {
      type: 'RANGE',
      start: nextStart,
      end: nextEnd,
    };
  }

  private static remapAxisBounds(
    start: number,
    end: number,
    mapper: (index: number) => number | null
  ): { min: number; max: number } | null {
    const lo = Math.min(start, end);
    const hi = Math.max(start, end);
    if (hi - lo > MAX_RANGE_REMAP_SCAN) {
      const mappedStart = mapper(lo);
      const mappedEnd = mapper(hi);
      if (mappedStart === null || mappedEnd === null) return null;
      return {
        min: Math.min(mappedStart, mappedEnd),
        max: Math.max(mappedStart, mappedEnd),
      };
    }

    let min = Infinity;
    let max = -Infinity;

    for (let i = lo; i <= hi; i++) {
      const mapped = mapper(i);
      if (mapped === null) return null;
      min = Math.min(min, mapped);
      max = Math.max(max, mapped);
    }

    return min === Infinity ? null : { min, max };
  }

  private static remapCellRef(token: CellRefToken, mapper: ReferenceMapper): CellRefToken | SymbolToken {
    const mapped = mapper({ row: token.row, col: token.col });
    if (!mapped) return { type: 'SYMBOL', value: '#REF!' };
    return this.rebuildCellRef(token, mapped.row, mapped.col);
  }

  private static rebuildCellRef(token: CellRefToken, row: number, col: number): CellRefToken | SymbolToken {
    if (row < 0 || row > MAX_ROW || col < 0 || col > MAX_COL) {
      return { type: 'SYMBOL', value: '#REF!' };
    }

    const colAbs = token.colAbs ? '$' : '';
    const rowAbs = token.rowAbs ? '$' : '';
    const colStr = this.columnIndexToLetters(col);
    const rowStr = (row + 1).toString();

    return {
      type: 'CELL_REF',
      value: `${colAbs}${colStr}${rowAbs}${rowStr}`,
      row,
      col,
      rowAbs: token.rowAbs,
      colAbs: token.colAbs,
    };
  }

  /**
   * Shift a single cell reference token
   * 
   * @returns Shifted CellRefToken or error SymbolToken (#REF!)
   */
  private static shiftCellRef(
    token: CellRefToken,
    rowOffset: number,
    colOffset: number
  ): CellRefToken | SymbolToken {
    let newRow = token.row;
    let newCol = token.col;

    // Shift column if not absolute
    if (!token.colAbs) {
      newCol += colOffset;
      if (newCol < 0 || newCol > MAX_COL) {
        return { type: 'SYMBOL', value: '#REF!' };
      }
    }

    // Shift row if not absolute
    if (!token.rowAbs) {
      newRow += rowOffset;
      if (newRow < 0 || newRow > MAX_ROW) {
        return { type: 'SYMBOL', value: '#REF!' };
      }
    }

    // Rebuild reference string
    const colAbs = token.colAbs ? '$' : '';
    const rowAbs = token.rowAbs ? '$' : '';
    const colStr = this.columnIndexToLetters(newCol);
    const rowStr = (newRow + 1).toString();  // 1-indexed display

    return {
      type: 'CELL_REF',
      value: `${colAbs}${colStr}${rowAbs}${rowStr}`,
      row: newRow,
      col: newCol,
      rowAbs: token.rowAbs,
      colAbs: token.colAbs
    };
  }

  /**
   * Rebuild formula string from tokens
   */
  private static rebuild(tokens: Token[]): string {
    let formula = '=';

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];

      if (token.type === 'RANGE') {
        formula += token.start.value + ':' + token.end.value;
      } else if (token.type === 'SHEET_REF') {
        // Sheet reference: Sheet1!A1
        const refStr = token.ref.type === 'RANGE' 
          ? `${token.ref.start.value}:${token.ref.end.value}`
          : token.ref.value;
        formula += `${token.sheet}!${refStr}`;
      } else {
        formula += token.value;
      }
    }

    return formula;
  }

  // ==================== LEXER HELPERS ====================

  /**
   * Check if position starts a cell reference
   * 
   * Cell ref pattern: $?[A-Z]+$?[0-9]+
   * Must have digits after column letters
   */
  private static isCellRefStart(formula: string, pos: number): boolean {
    let i = pos;
    const char = formula[i];
    
    // Check for optional column $
    if (char === '$') {
      i++;
      if (i >= formula.length || !this.isUpperAlpha(formula[i])) {
        return false;
      }
    } else if (!this.isUpperAlpha(char)) {
      return false;
    }
    
    // Skip column letters
    while (i < formula.length && this.isUpperAlpha(formula[i])) {
      i++;
    }
    
    // Check for optional row $
    if (i < formula.length && formula[i] === '$') {
      i++;
    }
    
    // MUST have at least one digit after column letters
    return i < formula.length && this.isDigit(formula[i]);
  }

  /**
   * Parse cell reference at position
   * Returns token and end position
   */
  private static parseCellRef(formula: string, pos: number): { token: CellRefToken; endIndex: number } {
    let i = pos;
    
    // Parse column absolute marker
    const colAbs = formula[i] === '$';
    if (colAbs) i++;
    
    // Parse column letters
    const colStart = i;
    while (i < formula.length && this.isUpperAlpha(formula[i])) {
      i++;
    }
    
    if (i === colStart) {
      throw new Error(`Expected column letters at position ${pos}`);
    }
    
    const colLetters = formula.substring(colStart, i);
    const col = this.columnLettersToIndex(colLetters);
    
    // Parse row absolute marker
    const rowAbs = i < formula.length && formula[i] === '$';
    if (rowAbs) i++;
    
    // Parse row number
    const rowStart = i;
    while (i < formula.length && this.isDigit(formula[i])) {
      i++;
    }
    
    if (i === rowStart) {
      throw new Error(`Expected row number at position ${pos}`);
    }
    
    const rowNum = parseInt(formula.substring(rowStart, i));
    const row = rowNum - 1;  // Convert to 0-indexed
    
    // Build display value
    const colAbsStr = colAbs ? '$' : '';
    const rowAbsStr = rowAbs ? '$' : '';
    const value = `${colAbsStr}${colLetters}${rowAbsStr}${rowNum}`;
    
    return {
      token: {
        type: 'CELL_REF',
        value,
        row,
        col,
        rowAbs,
        colAbs
      },
      endIndex: i
    };
  }

  /**
   * Parse number literal
   */
  private static parseNumber(formula: string, pos: number): { value: string; endIndex: number } {
    let i = pos;
    
    // Integer part
    while (i < formula.length && this.isDigit(formula[i])) {
      i++;
    }
    
    // Decimal part
    if (i < formula.length && formula[i] === '.') {
      i++;
      while (i < formula.length && this.isDigit(formula[i])) {
        i++;
      }
    }
    
    return {
      value: formula.substring(pos, i),
      endIndex: i
    };
  }

  /**
   * Parse text literal (string in quotes)
   */
  private static parseText(formula: string, pos: number): { value: string; endIndex: number } {
    let i = pos + 1;  // Skip opening quote
    let value = '"';
    
    while (i < formula.length) {
      const char = formula[i];
      
      // Handle escaped quotes ("")
      if (char === '"') {
        if (i + 1 < formula.length && formula[i + 1] === '"') {
          value += '""';
          i += 2;
        } else {
          // End of string
          value += '"';
          return { value, endIndex: i + 1 };
        }
      } else {
        value += char;
        i++;
      }
    }
    
    // Unterminated string - include what we have
    return { value, endIndex: i };
  }

  /**
   * Peek ahead to check if this is a function name
   * Function = uppercase letters followed by '('
   */
  private static peekFunction(formula: string, pos: number): { isFunction: boolean; name: string; endIndex: number } {
    let i = pos;
    
    // Read uppercase letters
    while (i < formula.length && this.isUpperAlpha(formula[i])) {
      i++;
    }
    
    if (i === pos) {
      return { isFunction: false, name: '', endIndex: pos };
    }
    
    // Skip whitespace
    const nextPos = this.skipWhitespace(formula, i);
    
    // Check for opening parenthesis
    if (nextPos < formula.length && formula[nextPos] === '(') {
      return {
        isFunction: true,
        name: formula.substring(pos, i),
        endIndex: i
      };
    }
    
    return { isFunction: false, name: '', endIndex: pos };
  }

  /**
   * Skip whitespace and return next non-space position
   */
  private static skipWhitespace(formula: string, pos: number): number {
    let i = pos;
    while (i < formula.length && (formula[i] === ' ' || formula[i] === '\t' || formula[i] === '\n')) {
      i++;
    }
    return i;
  }

  /**
   * Check if character is uppercase letter A-Z
   */
  private static isUpperAlpha(char: string): boolean {
    const code = char.charCodeAt(0);
    return code >= 65 && code <= 90;  // A-Z
  }

  /**
   * Check if character is digit 0-9
   */
  private static isDigit(char: string): boolean {
    const code = char.charCodeAt(0);
    return code >= 48 && code <= 57;  // 0-9
  }

  // ==================== COLUMN CONVERSION ====================

  /**
   * Convert column letters to 0-indexed column number
   * A=0, B=1, ..., Z=25, AA=26, etc.
   * 
   * @example
   * columnLettersToIndex("A")   // 0
   * columnLettersToIndex("Z")   // 25
   * columnLettersToIndex("AA")  // 26
   * columnLettersToIndex("AB")  // 27
   */
  static columnLettersToIndex(letters: string): number {
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 65 + 1);
    }
    return index - 1;  // Convert to 0-indexed
  }

  /**
   * Convert 0-indexed column number to letters
   * 0=A, 1=B, ..., 25=Z, 26=AA, etc.
   * 
   * @example
   * columnIndexToLetters(0)   // "A"
   * columnIndexToLetters(25)  // "Z"
   * columnIndexToLetters(26)  // "AA"
   * columnIndexToLetters(27)  // "AB"
   */
  static columnIndexToLetters(index: number): string {
    let letters = '';
    let num = index + 1;  // Convert to 1-indexed
    
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letters = String.fromCharCode(65 + remainder) + letters;
      num = Math.floor((num - 1) / 26);
    }
    
    return letters;
  }
}
