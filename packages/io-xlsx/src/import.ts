/**
 * High-Fidelity Excel Import API
 * 
 * Lightweight alternative to ExcelJS/XLSX.js with:
 * - Incremental parsing
 * - Viewport-based lazy loading
 * - Lazy sheet loading
 * - Minimal memory footprint
 */

import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Workbook, Worksheet } from '@cyber-sheet/core';
import { LightweightXLSXParser, XLSXParseOptions, XLSXMetadata } from './LightweightParser';

export interface ImportOptions extends XLSXParseOptions {
  /** Create Workbook instance automatically */
  createWorkbook?: boolean;
  
  /** Initial viewport to load (for large files) */
  initialViewport?: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  };
  
  /** Lazy load sheets (don't load until accessed) */
  lazySheets?: boolean;
  
  /** Memory limit in MB (auto-throttle parsing) */
  memoryLimit?: number;

  /** Number of worksheet rows yielded per progressive chunk */
  chunkRows?: number;
}

export interface ImportResult {
  /** Parsed workbook */
  workbook: Workbook;
  
  /** Metadata from file */
  metadata: XLSXMetadata;
  
  /** Lazy sheet loader */
  loadSheet: (nameOrIndex: string | number, options?: XLSXParseOptions) => Promise<Worksheet>;
  
  /** Expand viewport for already-loaded sheet */
  expandViewport: (
    sheetNameOrIndex: string | number,
    newViewport: { startRow: number; endRow: number; startCol: number; endCol: number }
  ) => Promise<void>;
}

function sheetDisplayDimensions(dims: { rows: number; cols: number }): { rows: number; cols: number } {
  return {
    rows: Math.max(DEFAULT_WORKSHEET_ROWS, dims.rows),
    cols: Math.max(DEFAULT_WORKSHEET_COLS, dims.cols),
  };
}

/**
 * Lazy-loading worksheet wrapper
 * Loads cells on-demand as viewport expands
 */
class LazyWorksheet extends Worksheet {
  private parser: LightweightXLSXParser;
  private sheetIndex: number;
  private loadedViewport: { startRow: number; endRow: number; startCol: number; endCol: number } | null = null;
  
  constructor(
    name: string,
    rows: number,
    cols: number,
    parser: LightweightXLSXParser,
    sheetIndex: number
  ) {
    super(name, rows, cols);
    this.parser = parser;
    this.sheetIndex = sheetIndex;
  }
  
  /**
   * Load cells in viewport
   */
  async loadViewport(viewport: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  }): Promise<void> {
    const cells = await this.parser.parseSheet(this.sheetIndex, {
      viewport,
      includeStyles: true,
      includeFormulas: true
    });
    
    // Apply cells to worksheet
    for (const [ref, cell] of cells) {
      const match = ref.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      
      const colStr = match[1];
      const row = parseInt(match[2]);
      
      let col = 0;
      for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
      }
      
      const addr = { row, col };
      
      if (cell.value !== null && cell.value !== undefined) {
        this.setCellValue(addr, cell.value);
      }
      
      if (cell.formula) {
        // Store formula in cell
        const k = `${addr.row}:${addr.col}`;
        const existingCell = (this as any).cells.get(k) || { value: null };
        existingCell.formula = cell.formula;
        (this as any).cells.set(k, existingCell);
      }
      
      if (cell.style) {
        this.setCellStyle(addr, cell.style);
      }
    }
    
    this.loadedViewport = viewport;
  }
  
  /**
   * Expand loaded viewport if needed
   */
  async ensureViewportLoaded(viewport: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  }): Promise<void> {
    // Check if viewport is already loaded
    if (this.loadedViewport) {
      const needsExpansion = 
        viewport.startRow < this.loadedViewport.startRow ||
        viewport.endRow > this.loadedViewport.endRow ||
        viewport.startCol < this.loadedViewport.startCol ||
        viewport.endCol > this.loadedViewport.endCol;
      
      if (!needsExpansion) {
        return; // Already loaded
      }
      
      // Expand viewport
      const expandedViewport = {
        startRow: Math.min(viewport.startRow, this.loadedViewport.startRow),
        endRow: Math.max(viewport.endRow, this.loadedViewport.endRow),
        startCol: Math.min(viewport.startCol, this.loadedViewport.startCol),
        endCol: Math.max(viewport.endCol, this.loadedViewport.endCol)
      };
      
      await this.loadViewport(expandedViewport);
    } else {
      // First load
      await this.loadViewport(viewport);
    }
  }
  
  /**
   * Get loaded viewport
   */
  getLoadedViewport() {
    return this.loadedViewport;
  }
}

/**
 * Import Excel file with high-fidelity parsing and lazy loading
 */
export async function importXLSX(
  buffer: ArrayBuffer,
  options: ImportOptions = {}
): Promise<ImportResult> {
  const parser = new LightweightXLSXParser();
  
  // Step 1: Parse metadata (fast - no cell data)
  const metadata = await parser.parseMetadata(buffer);
  
  // Step 2: Create workbook
  const workbook = new Workbook();
  const lazySheets = new Map<string, LazyWorksheet>();
  
  // Step 3: Create worksheets (lazy or eager)
  for (let i = 0; i < metadata.sheetNames.length; i++) {
    const sheetName = metadata.sheetNames[i];
    const dims = metadata.sheetDimensions.get(sheetName) || { rows: 1000, cols: 26 };
    const displayDims = sheetDisplayDimensions(dims);
    
    if (options.lazySheets) {
      // Create lazy worksheet
      const sheet = new LazyWorksheet(
        sheetName,
        displayDims.rows,
        displayDims.cols,
        parser,
        i
      );
      
      // Add sheet to workbook by creating it first, then replacing
      workbook.addSheet(sheetName, displayDims.rows, displayDims.cols);
      const addedSheet = workbook.getSheet(sheetName)!;
      // Copy lazy sheet reference
      lazySheets.set(sheetName, sheet);
      // Replace with lazy version
      (workbook as any).sheets.set(sheetName, sheet);
    } else {
      // Load sheet immediately
      const sheet = workbook.addSheet(sheetName, displayDims.rows, displayDims.cols);
      
      // Load cells
      const cells = await parser.parseSheet(i, {
        viewport: options.initialViewport,
        maxRows: options.maxRows,
        maxCols: options.maxCols,
        includeStyles: options.includeStyles !== false,
        includeFormulas: options.includeFormulas !== false,
        onProgress: options.onProgress
      });
      
      // Apply cells
      for (const [ref, cell] of cells) {
        const match = ref.match(/^([A-Z]+)(\d+)$/);
        if (!match) continue;
        
        const colStr = match[1];
        const row = parseInt(match[2]);
        
        let col = 0;
        for (let i = 0; i < colStr.length; i++) {
          col = col * 26 + (colStr.charCodeAt(i) - 64);
        }
        
        const addr = { row, col };
        
        if (cell.value !== null && cell.value !== undefined) {
          sheet.setCellValue(addr, cell.value);
        }
        
        if (cell.formula) {
          // Store formula in cell
          const k = `${addr.row}:${addr.col}`;
          const existingCell = (sheet as any).cells.get(k) || { value: null };
          existingCell.formula = cell.formula;
          (sheet as any).cells.set(k, existingCell);
        }
        
        if (cell.style) {
          sheet.setCellStyle(addr, cell.style);
        }
      }
    }
  }
  
  // Helper functions
  const loadSheet = async (
    nameOrIndex: string | number,
    loadOptions?: XLSXParseOptions
  ): Promise<Worksheet> => {
    const sheetName = typeof nameOrIndex === 'number' 
      ? metadata.sheetNames[nameOrIndex]
      : nameOrIndex;
    
    const lazySheet = lazySheets.get(sheetName);
    if (lazySheet) {
      // Load viewport if not already loaded
      if (!lazySheet.getLoadedViewport() && options.initialViewport) {
        await lazySheet.loadViewport(options.initialViewport);
      }
      return lazySheet;
    }
    
    // Sheet already loaded eagerly
    const sheet = workbook.getSheet(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    return sheet;
  };
  
  const expandViewport = async (
    sheetNameOrIndex: string | number,
    newViewport: { startRow: number; endRow: number; startCol: number; endCol: number }
  ): Promise<void> => {
    const sheetName = typeof sheetNameOrIndex === 'number'
      ? metadata.sheetNames[sheetNameOrIndex]
      : sheetNameOrIndex;
    
    const lazySheet = lazySheets.get(sheetName);
    if (lazySheet) {
      await lazySheet.ensureViewportLoaded(newViewport);
    } else {
      throw new Error('Sheet is not lazy-loaded, cannot expand viewport');
    }
  };
  
  return {
    workbook,
    metadata,
    loadSheet,
    expandViewport
  };
}

/**
 * Streaming import for very large files
 * Yields chunks of cells as they're parsed
 */
export async function* streamXLSX(
  buffer: ArrayBuffer,
  options: ImportOptions = {}
): AsyncGenerator<{
  sheetName: string;
  cells: Map<string, { value: any; style?: any; formula?: string }>;
  progress: { row: number; total: number };
}> {
  const parser = new LightweightXLSXParser();
  const metadata = await parser.parseMetadata(buffer);
  
  // Stream each sheet
  for (let i = 0; i < metadata.sheetNames.length; i++) {
    const sheetName = metadata.sheetNames[i];
    
    // Should we process this sheet?
    if (options.sheets) {
      const shouldProcess = 
        (typeof options.sheets[0] === 'string' && (options.sheets as string[]).includes(sheetName)) ||
        (typeof options.sheets[0] === 'number' && (options.sheets as number[]).includes(i));
      
      if (!shouldProcess) continue;
    }
    
    // Parse sheet in chunks
    const chunkSize = options.chunkRows ?? 500;
    const dims = metadata.sheetDimensions.get(sheetName) || { rows: 1000, cols: 26 };
    
    for (let startRow = 1; startRow <= dims.rows; startRow += chunkSize) {
      const endRow = Math.min(startRow + chunkSize - 1, dims.rows);
      
      const cells = await parser.parseSheet(i, {
        viewport: {
          startRow,
          endRow,
          startCol: 1,
          endCol: dims.cols
        },
        includeStyles: options.includeStyles !== false,
        includeFormulas: options.includeFormulas !== false
      });
      
      yield {
        sheetName,
        cells,
        progress: {
          row: endRow,
          total: dims.rows
        }
      };
    }
  }
}

/**
 * Quick metadata preview without loading any cells
 */
export async function previewXLSX(buffer: ArrayBuffer): Promise<{
  sheetNames: string[];
  sheetDimensions: Map<string, { rows: number; cols: number }>;
  fileSize: number;
  estimatedCells: number;
}> {
  const parser = new LightweightXLSXParser();
  const metadata = await parser.parseMetadata(buffer);
  
  let estimatedCells = 0;
  for (const dims of metadata.sheetDimensions.values()) {
    estimatedCells += dims.rows * dims.cols;
  }
  
  return {
    sheetNames: metadata.sheetNames,
    sheetDimensions: metadata.sheetDimensions,
    fileSize: metadata.fileSize,
    estimatedCells
  };
}

export { LightweightXLSXParser };
export type { XLSXMetadata, XLSXParseOptions };
