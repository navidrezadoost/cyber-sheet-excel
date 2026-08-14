import { Worksheet, Address, CellStyle, CellEvent, SheetEvents, resolveExcelColor, ExcelColorSpec, Emitter, assertInternedStyle, computeVerticalOffset, isHyperlinkVisited, HYPERLINK_COLOR, HYPERLINK_VISITED_COLOR, ConditionalFormattingEngine, type ConditionalFormattingRule, type ConditionalFormattingResult, type DataBarRender, type IconRender, renderIconOnCanvas, expandHeaderFooterText, sectionHasContent, type HeaderFooterSection, type CellValue } from '@cyber-sheet/core';
import { TextMeasureCache } from './TextMeasureCache';
import { Theme, ExcelLightTheme, mergeTheme, ThemePresetName, resolveThemePreset } from './Theme';
import { FormatCache } from './FormatCache';
import { RenderPlugin } from './plugins';
import {
  type ViewMode,
  getDefaultPageMetrics,
  computeRowPageBreaks,
  computeColPageBreaks,
} from './ViewModeRenderer';

export type CanvasRendererOptions = {
  headerHeight?: number; // px
  headerWidth?: number; // px
  // Theme tokens
  theme?: Partial<Theme>;
  debug?: boolean;
  onRender?: (info: { ms: number; regions: { x: number; y: number; w: number; h: number }[] }) => void;
  // Locale for ICU formatting (Intl)
  locale?: string;
  // Auto-size columns at init using measured text widths
  autoSizeColumns?: boolean | {
    mode?: 'visible' | 'all';
    includeHeader?: boolean;
    padding?: number; // px added to measured text
    minWidth?: number; // px lower bound
    maxWidth?: number; // px upper bound
    maxRowsToScan?: number; // when mode='all', cap scanning for performance
  };
  onRequestColumnFilterMenu?: (info: { col: number; anchor: { x: number; y: number };
    values: Array<{ value: string; count: number }>; apply: (selected: string[]) => void; clear: () => void; }) => void;
  formulaWorker?: FormulaWorkerBridge;
  formulaWorkerFactory?: () => FormulaWorkerBridge;
  formulaWorkerThreshold?: number;
};

export type FormulaWorkerBridge = {
  setCellValue(row: number, col: number, value: CellValue): Promise<void>;
  setCellFormula(row: number, col: number, formula: string, displayValue?: CellValue): Promise<void>;
  evaluateBatch(addresses: Address[]): Promise<{ values: Float64Array; evaluated: number; hasCycles: boolean }>;
  terminate?: () => void;
};

export type RenderStage = 'background' | 'grid' | 'headers' | 'cells' | 'selection' | 'overlays' | 'after';
export type RenderLayer = {
  id: string;
  stage: RenderStage;
  zIndex?: number;
  draw: (ctx: CanvasRenderingContext2D, api: {
    width: number; height: number; dpr: number; options: Required<CanvasRendererOptions>;
    sheet: Worksheet; getScroll: () => { x: number; y: number };
    rectForRange: (r1: number, c1: number, r2: number, c2: number) => { x: number; y: number; w: number; h: number } | null;
  }) => void;
};

export class CanvasRenderer {
  private container: HTMLElement;
  private sheet: Worksheet;
  private options: Required<CanvasRendererOptions>;
  private theme: Theme = ExcelLightTheme;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private dpr = 1;
  private scrollX = 0;
  private scrollY = 0;
  private zoom = 1; // scaling factor for grid (headers remain fixed)
  private selection: { start: Address; end: Address } | null = null;
  private selections: { start: Address; end: Address }[] = [];
  private textCache = new TextMeasureCache();
  private formatCache = new FormatCache();
  private layers: RenderLayer[] = [];
  private dirty: { x: number; y: number; w: number; h: number } | null = null;
  private rafId: number | null = null;
  private _lastRenderMs = 0;
  private lastDirtyRects: { x: number; y: number; w: number; h: number }[] = [];
  private valueFmtCache = new Map<string, string>();
  private plugins: RenderPlugin[] = [];
  private heatmapRange: { min: number; max: number } | null = null;
  private cfEngine = new ConditionalFormattingEngine();
  private cachedCFRules: ConditionalFormattingRule[] = [];
  private hoveredCell: Address | null = null;
  private clickStartTime = 0;
  private clickStartCell: Address | null = null;
  private isDragging = false;
  private dragStartCell: Address | null = null;
  private resizeState: { type: 'col' | 'row'; index: number; startPos: number; startSize: number; applyToAll?: boolean } | null = null;
  private scrollEmitter = new Emitter<{ type: 'scroll'; scroll: { x: number; y: number; maxX: number; maxY: number } }>();
  private selectionEmitter = new Emitter<{ type: 'selection'; selections: { start: Address; end: Address }[] }>();
  
  // Mouse hover throttling (rAF-based to reduce RAM/CPU)
  private pendingMouseX = 0;
  private pendingMouseY = 0;
  private hoverProcessPending = false;
  private lastHoveredRow = -1;
  private lastHoveredCol = -1;
  private isVisible = true; // Track tab visibility to pause processing when hidden
  
  // Dirty-rectangle rendering: track cells that need redraw
  private dirtyCells = new Set<string>(); // Format: "row:col"
  private gridLinesCanvas: HTMLCanvasElement | null = null; // Static gridlines rendered once
  private gridLinesNeedRedraw = true; // Flag to redraw gridlines on scroll/resize
  private viewMode: ViewMode = 'normal';
  private readonly workerPendingCells = new Set<string>();
  private formulaWorker?: FormulaWorkerBridge;
  private formulaWorkerHydrated = false;
  private layoutCacheDirty = true;
  private visibleRowsCache: number[] = [];
  private rowTopsCache: number[] = [0];
  private colLeftsCache: number[] = [0];
  
  // Distinct value cache per column for filter menus

  private clearValueCacheForColumn(col?: number) {
    if (col != null) this._valueCache.delete(col); else this._valueCache.clear();
  }

  getColumnDistinctValues(col: number, options?: { domain?: 'all' | 'visible' | 'filtered'; search?: string }): Array<{ value: string; count: number }> {
    const domain = options?.domain ?? 'all';
    const searchRaw = options?.search?.toLowerCase() ?? '';
    let map: Map<string, number> | undefined;
    if (domain === 'all') {
      map = this._valueCache.get(col);
      if (!map) {
        map = new Map();
        for (let r = 1; r <= this.sheet.rowCount; r++) {
          const v = this.sheet.getCellValue({ row: r, col });
          const key = v == null ? '' : String(v);
          map.set(key, (map.get(key) || 0) + 1);
        }
        this._valueCache.set(col, map);
      }
    } else if (domain === 'visible') {
      map = new Map();
      const rows = this.sheet.getVisibleRowIndices();
      for (const r of rows) {
        const v = this.sheet.getCellValue({ row: r, col });
        const key = v == null ? '' : String(v);
        map.set(key, (map.get(key) || 0) + 1);
      }
    } else if (domain === 'filtered') {
      // Rows that pass other filters except this column's own filter
      map = new Map();
      const rows = (this.sheet as any).getVisibleRowIndicesExcluding?.(col) || [];
      for (const r of rows) {
        const v = this.sheet.getCellValue({ row: r, col });
        const key = v == null ? '' : String(v);
        map.set(key, (map.get(key) || 0) + 1);
      }
    }
  if (!map) return [];
  let entries = Array.from(map.entries());
    if (searchRaw) entries = entries.filter(([value]) => value.toLowerCase().includes(searchRaw));
    return entries.map(([value, count]) => ({ value, count })).sort((a,b)=>a.value.localeCompare(b.value));
  }
  // Distinct value cache per column for filter menus
  private _valueCache: Map<number, Map<string, number>> = new Map();

  constructor(container: HTMLElement, sheet: Worksheet, options: CanvasRendererOptions = {}) {
    this.container = container;
    this.sheet = sheet;
    this.options = {
      headerHeight: options.headerHeight ?? 24,
      headerWidth: options.headerWidth ?? 48,
      theme: options.theme ?? {},
      debug: options.debug ?? false,
      onRender: options.onRender,
      locale: options.locale ?? undefined,
      formulaWorker: options.formulaWorker,
      formulaWorkerFactory: options.formulaWorkerFactory,
      formulaWorkerThreshold: options.formulaWorkerThreshold ?? 10,
    } as Required<CanvasRendererOptions>;
    this.formulaWorker = options.formulaWorker;
    this.theme = mergeTheme(ExcelLightTheme, this.options.theme);
    this.canvas = document.createElement('canvas');
    const ctx = this.canvas.getContext('2d');
    if (!ctx) throw new Error('Canvas 2D context not available');
  this.ctx = ctx;
    this.container.appendChild(this.canvas);
  // Initialize formatting cache locale
  this.formatCache = new FormatCache(this.options.locale);
    this.observeResize();
    this.resize();
    if (this.options.debug) this.setDebugEnabled(true);
    // Optional auto-size pass before first draw
    if (options.autoSizeColumns) {
      const cfg = typeof options.autoSizeColumns === 'object' ? options.autoSizeColumns : {};
      try { this.autoSizeColumns(cfg); } catch {}
    }
    this.setupEventListeners();
    this.cachedCFRules = sheet.getConditionalFormattingRules?.() ?? [];
    // Invalidate distinct value cache on relevant sheet events
    this.sheet.on((ev: SheetEvents) => {
      const t = (ev as any).type;
      if (t === 'cell-changed') this.clearValueCacheForColumn((ev as any).address.col);
      else if (t === 'sheet-mutated' || t === 'filter-changed') {
        this.invalidateLayoutCache();
        this.cachedCFRules = this.sheet.getConditionalFormattingRules?.() ?? [];
        this.clearValueCacheForColumn();
        this.gridLinesNeedRedraw = true;
        this.scheduleRedraw();
      } else if (
        t === 'row-hidden' ||
        t === 'row-shown' ||
        t === 'col-hidden' ||
        t === 'col-shown' ||
        t === 'progressive-load-changed'
      ) {
        if (t !== 'progressive-load-changed') this.invalidateLayoutCache();
        this.gridLinesNeedRedraw = true;
        this.scheduleRedraw();
      }
      // Invalidate cell region for style/comment/icon changes so visual updates (color, borders, indicators) appear
      if (t === 'style-changed') {
        const a = (ev as any).address; this.invalidateRange(a.row, a.col, a.row, a.col);
      } else if (t === 'comment-added' || t === 'comment-updated' || t === 'comment-deleted' || t === 'icon-changed') {
        const a = (ev as any).address; this.invalidateRange(a.row, a.col, a.row, a.col);
      } else if (t === 'cell-component-changed') {
        const a = (ev as any).address; this.invalidateRange(a.row, a.col, a.row, a.col);
        this.scheduleRedraw();
      } else if (t === 'header-footer-changed') {
        this.scheduleRedraw();
      }
      if (t === 'cell-changed') {
        this.syncFormulaWorkerCell((ev as any).address, (ev as any).cell);
      }
    });
    this.redraw();
  }

  dispose() {
    // Clean up resize timer
    if (this.resizeTimer !== null) {
      clearTimeout(this.resizeTimer);
      this.resizeTimer = null;
    }
    this.resizeObserver?.disconnect();
    this.removeEventListeners();
    if (this.formulaWorker && this.formulaWorker !== this.options.formulaWorker) {
      this.formulaWorker.terminate?.();
      this.formulaWorker = undefined;
    }
    this.canvas.remove();
  }

  setScroll(x: number, y: number) {
    const max = this.getMaxScroll();
    // Clamp to bounds
    const clampedX = Math.min(Math.max(0, x), max.x);
    const clampedY = Math.min(Math.max(0, y), max.y);
    
    // Only redraw if scroll actually changed
    if (this.scrollX === clampedX && this.scrollY === clampedY) return;
    
    this.scrollX = clampedX;
    this.scrollY = clampedY;
    
    // Scrolling changes viewport - trigger full redraw
    this.gridLinesNeedRedraw = true;
    this.dirtyCells.clear(); // Clear dirty cells since we're redrawing everything
    
    // Emit scroll event for adapters to listen
    this.scrollEmitter.emit({
      type: 'scroll',
      scroll: { x: clampedX, y: clampedY, maxX: max.x, maxY: max.y }
    });
    
    // Schedule redraw via RAF to batch multiple scroll updates
    this.scheduleRedraw();
  }
  
  private scheduleRedraw() {
    if (this.rafId !== null) return; // Already scheduled
    this.rafId = requestAnimationFrame(() => {
      this.rafId = null;
      this.redraw();
    });
  }
  setSelection(sel: { start: Address; end: Address } | null) { 
    // Mark old selection dirty
    if (this.selection) this.markRangeDirty(this.selection.start, this.selection.end);
    this.selection = sel; 
    this.selections = sel ? [sel] : []; 
    // Mark new selection dirty
    if (sel) this.markRangeDirty(sel.start, sel.end);
    this.scheduleRedraw(); 
  }
  setSelections(ranges: { start: Address; end: Address }[]) { 
    // Check if selection actually changed to avoid unnecessary events
    const changed = ranges.length !== this.selections.length || 
      ranges.some((r, i) => {
        const old = this.selections[i];
        return !old || 
          r.start.row !== old.start.row || r.start.col !== old.start.col ||
          r.end.row !== old.end.row || r.end.col !== old.end.col;
      });
    
    // Mark old selections dirty
    for (const range of this.selections) this.markRangeDirty(range.start, range.end);
    this.selections = ranges.slice(); 
    this.selection = ranges[0] ?? null; 
    // Mark new selections dirty
    for (const range of ranges) this.markRangeDirty(range.start, range.end);
    this.scheduleRedraw(); 
    
    // Only emit selection change event if selection actually changed
    if (changed) {
      this.selectionEmitter.emit({ type: 'selection', selections: this.selections });
    }
  }
  getSelections(): { start: Address; end: Address }[] { return this.selections.slice(); }
  getScroll() { return { x: this.scrollX, y: this.scrollY }; }
  scrollBy(dx: number, dy: number) { this.setScroll(this.scrollX + dx, this.scrollY + dy); }
  
  /**
   * Subscribe to scroll change events. Returns a disposable to unsubscribe.
   * This is the proper way for adapters to listen to scroll changes instead of polling.
   */
  onScrollChange(listener: (event: { x: number; y: number; maxX: number; maxY: number }) => void): { dispose: () => void } {
    return this.scrollEmitter.on((event: any) => {
      if (event.type === 'scroll') {
        listener(event.scroll);
      }
    });
  }

  onSelectionChange(listener: (selections: { start: Address; end: Address }[]) => void): { dispose: () => void } {
    return this.selectionEmitter.on((event: any) => {
      if (event.type === 'selection') {
        listener(event.selections);
      }
    });
  }
  
  // Expose read-only handles for host wrappers (React) that need to align/snapping or edge checks
  get optionsReadonly(): Readonly<Required<CanvasRendererOptions>> { return this.options; }
  get sheetReadonly(): Worksheet { return this.sheet; }
  // The size of the drawable content viewport (excluding headers), in CSS pixels.
  getViewportSize(): { width: number; height: number } {
    const width = this.canvas.width / this.dpr;
    const height = this.canvas.height / this.dpr;
    return { width: Math.max(0, width - this.options.headerWidth), height: Math.max(0, height - this.options.headerHeight) };
  }
  // Visible rows considering filters (fallback to all rows)
  private getVisibleRows(): number[] {
    this.ensureLayoutCache();
    return this.visibleRowsCache;
  }

  private invalidateLayoutCache(): void {
    this.layoutCacheDirty = true;
  }

  private ensureLayoutCache(): void {
    if (!this.layoutCacheDirty) return;

    const anySheet: any = this.sheet as any;
    const rows: number[] = typeof anySheet.getVisibleRowIndices === 'function'
      ? anySheet.getVisibleRowIndices()
      : Array.from({ length: this.sheet.rowCount }, (_, i) => i + 1);

    const rowTops = new Array(rows.length + 1);
    rowTops[0] = 0;
    for (let i = 0; i < rows.length; i++) {
      rowTops[i + 1] = rowTops[i] + this.sheet.getRowHeight(rows[i]) * this.zoom;
    }

    const colLefts = new Array(this.sheet.colCount + 1);
    colLefts[0] = 0;
    for (let col = 1; col <= this.sheet.colCount; col++) {
      colLefts[col] = colLefts[col - 1] + this.sheet.getColumnWidth(col) * this.zoom;
    }

    this.visibleRowsCache = rows;
    this.rowTopsCache = rowTops;
    this.colLeftsCache = colLefts;
    this.layoutCacheDirty = false;
  }

  private upperBound(values: number[], target: number): number {
    let lo = 0;
    let hi = values.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (values[mid] <= target) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  private lowerBound(values: number[], target: number): number {
    let lo = 0;
    let hi = values.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (values[mid] < target) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  private columnIndexAt(contentX: number): number | null {
    this.ensureLayoutCache();
    if (contentX < 0 || contentX >= this.colLeftsCache[this.sheet.colCount]) return null;
    const col = this.upperBound(this.colLeftsCache, contentX);
    return col >= 1 && col <= this.sheet.colCount ? col : null;
  }

  private visibleRowIndexAt(contentY: number): number | null {
    this.ensureLayoutCache();
    if (contentY < 0 || contentY >= this.rowTopsCache[this.visibleRowsCache.length]) return null;
    const rowIndex = this.upperBound(this.rowTopsCache, contentY) - 1;
    return rowIndex >= 0 && rowIndex < this.visibleRowsCache.length ? rowIndex : null;
  }

  // The full content size based on row/column sizes (excluding headers), in CSS pixels.
  getContentSize(): { width: number; height: number } {
    this.ensureLayoutCache();
    return {
      width: this.colLeftsCache[this.sheet.colCount] ?? 0,
      height: this.rowTopsCache[this.visibleRowsCache.length] ?? 0,
    };
  }
  // The maximum scroll offsets allowed (content minus viewport), never negative.
  getMaxScroll(): { x: number; y: number } {
    const vp = this.getViewportSize();
    const ct = this.getContentSize();
    return { x: Math.max(0, ct.width - vp.width), y: Math.max(0, ct.height - vp.height) };
  }
  get lastRenderMs() { return this._lastRenderMs; }
  getTheme(): Theme { return this.theme; }
  setTheme(next: Partial<Theme>) { this.theme = mergeTheme(this.theme, next); this.ctx.font = `${this.theme.fontSize}px ${this.theme.fontFamily}`; this.redraw(); }
  setThemePreset(name: ThemePresetName, overrides?: Partial<Theme>) {
    const preset = resolveThemePreset(name);
    this.theme = overrides ? mergeTheme(preset, overrides) : preset;
    this.ctx.font = `${this.theme.fontSize}px ${this.theme.fontFamily}`;
    this.redraw();
  }
  setLocale(locale: string) {
    this.formatCache.setLocale(locale);
    this.valueFmtCache.clear(); // invalidate cached formatted strings
    this.redraw();
  }
  setZoom(factor: number) {
    const z = Math.min(4, Math.max(0.25, factor));
    if (z === this.zoom) return;
    this.zoom = z;
    this.invalidateLayoutCache();
    this.invalidateRect(0, 0, this.canvas.width / this.dpr, this.canvas.height / this.dpr);
  }
  getZoom() { return this.zoom; }
  setViewMode(mode: ViewMode) {
    if (this.viewMode === mode) return;
    this.viewMode = mode;
    this.gridLinesNeedRedraw = true;
    this.redraw();
  }
  getViewMode(): ViewMode {
    return this.viewMode;
  }
  /**
   * Compute and apply column widths based on measured text of cells.
   * By default scans visible rows for performance, includes header labels, and applies padding/minWidth.
   */
  autoSizeColumns(opts?: { mode?: 'visible'|'all'; includeHeader?: boolean; padding?: number; minWidth?: number; maxWidth?: number; maxRowsToScan?: number; }) {
    const mode = opts?.mode ?? 'visible';
    const includeHeader = opts?.includeHeader ?? true;
    const padding = opts?.padding ?? 12;
    const minWidth = opts?.minWidth ?? 48;
    const maxWidth = opts?.maxWidth ?? 800;
    const maxRowsToScan = opts?.maxRowsToScan ?? 5000;
    const ctx = this.ctx;
    const fontBaseSize = this.theme.fontSize;
    const fontFamily = this.theme.fontFamily;
    const measure = (text: string, style?: CellStyle) => {
      const size = style?.fontSize || fontBaseSize;
      const weight = style?.bold ? 'bold ' : '';
      const italic = style?.italic ? 'italic ' : '';
      const font = `${italic}${weight}${size}px ${fontFamily}`;
      let w = this.textCache.get(font, text);
      if (w == null) { ctx.save(); ctx.font = font; w = ctx.measureText(text).width; ctx.restore(); this.textCache.set(font, text, w); }
      return w;
    };
    const rows: number[] = mode === 'visible' ? this.sheet.getVisibleRowIndices() : Array.from({ length: this.sheet.rowCount }, (_, i) => i + 1);
    const limitedRows = rows.length > maxRowsToScan ? rows.slice(0, maxRowsToScan) : rows;
    for (let c = 1; c <= this.sheet.colCount; c++) {
      let maxW = includeHeader ? measure(this.colLabel(c)) : 0;
      for (const r of limitedRows) {
        const v = this.sheet.getCellValue({ row: r, col: c });
        if (v == null) continue;
        const style = this.sheet.getCellStyle({ row: r, col: c });
        // use formatted textual value if applicable
        let text: string;
        if (typeof v === 'number') {
          const fmt = style?.numberFormat;
          const res = this.formatCache.formatValue(v, fmt);
          text = res.text ?? String(v);
        } else {
          text = String(v);
        }
        const w = measure(text, style);
        if (w > maxW) maxW = w;
      }
      const finalW = Math.min(maxWidth, Math.max(minWidth, Math.ceil(maxW + padding)));
      this.sheet.setColumnWidth(c, finalW);
    }
    this.redraw();
  }
  addPlugin(plugin: RenderPlugin) {
    this.plugins.push(plugin);
    this.plugins.sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
    this.redraw();
  }
  removePlugin(id: string) {
    this.plugins = this.plugins.filter(p => p.id !== id);
    this.redraw();
  }
  getPlugin(id: string): RenderPlugin | undefined {
    return this.plugins.find(p => p.id === id);
  }
  computeHeatmapRange() {
    let min = Infinity, max = -Infinity;
    for (let r = 1; r <= this.sheet.rowCount; r++) {
      for (let c = 1; c <= this.sheet.colCount; c++) {
        const v = this.sheet.getCellValue({ row: r, col: c });
        if (typeof v === 'number') {
          min = Math.min(min, v);
          max = Math.max(max, v);
        }
      }
    }
    this.heatmapRange = min !== Infinity ? { min, max } : null;
    this.redraw();
  }
  
  /**
   * Resolve Excel color specification to CSS color string
   * Handles theme colors, indexed colors, RGB, and direct CSS strings
   */
  private resolveColor(color: string | ExcelColorSpec | undefined | null, defaultColor: string = '#000000'): string {
    if (!color) return defaultColor;
    if (typeof color === 'string') return color;
    return resolveExcelColor(color, { defaultColor });
  }

  private applyConditionalFormatting(
    value: unknown,
    addr: Address,
    baseStyle?: CellStyle,
  ): { result?: ConditionalFormattingResult; style?: CellStyle } {
    if (!this.cachedCFRules.length) return { style: baseStyle };

    const result = this.cfEngine.applyRules(value as any, this.cachedCFRules, {
      address: addr,
      valueRange: this.heatmapRange ?? undefined,
    });

    const mergedStyle = result.style
      ? { ...(baseStyle ?? {}), ...result.style }
      : baseStyle;
    return { result, style: mergedStyle };
  }

  private renderConditionalDataBar(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    w: number,
    h: number,
    dataBar: DataBarRender,
  ) {
    const margin = 3;
    const barHeight = Math.max(4, h - margin * 2);
    const barWidth = Math.max(0, (w - margin * 2) * Math.min(1, Math.max(0, dataBar.percent)));

    ctx.save();
    if (dataBar.gradient) {
      const grad = ctx.createLinearGradient(x, y, x + barWidth, y);
      grad.addColorStop(0, dataBar.color);
      grad.addColorStop(1, this.applyAlpha(dataBar.color, 0.65));
      ctx.fillStyle = grad;
    } else {
      ctx.fillStyle = dataBar.color;
    }
    ctx.fillRect(x + margin, y + (h - barHeight) / 2, barWidth, barHeight);
    ctx.restore();
  }

  private renderConditionalIcon(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    w: number,
    h: number,
    icon: IconRender,
  ) {
    const size = Math.min(14, h - 4);
    const px = x + 4;
    const py = y + (h - size) / 2;
    ctx.save();
    renderIconOnCanvas(ctx, icon.iconSet, icon.iconIndex, px, py, size);
    ctx.restore();
  }

  private renderCustomCellIcon(
    ctx: CanvasRenderingContext2D,
    component: import('@cyber-sheet/core').CustomCellComponent,
    x: number,
    y: number,
    drawW: number,
    drawH: number,
  ) {
    const size = (component.size ?? 16) * this.zoom;
    const pad = 2 * this.zoom;
    let ix = x + pad;
    let iy = y + pad;

    switch (component.position) {
      case 'right':
        ix = x + drawW - size - pad;
        iy = y + (drawH - size) / 2;
        break;
      case 'overlay':
        ix = x + (drawW - size) / 2;
        iy = y + (drawH - size) / 2;
        break;
      case 'left':
      default:
        ix = x + pad;
        iy = y + (drawH - size) / 2;
        break;
    }

    const iconType = (component.props?.iconType as string | undefined)
      ?? (typeof component.props?.source === 'string' && component.props.source.startsWith('http') ? 'url' : 'emoji');
    const source = String(component.props?.source ?? component.id);

    ctx.save();
    if (iconType === 'emoji') {
      ctx.font = `${size}px ${this.theme.fontFamily}`;
      ctx.fillText(source, ix, iy + size - 4);
    } else if (iconType === 'builtin') {
      if (source === 'warning') {
        ctx.fillStyle = '#ffcc00';
        ctx.beginPath();
        ctx.moveTo(ix + size / 2, iy);
        ctx.lineTo(ix + size, iy + size);
        ctx.lineTo(ix, iy + size);
        ctx.closePath();
        ctx.fill();
      } else {
        ctx.fillStyle = '#0078d4';
        ctx.beginPath();
        ctx.arc(ix + size / 2, iy + size / 2, size / 2, 0, Math.PI * 2);
        ctx.fill();
      }
    } else if (iconType === 'url') {
      const cacheKey = `img:${source}`;
      let img = (this as any)._imageCache?.get(cacheKey);
      if (!img) {
        img = new Image();
        img.crossOrigin = 'anonymous';
        img.src = source;
        (this as any)._imageCache = (this as any)._imageCache || new Map();
        (this as any)._imageCache.set(cacheKey, img);
        img.onload = () => { this.invalidateRect(x, y, drawW, drawH); };
      }
      if (img.complete && img.naturalWidth) {
        ctx.drawImage(img, ix, iy, size, size);
      }
    }
    ctx.restore();
  }

  private applyAlpha(color: string, alpha: number): string {
    if (color.startsWith('#') && color.length === 7) {
      const r = parseInt(color.slice(1, 3), 16);
      const g = parseInt(color.slice(3, 5), 16);
      const b = parseInt(color.slice(5, 7), 16);
      return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }
    return color;
  }
  
  /**
   * Mark a specific cell as dirty for incremental rendering
   * Only this cell will be redrawn on next frame (unless full redraw triggered)
   */
  markCellDirty(row: number, col: number): void {
    const key = `${row}:${col}`;
    this.dirtyCells.add(key);
  }
  
  /**
   * Mark a range of cells as dirty (for selection changes, etc.)
   */
  private markRangeDirty(start: Address, end: Address): void {
    const r1 = Math.min(start.row, end.row);
    const r2 = Math.max(start.row, end.row);
    const c1 = Math.min(start.col, end.col);
    const c2 = Math.max(start.col, end.col);
    
    // Avoid marking too many cells (fallback to full redraw for large ranges)
    const cellCount = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (cellCount > 100) {
      // For large selections, trigger full redraw
      this.gridLinesNeedRedraw = true;
      this.dirtyCells.clear();
      return;
    }
    
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        this.markCellDirty(r, c);
      }
    }
  }
  
  /**
   * Render static gridlines to offscreen canvas
   * Only called when gridLinesNeedRedraw is true (scroll, resize, zoom)
   */
  private renderGridLines(): void {
    if (!this.gridLinesCanvas) {
      this.gridLinesCanvas = document.createElement('canvas');
      this.gridLinesCanvas.width = this.canvas.width;
      this.gridLinesCanvas.height = this.canvas.height;
    }
    
    const ctx = this.gridLinesCanvas.getContext('2d')!;
    if (!ctx) return;
    
    const width = this.canvas.width / this.dpr;
    const height = this.canvas.height / this.dpr;
    const { headerHeight, headerWidth } = this.options;
    const { gridColor, sheetBg } = this.theme;
    
    ctx.save();
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.clearRect(0, 0, width, height);
    
    // Fill background
    ctx.fillStyle = sheetBg;
    ctx.fillRect(headerWidth, headerHeight, width - headerWidth, height - headerHeight);
    
    // Draw gridlines
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    
    const { firstRow, firstCol, xOffset, yOffset } = this.visibleRange();
    
    // Vertical gridlines (columns)
    let x = xOffset;
    let col = firstCol;
    while (x < width && col <= this.sheet.colCount) {
      const cw = this.sheet.getColumnWidth(col) * this.zoom;
      const px = Math.round(x + cw) + 0.5;
      ctx.beginPath();
      ctx.moveTo(px, headerHeight);
      ctx.lineTo(px, height);
      ctx.stroke();
      x += cw;
      col++;
    }
    
    // Horizontal gridlines (rows - respect filters)
    const visRows = this.getVisibleRows();
    let y = yOffset;
    let rowIndex = (this.visibleRange().firstRowIndex ?? 0);
    while (y < height && rowIndex < visRows.length) {
      const row = visRows[rowIndex];
      const rh = this.sheet.getRowHeight(row) * this.zoom;
      const py = Math.round(y + rh) + 0.5;
      ctx.beginPath();
      ctx.moveTo(headerWidth, py);
      ctx.lineTo(width, py);
      ctx.stroke();
      y += rh;
      rowIndex++;
    }
    
    ctx.restore();
  }
  
  addLayer(layer: RenderLayer) { this.layers.push(layer); this.layers.sort((a, b) => (stageOrder(a.stage) - stageOrder(b.stage)) || ((a.zIndex ?? 0) - (b.zIndex ?? 0))); this.redraw(); }
  removeLayer(id: string) { this.layers = this.layers.filter(l => l.id !== id); this.redraw(); }
  invalidateRange(r1: number, c1: number, r2: number, c2: number) {
    if (this.options.debug) {
      console.log('🎨 [CanvasRenderer] invalidateRange called:', `(${r1},${c1}) to (${r2},${c2})`);
    }
    const rect = this.rectForRange(r1, c1, r2, c2);
    if (rect) {
      if (this.options.debug) {
        console.log('🎨 [CanvasRenderer] Invalidating rect:', rect);
      }
      this.invalidateRect(rect.x, rect.y, rect.w, rect.h);
    } else if (this.options.debug) {
      console.warn('⚠️ [CanvasRenderer] rectForRange returned null!');
    }
  }
  invalidateRect(x: number, y: number, w: number, h: number) {
    const r = { x, y, w, h };
    this.dirty = this.dirty ? unionRect(this.dirty, r) : r;
    if (this.rafId == null) this.rafId = requestAnimationFrame(() => { this.rafId = null; this.redraw(); });
  }
  exportToDataURL(range?: { r1: number; c1: number; r2: number; c2: number }, type: 'image/png' | 'image/jpeg' = 'image/png'): string {
    return this.canvas.toDataURL(type);
  }
  setDebugEnabled(enabled: boolean) {
    if (enabled) {
      if (!this.layers.find(l => l.id === 'debug-overlay')) {
        this.addLayer({ id: 'debug-overlay', stage: 'after', zIndex: 9999, draw: (ctx) => {
          ctx.save();
          ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
          ctx.font = `10px ${this.theme.fontFamily}`;
          ctx.fillStyle = '#008000';
          const fps = this._lastRenderMs > 0 ? (1000 / this._lastRenderMs).toFixed(1) : '-';
          ctx.fillText(`render: ${this._lastRenderMs.toFixed(2)} ms  FPS~${fps}`, 8, 12);
          ctx.strokeStyle = 'rgba(255,0,128,0.7)';
          for (const r of this.lastDirtyRects) ctx.strokeRect(r.x, r.y, r.w, r.h);
          ctx.restore();
        }});
      }
    } else {
      this.removeLayer('debug-overlay');
    }
  }

  private resizeObserver?: ResizeObserver;
  private resizeTimer: number | null = null;
  
  private observeResize() {
    this.resizeObserver = new ResizeObserver(() => {
      // Debounce: wait 150ms after the LAST resize event before redrawing
      if (this.resizeTimer !== null) {
        clearTimeout(this.resizeTimer);
      }
      
      this.resizeTimer = window.setTimeout(() => {
        this.resize();
        this.scheduleRedraw();
        this.resizeTimer = null;
      }, 150);
    });
    this.resizeObserver.observe(this.container);
  }

  resize() {
    this.dpr = window.devicePixelRatio || 1;
    const w = this.container.clientWidth;
    const h = this.container.clientHeight;
    this.canvas.width = Math.max(1, Math.floor(w * this.dpr));
    this.canvas.height = Math.max(1, Math.floor(h * this.dpr));
    this.canvas.style.width = `${w}px`;
    this.canvas.style.height = `${h}px`;
    this.ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    this.ctx.font = `${this.theme.fontSize}px ${this.theme.fontFamily}`;
    
    // Resize changes gridlines and viewport - trigger full redraw
    this.gridLinesNeedRedraw = true;
    this.dirtyCells.clear();
    if (this.gridLinesCanvas) {
      this.gridLinesCanvas.width = this.canvas.width;
      this.gridLinesCanvas.height = this.canvas.height;
    }
  }

  private visibleRange(): { firstRow: number; firstCol: number; xOffset: number; yOffset: number; firstRowIndex?: number } {
    this.ensureLayoutCache();
    const col = this.columnIndexAt(this.scrollX) ?? this.sheet.colCount;
    const rowIndex = this.visibleRowIndexAt(this.scrollY) ?? Math.max(0, this.visibleRowsCache.length - 1);
    const colLeft = this.colLeftsCache[col - 1] ?? 0;
    const rowTop = this.rowTopsCache[rowIndex] ?? 0;
    const x = this.options.headerWidth - (this.scrollX - colLeft);
    const y = this.options.headerHeight - (this.scrollY - rowTop);
    const firstRow = this.visibleRowsCache[rowIndex] ?? 1;
    return { firstRow, firstCol: col, xOffset: x, yOffset: y, firstRowIndex: rowIndex };
  }

  private getVisibleCellAddresses(): Address[] {
    const addresses: Address[] = [];
    const viewport = this.getViewportSize();
    const width = viewport.width + this.options.headerWidth;
    const height = viewport.height + this.options.headerHeight;
    const visRows = this.getVisibleRows();
    const { firstCol, xOffset, yOffset, firstRowIndex } = this.visibleRange();

    let y = yOffset;
    let rowIndex = firstRowIndex ?? 0;
    while (y < height && rowIndex < visRows.length) {
      const row = visRows[rowIndex];
      let x = xOffset;
      let col = firstCol;
      while (x < width && col <= this.sheet.colCount) {
        addresses.push({ row, col });
        x += this.sheet.getColumnWidth(col) * this.zoom;
        col++;
      }
      y += this.sheet.getRowHeight(row) * this.zoom;
      rowIndex++;
    }

    return addresses;
  }

  private evaluateVisibleFormulaCells(): void {
    const visibleAddresses = this.getVisibleCellAddresses();
    const isDirty = (this.sheet as any).isDirty;
    const dirtyAddresses = typeof isDirty === 'function'
      ? visibleAddresses.filter(addr => isDirty.call(this.sheet, addr))
      : visibleAddresses;
    if (dirtyAddresses.length === 0) return;

    const worker = this.getFormulaWorker(true);
    const threshold = this.options.formulaWorkerThreshold;
    if (!worker || dirtyAddresses.length < threshold) {
      const evaluateBatch = (this.sheet as any).evaluateBatch;
      if (typeof evaluateBatch === 'function') evaluateBatch.call(this.sheet, dirtyAddresses);
      return;
    }

    const pending = dirtyAddresses.filter(addr => {
      const key = `${addr.row}:${addr.col}`;
      if (this.workerPendingCells.has(key)) return false;
      this.workerPendingCells.add(key);
      return true;
    });
    if (pending.length === 0) return;

    worker.evaluateBatch(pending).then(result => {
      const applyFormulaResults = (this.sheet as any).applyFormulaResults;
      if (typeof applyFormulaResults === 'function') {
        applyFormulaResults.call(this.sheet, pending, result.values);
      }
    }).catch(error => {
      console.warn('Formula worker batch failed; falling back to main-thread evaluation.', error);
      const evaluateBatch = (this.sheet as any).evaluateBatch;
      if (typeof evaluateBatch === 'function') evaluateBatch.call(this.sheet, pending);
    }).finally(() => {
      for (const addr of pending) this.workerPendingCells.delete(`${addr.row}:${addr.col}`);
      this.scheduleRedraw();
    });
  }

  private syncFormulaWorkerCell(addr: Address, cell: { value?: unknown; formula?: string }): void {
    const worker = this.getFormulaWorker(!!cell.formula);
    if (!worker) return;
    if (cell.formula) {
      worker.setCellFormula(addr.row, addr.col, cell.formula, typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean' || cell.value === null ? cell.value : undefined).catch(() => {});
    } else {
      const value = typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean' || cell.value === null ? cell.value : null;
      worker.setCellValue(addr.row, addr.col, value).catch(() => {});
    }
  }

  private hydrateFormulaWorker(worker: FormulaWorkerBridge): void {
    if (this.formulaWorkerHydrated) return;
    this.formulaWorkerHydrated = true;

    const forEachNonEmptyCell = (this.sheet as any).forEachNonEmptyCell;
    if (typeof forEachNonEmptyCell !== 'function') return;

    forEachNonEmptyCell.call(this.sheet, (row: number, col: number, cell: { value?: unknown; formula?: string }) => {
      const value = typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean' || cell.value === null
        ? cell.value
        : null;

      if (cell.formula) {
        worker.setCellFormula(row, col, cell.formula, value).catch(() => {});
      } else {
        worker.setCellValue(row, col, value).catch(() => {});
      }
    });
  }

  private getFormulaWorker(create: boolean): FormulaWorkerBridge | undefined {
    if (this.formulaWorker) {
      this.hydrateFormulaWorker(this.formulaWorker);
      return this.formulaWorker;
    }
    if (!create) return undefined;
    const factory = this.options.formulaWorkerFactory;
    if (typeof factory !== 'function') return undefined;
    try {
      const worker = factory();
      this.formulaWorker = worker;
      this.hydrateFormulaWorker(worker);
    } catch (error) {
      console.warn('Failed to create formula worker; using main-thread formula evaluation.', error);
    }
    return this.formulaWorker;
  }

  redraw() {
    if (this.options.debug) {
      console.log('🎨 [CanvasRenderer] redraw() called, dirty rect:', this.dirty, 'dirtyCells:', this.dirtyCells.size);
    }
    const t0 = performance.now();
    this.evaluateVisibleFormulaCells();
    
    // OPTIMIZATION: Dirty-cell-only rendering
    // If gridlines don't need redraw and we have a small set of dirty cells,
    // only redraw those specific cells for massive performance gain
    const canUseDirtyRectOptimization = 
      this.dirtyCells.size > 0 && 
      this.dirtyCells.size <= 20 && 
      !this.gridLinesNeedRedraw;
    
    if (canUseDirtyRectOptimization) {
      if (this.options.debug) {
        console.log('🎨 [CanvasRenderer] Using dirty-cell optimization for', this.dirtyCells.size, 'cells');
      }
      const dirtyCellsArray = Array.from(this.dirtyCells);
      this.dirtyCells.clear();
      
      // Render gridlines to offscreen canvas if needed
      if (!this.gridLinesCanvas || this.gridLinesCanvas.width !== this.canvas.width) {
        this.renderGridLines();
        this.gridLinesNeedRedraw = false;
      }
      
      // Copy gridlines from offscreen canvas (fast blit)
      const ctx = this.ctx;
      const width = this.canvas.width / this.dpr;
      const height = this.canvas.height / this.dpr;
      
      if (this.gridLinesCanvas) {
        ctx.save();
        ctx.clearRect(0, 0, width, height);
        ctx.drawImage(this.gridLinesCanvas, 0, 0);
        ctx.restore();
      }
      
      // Redraw only dirty cells
      // TODO: Implement incremental cell redraw
      // For now, fall back to full redraw to maintain correctness
      // This will be completed in next iteration
      this.gridLinesNeedRedraw = true; // Force full redraw for now
    } else {
      if (this.options.debug) {
        console.log('🎨 [CanvasRenderer] Using full redraw, dirty rect:', this.dirty);
      }
    }
    
    // BEGIN: Standard full-frame render
    // Begin frame: optionally keep cache; could purge if memory grows. For now, keep across frames.
  const { headerHeight, headerWidth } = this.options;
  const { gridColor, headerBg, headerFg, sheetBg, selectionColor } = this.theme;
    const ctx = this.ctx;
    const width = this.canvas.width / this.dpr;
    const height = this.canvas.height / this.dpr;

    // Split dirty to regions (content, top header, left header, corner)
    const dirty = this.dirty;
    const clips: ({ x: number; y: number; w: number; h: number } | null)[] = [];
    if (!dirty) clips.push(null);
    else {
      const content = { x: headerWidth, y: headerHeight, w: width - headerWidth, h: height - headerHeight };
      const topHdr = { x: headerWidth, y: 0, w: width - headerWidth, h: headerHeight };
      const leftHdr = { x: 0, y: headerHeight, w: headerWidth, h: height - headerHeight };
      const corner = { x: 0, y: 0, w: headerWidth, h: headerHeight };
      const inter = (a: any, b: any) => {
        const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
        const r = Math.min(a.x + a.w, b.x + b.w), bt = Math.min(a.y + a.h, b.y + b.h);
        const w = r - x, h = bt - y; return w > 0 && h > 0 ? { x, y, w, h } : null;
      };
      const c1 = inter(dirty, content); if (c1) clips.push(c1);
      const c2 = inter(dirty, topHdr); if (c2) clips.push(c2);
      const c3 = inter(dirty, leftHdr); if (c3) clips.push(c3);
      const c4 = inter(dirty, corner); if (c4) clips.push(c4);
    }
    this.lastDirtyRects = clips.filter((c): c is { x: number; y: number; w: number; h: number } => !!c);

    for (const clip of clips) {
      ctx.save();
      if (clip) { ctx.beginPath(); ctx.rect(clip.x, clip.y, clip.w, clip.h); ctx.clip(); ctx.clearRect(clip.x, clip.y, clip.w, clip.h); }
      else { ctx.clearRect(0, 0, width, height); }

      // background
      ctx.fillStyle = this.viewMode === 'pageLayout' ? '#BFBFBF' : sheetBg;
      if (clip) ctx.fillRect(clip.x, clip.y, clip.w, clip.h); else ctx.fillRect(0, 0, width, height);
      // headers background
  ctx.fillStyle = headerBg;
      ctx.fillRect(0, 0, width, headerHeight);
      ctx.fillRect(0, 0, headerWidth, height);
      
      // Draw select-all button (top-left corner) with triangle icon
      ctx.fillStyle = headerBg;
      ctx.fillRect(0, 0, headerWidth, headerHeight);
      ctx.strokeStyle = gridColor;
      ctx.strokeRect(0, 0, headerWidth, headerHeight);
      
      // Draw triangle icon (like Excel's select-all)
      ctx.fillStyle = '#666666';
      ctx.beginPath();
      const triSize = 6;
      const triX = headerWidth / 2;
      const triY = headerHeight / 2;
      ctx.moveTo(triX - triSize / 2, triY + triSize / 3);
      ctx.lineTo(triX + triSize / 2, triY + triSize / 3);
      ctx.lineTo(triX + triSize / 2, triY - triSize / 3);
      ctx.closePath();
      ctx.fill();
      
      this.drawLayers('background', ctx, width, height);

      // grid background
      if (this.viewMode === 'pageLayout') {
        ctx.fillStyle = '#BFBFBF';
        ctx.fillRect(headerWidth, headerHeight, width - headerWidth, height - headerHeight);
        this.drawViewModePageBackground(ctx, width, height);
      } else {
        ctx.fillStyle = sheetBg;
        ctx.fillRect(headerWidth, headerHeight, width - headerWidth, height - headerHeight);
      }

      // grid lines and headers
  ctx.strokeStyle = gridColor;
      ctx.lineWidth = 1;
      const { firstRow, firstCol, xOffset, yOffset } = this.visibleRange();
      this.drawLayers('grid', ctx, width, height);

      // Column headers
      let x = xOffset; let col = firstCol;
  ctx.fillStyle = headerFg;
      if (clip) { while (x + this.sheet.getColumnWidth(col) * this.zoom < clip.x && col <= this.sheet.colCount) { x += this.sheet.getColumnWidth(col) * this.zoom; col++; } }
      while (x < width && col <= this.sheet.colCount) {
        const cw = this.sheet.getColumnWidth(col) * this.zoom;
        if (clip && x > (clip.x + clip.w)) break;
  ctx.fillText(this.colLabel(col), x + cw / 2 - ctx.measureText(this.colLabel(col)).width / 2, headerHeight / 2 + this.theme.fontSize / 2 - 2);
        const px = Math.round(x + cw) + 0.5; ctx.beginPath(); ctx.moveTo(px, 0); ctx.lineTo(px, height); ctx.stroke();
        x += cw; col++;
      }

      // Row headers (respect filters)
      const visRows = this.getVisibleRows();
      let y = yOffset; let rowIndex = (this.visibleRange().firstRowIndex ?? 0);
      if (clip) { while (rowIndex < visRows.length && y + this.sheet.getRowHeight(visRows[rowIndex]) * this.zoom < clip.y) { y += this.sheet.getRowHeight(visRows[rowIndex]) * this.zoom; rowIndex++; } }
      while (y < height && rowIndex < visRows.length) {
        const row = visRows[rowIndex];
        const rh = this.sheet.getRowHeight(row) * this.zoom;
        if (clip && y > (clip.y + clip.h)) break;
        const label = String(row);
  ctx.fillText(label, headerWidth - 4 - ctx.measureText(label).width, y + rh / 2 + this.theme.fontSize / 2 - 2);
        const py = Math.round(y + rh) + 0.5; ctx.beginPath(); ctx.moveTo(0, py); ctx.lineTo(width, py); ctx.stroke();
        y += rh; rowIndex++;
      }
      this.drawLayers('headers', ctx, width, height);

      // Cells (respect filters)
      y = yOffset; rowIndex = (this.visibleRange().firstRowIndex ?? 0);
  if (clip) { while (rowIndex < visRows.length && y + this.sheet.getRowHeight(visRows[rowIndex]) * this.getZoom() < clip.y) { y += this.sheet.getRowHeight(visRows[rowIndex]) * this.getZoom(); rowIndex++; } }
      while (y < height && rowIndex < visRows.length) {
        const row = visRows[rowIndex];
        x = xOffset; col = firstCol;
        if (clip) { while (x + this.sheet.getColumnWidth(col) * this.zoom < clip.x && col <= this.sheet.colCount) { x += this.sheet.getColumnWidth(col) * this.zoom; col++; } }
        const rh = this.sheet.getRowHeight(row) * this.zoom;
        while (x < width && col <= this.sheet.colCount) {
          if (clip && y > (clip.y + clip.h)) break;
          const cw = this.sheet.getColumnWidth(col) * this.zoom;
          if (clip && x > (clip.x + clip.w)) break;
          const addr = { row, col };
          const merged = (this.sheet as any).getMergedRangeForCell?.(addr);
          let spanW = cw, spanH = rh, isAnchor = true;
          if (merged) {
            isAnchor = addr.row === merged.start.row && addr.col === merged.start.col;
            if (isAnchor) {
              spanW = 0; for (let c = merged.start.col; c <= merged.end.col; c++) spanW += this.sheet.getColumnWidth(c) * this.zoom;
              spanH = 0; for (let r2 = merged.start.row; r2 <= merged.end.row; r2++) spanH += this.sheet.getRowHeight(r2) * this.zoom;
            }
          }
          const v = this.sheet.getCellValue(addr);
          let style: CellStyle | undefined = this.sheet.getCellStyle(addr);
          const hyperlink = this.sheet.getHyperlink?.(addr);
          const rowLoaded = typeof (this.sheet as any).isRowLoaded === 'function'
            ? (this.sheet as any).isRowLoaded(row)
            : true;

          const { result: cfResult, style: cfStyle } = this.applyConditionalFormatting(v, addr, style);
          style = cfStyle;

          // Phase 1 UI: Validate style is interned (dev mode only)
          // Prevents ecosystem integration drift (React, XLSX, toolbar mutations)
          assertInternedStyle(style, 'CanvasRenderer.renderCells');

          // Apply plugin-based heatmap background
          let pluginBg: string | undefined;
          for (const plugin of this.plugins) {
            if (plugin.getCellBackground) {
              const bg = plugin.getCellBackground({ 
                addr, 
                value: v, 
                style, 
                min: this.heatmapRange?.min, 
                max: this.heatmapRange?.max 
              });
              if (bg) { pluginBg = bg; break; }
            }
          }
          
          if (isAnchor && pluginBg) {
            ctx.fillStyle = pluginBg;
            ctx.fillRect(x + 1, y + 1, (merged ? spanW : cw) - 2, (merged ? spanH : rh) - 2);
          } else if (isAnchor && style?.fill) {
            // Resolve Excel color to CSS string
            let fillColor = this.resolveColor(style.fill, '#FFFFFF');
            // Apply color transform plugins
            for (const plugin of this.plugins) {
              if (plugin.transformColor) {
                fillColor = plugin.transformColor(fillColor, { addr, value: v, style });
              }
            }
            ctx.fillStyle = fillColor;
            ctx.fillRect(x + 1, y + 1, (merged ? spanW : cw) - 2, (merged ? spanH : rh) - 2);
          }

          if (isAnchor && cfResult?.dataBar) {
            this.renderConditionalDataBar(ctx, x, y, merged ? spanW : cw, merged ? spanH : rh, cfResult.dataBar);
          }

          if (isAnchor && !rowLoaded) {
            const w = merged ? spanW : cw;
            const h = merged ? spanH : rh;
            ctx.save();
            ctx.fillStyle = '#F3F4F6';
            ctx.fillRect(x + 1, y + 1, w - 2, h - 2);
            ctx.fillStyle = '#D1D5DB';
            ctx.fillRect(x + 8, y + Math.max(6, h / 2 - 3), Math.max(12, w - 16), 6);
            ctx.restore();
          }

          if (style?.border) {
            const bw = merged ? spanW : cw; const bh = merged ? spanH : rh;
            if (!merged || isAnchor) {
              type BorderEdgeLike = string | ExcelColorSpec | { color?: string | ExcelColorSpec; style?: string };
              const resolveBorderEdge = (edge?: BorderEdgeLike | null) => {
                if (!edge) return null;
                const edgeColor = typeof edge === 'object' && 'color' in edge ? edge.color : edge;
                const edgeStyle = typeof edge === 'object' && 'style' in edge ? edge.style : 'thin';
                let resolved = this.resolveColor(edgeColor as string | ExcelColorSpec, '#000000');
                // Apply plugin transforms
                for (const plugin of this.plugins) {
                  if (plugin.transformColor) {
                    resolved = plugin.transformColor(resolved, { addr, value: v, style });
                  }
                }
                return { color: resolved, style: edgeStyle ?? 'thin' };
              };
              const configureBorderStroke = (edge: { color: string; style: string }) => {
                ctx.strokeStyle = edge.color;
                ctx.setLineDash([]);
                switch (edge.style) {
                  case 'medium':
                  case 'mediumDashed':
                  case 'mediumDashDot':
                  case 'mediumDashDotDot':
                    ctx.lineWidth = 2;
                    break;
                  case 'thick':
                    ctx.lineWidth = 3;
                    break;
                  case 'hairline':
                    ctx.lineWidth = 0.5;
                    break;
                  default:
                    ctx.lineWidth = 1;
                }
                if (edge.style === 'dashed' || edge.style === 'mediumDashed') ctx.setLineDash([6, 3]);
                if (edge.style === 'dotted') ctx.setLineDash([1, 2]);
                if (edge.style === 'dashDot' || edge.style === 'mediumDashDot') ctx.setLineDash([6, 3, 1, 3]);
                if (edge.style === 'dashDotDot' || edge.style === 'mediumDashDotDot') ctx.setLineDash([6, 3, 1, 3, 1, 3]);
              };
              const drawLine = (edgeSpec: BorderEdgeLike | undefined, draw: () => void) => {
                const edge = resolveBorderEdge(edgeSpec);
                if (!edge) return;
                if (edge.style === 'double') {
                  configureBorderStroke({ ...edge, style: 'thin' });
                  draw();
                  ctx.save();
                  ctx.translate(0, edgeSpec === style.border.left || edgeSpec === style.border.right ? 0 : 3);
                  draw();
                  ctx.restore();
                  return;
                }
                configureBorderStroke(edge);
                draw();
                ctx.setLineDash([]);
              };
              drawLine(style.border.top as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x, y + 0.5); ctx.lineTo(x + bw, y + 0.5); ctx.stroke(); });
              drawLine(style.border.bottom as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x, y + bh + 0.5); ctx.lineTo(x + bw, y + bh + 0.5); ctx.stroke(); });
              drawLine(style.border.left as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x + 0.5, y); ctx.lineTo(x + 0.5, y + bh); ctx.stroke(); });
              drawLine(style.border.right as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x + bw + 0.5, y); ctx.lineTo(x + bw + 0.5, y + bh); ctx.stroke(); });
              drawLine(style.border.diagonalDown as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x, y); ctx.lineTo(x + bw, y + bh); ctx.stroke(); });
              drawLine(style.border.diagonalUp as BorderEdgeLike, () => { ctx.beginPath(); ctx.moveTo(x, y + bh); ctx.lineTo(x + bw, y); ctx.stroke(); });
            }
          }

          if (isAnchor && v !== null && v !== undefined && v !== '') {
            const fontSize = style?.fontSize ?? this.theme.fontSize;
            const fontFamily = style?.fontFamily ?? this.theme.fontFamily;
            let font = `${style?.italic ? 'italic ' : ''}${style?.bold ? 'bold ' : ''}${fontSize}px ${fontFamily}`;
            
            // Apply custom font transform
            for (const plugin of this.plugins) {
              if (plugin.transformFont) {
                font = plugin.transformFont(font, { addr, value: v, style });
              }
            }
            
            ctx.font = font;
            // Per-cell formatted string cache: address + fmt + primitive identity
            const key = `${addr.row}:${addr.col}|${style?.numberFormat ?? ''}|${typeof v === 'number' ? v : String(v)}`;
            let cached = this.valueFmtCache.get(key);
            let fmtResult: { text: string; color?: string };
            if (cached) {
              fmtResult = JSON.parse(cached);
            } else {
              fmtResult = this.formatCache.formatValue(v, style?.numberFormat);
              this.valueFmtCache.set(key, JSON.stringify(fmtResult));
            }
            const text = fmtResult.text;
            // Apply color from format if present, else use cell style color, then resolve Excel colors and apply plugins
            let textColor = hyperlink?.target
              ? (isHyperlinkVisited(hyperlink.target) ? HYPERLINK_VISITED_COLOR : HYPERLINK_COLOR)
              : this.resolveColor(fmtResult.color ?? style?.color, '#000000');
            for (const plugin of this.plugins) {
              if (plugin.transformColor) {
                textColor = plugin.transformColor(textColor, { addr, value: v, style });
              }
            }
            ctx.fillStyle = textColor;
            const drawW = merged ? spanW : cw; const drawH = merged ? spanH : rh; const maxWidth = Math.max(0, drawW - 8);
            let tx = x + 4;
            // Compute vertical offset using layout function (pure layout concern)
            let ty = y + computeVerticalOffset(style?.valign, drawH, fontSize, fontSize, 2, 4) - fontSize / 2 + 2;
            const align = style?.align ?? this.formatCache.preferredAlign(v, style?.numberFormat);
            let textWidth = this.textCache.get(font, text); if (textWidth === undefined) { textWidth = ctx.measureText(text).width; this.textCache.set(font, text, textWidth); }
            // Shrink-to-fit scaling if specified (tolerate style flag if present)
            const shrinkFlag = (style as any)?.shrinkToFit;
            const shrink = shrinkFlag ? this.formatCache.getTextScale(font, text, maxWidth, textWidth as number) : 1;
            if (shrink !== 1) {
              ctx.save();
              const cx = x + (drawW / 2);
              const cy = y + (drawH / 2);
              ctx.translate(cx, cy);
              ctx.scale(shrink, shrink);
              // After scaling, recompute tx relative to scaled center
              let localTx = -drawW / 2 + 4;
              if (align === 'right') localTx = drawW / 2 - 4 - (textWidth as number);
              else if (align === 'center') localTx = -(textWidth as number) / 2;
              // Compute vertical offset using layout function (scaled context)
              const valignOffset = computeVerticalOffset(style?.valign, drawH, fontSize, fontSize, 2, 4);
              let localTy = valignOffset - drawH / 2 - fontSize / 2 + 2;
              ctx.beginPath(); ctx.rect(-drawW / 2 + 1, -drawH / 2 + 1, drawW - 2, drawH - 2); ctx.clip();
              ctx.fillText(text, localTx, localTy, maxWidth);
              ctx.restore();
              x += this.sheet.getColumnWidth(col); col++;
              continue;
            }
            // Phase 1 UI: Apply indent (left-align only)
            let indentOffset = 0;
            if (style?.indent && align === 'left') {
              indentOffset = (style.indent) * 8; // ~8px per indent level
            }
            
            if (align === 'right') tx = x + drawW - 4 - (textWidth as number) + indentOffset; else if (align === 'center') tx = x + drawW / 2 - (textWidth as number) / 2; else if (indentOffset > 0) tx += indentOffset;
            const wrap = !!style?.wrap; const overflow = style?.textOverflow ?? 'clip';
            ctx.save(); ctx.beginPath(); ctx.rect(x + 1, y + 1, drawW - 2, drawH - 2); ctx.clip();
            
            // Phase 1 UI: Superscript/Subscript font scaling (fast path: check before computing)
            const hasScript = style?.superscript || style?.subscript;
            let scriptScale = 1;
            let scriptOffsetY = 0;
            if (hasScript) {
              scriptScale = 0.7; // 70% of normal font size
              const metrics = ctx.measureText(text);
              const ascent = metrics.actualBoundingBoxAscent || fontSize * 0.8;
              scriptOffsetY = style.superscript ? -ascent * 0.4 : ascent * 0.2;
            }
            
            if (style?.rotation && style.rotation !== 0) {
              const angle = (style.rotation * Math.PI) / 180; const cx = x + (drawW / 2); const cy = y + (drawH / 2);
              ctx.save(); ctx.translate(cx, cy); ctx.rotate(angle); 
              if (hasScript) {
                ctx.save();
                ctx.scale(scriptScale, scriptScale);
                const rx = -(textWidth as number) / 2 / scriptScale; 
                const ry = fontSize / 2 / scriptScale + scriptOffsetY / scriptScale;
                ctx.fillText(text, rx, ry);
                ctx.restore();
              } else {
                const rx = -(textWidth as number) / 2; const ry = fontSize / 2; 
                ctx.fillText(text, rx, ry); 
              }
              ctx.restore();
            } else if (!wrap) {
              let toDraw = text;
              if (overflow === 'ellipsis' && (textWidth as number) > maxWidth) {
                let lo = 0, hi = text.length;
                while (lo < hi) { const mid = Math.floor((lo + hi) / 2); const s = text.slice(0, mid) + '…'; const w2 = this.textCache.get(font, s) ?? ctx.measureText(s).width; if (w2 <= maxWidth) { lo = mid + 1; this.textCache.set(font, s, w2); } else { hi = mid; } }
                toDraw = text.slice(0, Math.max(0, lo - 1)) + '…';
              }
              
              // Apply superscript/subscript if present
              if (hasScript) {
                ctx.save();
                ctx.scale(scriptScale, scriptScale);
                ctx.fillText(toDraw, tx / scriptScale, (ty + scriptOffsetY) / scriptScale, maxWidth / scriptScale);
                ctx.restore();
              } else {
                ctx.fillText(toDraw, tx, ty, maxWidth);
              }
              
              // Phase 1 UI: Strikethrough rendering (fast path: only if property is true)
              if (style?.strikethrough) {
                const metrics = ctx.measureText(toDraw);
                const strikeY = ty - (metrics.actualBoundingBoxAscent || fontSize * 0.8) * 0.3;
                const strikeWidth = (typeof metrics.width === 'number') ? metrics.width : (textWidth as number);
                ctx.strokeStyle = textColor;
                ctx.lineWidth = Math.max(1, fontSize * 0.08); // ~8% of font size
                ctx.beginPath();
                ctx.moveTo(tx, strikeY);
                ctx.lineTo(tx + strikeWidth * scriptScale, strikeY);
                ctx.stroke();
              }

              if (style?.underline || hyperlink?.target) {
                const metrics = ctx.measureText(toDraw);
                const underlineY = ty + (metrics.actualBoundingBoxDescent || fontSize * 0.15) + 1;
                const underlineWidth = (typeof metrics.width === 'number') ? metrics.width : (textWidth as number);
                ctx.strokeStyle = textColor;
                ctx.lineWidth = Math.max(1, fontSize * 0.05);
                ctx.beginPath();
                ctx.moveTo(tx, underlineY);
                ctx.lineTo(tx + underlineWidth * scriptScale, underlineY);
                ctx.stroke();
              }
            } else {
              const raw = String(text); const paragraphs = raw.split(/\n/); const words = paragraphs.flatMap((p, i) => (i > 0 ? ['\n', ...p.split(/\s+/)] : p.split(/\s+/)));
              const lines: string[] = []; let line = '';
              for (const wtoken of words) { if (wtoken === '\n') { if (line) { lines.push(line); line = ''; } continue; } const candidate = line ? line + ' ' + wtoken : wtoken; const cw = this.textCache.get(font, candidate) ?? ctx.measureText(candidate).width; if (cw <= maxWidth || !line) { line = candidate; this.textCache.set(font, candidate, cw); } else { lines.push(line); line = wtoken; } }
              if (line) lines.push(line);
              // Compute vertical offset using layout function (multi-line case)
              const totalH = lines.length * (fontSize + 2);
              let lineY = y + computeVerticalOffset(style?.valign, drawH, totalH, fontSize, 2, 4) - fontSize;
              const startX = align === 'center' ? (x + drawW / 2) : (align === 'right' ? (x + drawW - 4) : (x + 4 + indentOffset));
              
              // Phase 1 UI: Apply superscript/subscript to wrapped text (fast path: only if needed)
              if (hasScript) {
                ctx.save();
                ctx.scale(scriptScale, scriptScale);
              }
              
              for (const ln of lines) { 
                let lx = startX; 
                const lw = this.textCache.get(font, ln) ?? ctx.measureText(ln).width; 
                if (align === 'center') lx -= (lw as number) / 2; 
                else if (align === 'right') lx -= (lw as number); 
                
                const finalLx = hasScript ? lx / scriptScale : lx;
                const finalLineY = hasScript ? (lineY + scriptOffsetY) / scriptScale : lineY;
                const finalMaxWidth = hasScript ? maxWidth / scriptScale : maxWidth;
                
                ctx.fillText(ln, finalLx, finalLineY, finalMaxWidth); 
                
                // Phase 1 UI: Strikethrough for wrapped lines (fast path: only if property is true)
                if (style?.strikethrough) {
                  const metrics = ctx.measureText(ln);
                  const strikeY = hasScript 
                    ? (lineY + scriptOffsetY - (metrics.actualBoundingBoxAscent || fontSize * 0.8) * 0.3) / scriptScale
                    : lineY - (metrics.actualBoundingBoxAscent || fontSize * 0.8) * 0.3;
                  const strikeWidth = (typeof metrics.width === 'number') ? metrics.width : (lw as number);
                  ctx.strokeStyle = textColor;
                  ctx.lineWidth = Math.max(1, fontSize * 0.08) / (hasScript ? scriptScale : 1);
                  ctx.beginPath();
                  ctx.moveTo(finalLx, strikeY);
                  ctx.lineTo(finalLx + strikeWidth, strikeY);
                  ctx.stroke();
                }
                
                lineY += fontSize + 2; 
                if (lineY > y + drawH) break; 
              }
              
              if (hasScript) {
                ctx.restore();
              }
            }
            ctx.restore();
            
            // Plugin after-render hook
            for (const plugin of this.plugins) {
              if (plugin.afterCellRender) {
                plugin.afterCellRender(ctx, { x, y, w: drawW, h: drawH }, { addr, value: v, style });
              }
            }
          }

          // Comment indicator (small red triangle top-right) if comments exist
          const cellObj = (this.sheet as any).getCell?.(addr);
          if (isAnchor && cfResult?.icon) {
            this.renderConditionalIcon(ctx, x, y, merged ? spanW : cw, merged ? spanH : rh, cfResult.icon);
          }

          if (cellObj?.comments?.length) {
            const drawW = merged ? spanW : cw; const drawH = merged ? spanH : rh;
            ctx.save();
            ctx.beginPath();
            const triSize = Math.min(8 * this.zoom, Math.max(6, 6 * this.zoom));
            ctx.moveTo(x + drawW - triSize, y + 1);
            ctx.lineTo(x + drawW - 1, y + 1);
            ctx.lineTo(x + drawW - 1, y + triSize);
            ctx.closePath();
            ctx.fillStyle = '#d40000';
            ctx.fill();
            ctx.restore();
          }

          // Icon overlay rendering
          if (cellObj?.icon) {
            const icon = cellObj.icon; const drawW = merged ? spanW : cw; const drawH = merged ? spanH : rh;
            const size = (icon.size ?? 16) * this.zoom;
            const pad = 2 * this.zoom;
            let ix = x + pad, iy = y + pad;
            switch (icon.position) {
              case 'top-right': ix = x + drawW - size - pad; iy = y + pad; break;
              case 'bottom-left': ix = x + pad; iy = y + drawH - size - pad; break;
              case 'bottom-right': ix = x + drawW - size - pad; iy = y + drawH - size - pad; break;
              case 'center': ix = x + (drawW - size) / 2; iy = y + (drawH - size) / 2; break;
              case 'top-left': default: ix = x + pad; iy = y + pad; break;
            }
            ctx.save();
            if (icon.type === 'emoji') {
              ctx.font = `${size}px ${this.theme.fontFamily}`;
              ctx.fillText(icon.source, ix, iy + size - 4);
            } else if (icon.type === 'builtin') {
              // Simple builtin shapes: e.g., 'warning', 'info'
              ctx.beginPath();
              if (icon.source === 'warning') {
                ctx.fillStyle = '#ffcc00';
                ctx.moveTo(ix + size / 2, iy);
                ctx.lineTo(ix + size, iy + size);
                ctx.lineTo(ix, iy + size);
                ctx.closePath();
                ctx.fill();
              } else if (icon.source === 'info') {
                ctx.fillStyle = '#0078d4';
                ctx.arc(ix + size / 2, iy + size / 2, size / 2, 0, Math.PI * 2);
                ctx.fill();
                ctx.fillStyle = '#ffffff';
                ctx.font = `${Math.floor(size * 0.6)}px ${this.theme.fontFamily}`;
                ctx.fillText('i', ix + size * 0.35, iy + size * 0.7);
              }
            } else if (icon.type === 'url') {
              // Defer image loading; cache per URL
              const cacheKey = `img:${icon.source}`;
              let img = (this as any)._imageCache?.get(cacheKey);
              if (!img) {
                img = new Image();
                img.crossOrigin = 'anonymous';
                img.src = icon.source;
                (this as any)._imageCache = (this as any)._imageCache || new Map();
                (this as any)._imageCache.set(cacheKey, img);
                img.onload = () => { this.invalidateRect(x, y, drawW, drawH); };
              }
              if (img.complete && img.naturalWidth) {
                ctx.drawImage(img, ix, iy, size, size);
              }
            }
            ctx.restore();
          }

          if (cellObj?.customComponent?.type === 'icon') {
            this.renderCustomCellIcon(
              ctx,
              cellObj.customComponent,
              x,
              y,
              merged ? spanW : cw,
              merged ? spanH : rh,
            );
          }
          x += this.sheet.getColumnWidth(col) * this.zoom; col++;
        }
        y += this.sheet.getRowHeight(row) * this.zoom; rowIndex++;
      }

      this.drawLayers('cells', ctx, width, height);

      // Draw selections with light gray fill (Excel-style)
      const sels = this.selections.length ? this.selections : (this.selection ? [this.selection] : []);
      for (const sel of sels) {
        const a = sel.start; const b = sel.end; const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row); const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
        const rect = this.rectForRange(r1, c1, r2, c2);
        if (rect) {
          // Fill with very light gray (Excel uses rgba(217, 217, 217, 0.3) or similar)
          ctx.fillStyle = 'rgba(217, 217, 217, 0.3)';
          ctx.fillRect(rect.x + 1, rect.y + 1, rect.w - 2, rect.h - 2);
          // Draw border
          ctx.strokeStyle = selectionColor; 
          ctx.lineWidth = 2; 
          ctx.strokeRect(rect.x + 1, rect.y + 1, rect.w - 2, rect.h - 2);
        }
      }
      
      // Draw auto-fill handle on active selection (last one)
      if (sels.length > 0) {
        const activeSel = sels[sels.length - 1];
        const a = activeSel.start; const b = activeSel.end;
        const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row);
        const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
        const rect = this.rectForRange(r1, c1, r2, c2);
        if (rect) {
          const handleSize = 6;
          ctx.fillStyle = selectionColor;
          ctx.fillRect(rect.x + rect.w - handleSize / 2, rect.y + rect.h - handleSize / 2, handleSize, handleSize);
        }
      }

      this.drawLayers('selection', ctx, width, height);
      if (this.viewMode !== 'normal') {
        this.drawViewModeOverlay(ctx, width, height);
      }
      this.drawLayers('overlays', ctx, width, height);
      ctx.restore();
    }

    this.drawLayers('after', ctx, width, height);
    
    // After full redraw, update offscreen gridlines canvas if needed
    if (this.gridLinesNeedRedraw) {
      this.renderGridLines();
      this.gridLinesNeedRedraw = false;
    }
    
    // Clear dirty cells since we just did a full redraw
    this.dirtyCells.clear();
    
    this._lastRenderMs = performance.now() - t0;
    this.dirty = null;
    if (this.options.debug) {
      console.log('🎨 [CanvasRenderer] redraw() completed in', this._lastRenderMs.toFixed(2), 'ms');
    }
    // observability callback
    if (this.options.onRender) {
      try { this.options.onRender({ ms: this._lastRenderMs, regions: this.lastDirtyRects.slice() }); } catch {}
    }
  }

  private sheetXForCol(col: number): number {
    this.ensureLayoutCache();
    const clampedCol = Math.max(1, Math.min(col, this.sheet.colCount + 1));
    return this.options.headerWidth + (this.colLeftsCache[clampedCol - 1] ?? 0) - this.scrollX;
  }

  private sheetYForRow(row: number): number {
    this.ensureLayoutCache();
    const rowIndex = this.lowerBound(this.visibleRowsCache, row);
    return this.options.headerHeight + (this.rowTopsCache[rowIndex] ?? 0) - this.scrollY;
  }

  private getPageBreakPositions(): { rowBreaks: number[]; colBreaks: number[] } {
    const metrics = getDefaultPageMetrics(this.zoom);
    const rows = this.getVisibleRows();
    const rowBreaks = computeRowPageBreaks(
      rows,
      (row) => this.sheet.getRowHeight(row) * this.zoom,
      metrics.contentHeightPx,
    );
    const colBreaks = computeColPageBreaks(
      this.sheet.colCount,
      (col) => this.sheet.getColumnWidth(col) * this.zoom,
      metrics.contentWidthPx,
    );
    return { rowBreaks, colBreaks };
  }

  private drawViewModePageBackground(ctx: CanvasRenderingContext2D, width: number, height: number) {
    const { headerHeight, headerWidth } = this.options;
    const metrics = getDefaultPageMetrics(this.zoom);
    const { rowBreaks, colBreaks } = this.getPageBreakPositions();
    const rowStarts = [this.getVisibleRows()[0] ?? 1, ...rowBreaks.map((r) => r + 1)];
    const colStarts = [1, ...colBreaks.map((c) => c + 1)];

    ctx.save();
    ctx.fillStyle = '#FFFFFF';
    ctx.strokeStyle = '#999999';
    ctx.lineWidth = 1;

    for (let pi = 0; pi < rowStarts.length; pi++) {
      const startRow = rowStarts[pi];
      const endRow = pi < rowBreaks.length ? rowBreaks[pi] : (this.getVisibleRows().at(-1) ?? startRow);
      for (let pj = 0; pj < colStarts.length; pj++) {
        const startCol = colStarts[pj];
        const endCol = pj < colBreaks.length ? colBreaks[pj] : this.sheet.colCount;
        const rect = this.rectForRange(startRow, startCol, endRow, endCol);
        if (!rect) continue;
        const pageX = rect.x - metrics.marginLeftPx;
        const pageY = rect.y - metrics.marginTopPx - metrics.headerPx;
        const pageW = rect.w + metrics.marginLeftPx + metrics.marginRightPx;
        const pageH = rect.h + metrics.marginTopPx + metrics.marginBottomPx + metrics.headerPx + metrics.footerPx;
        if (pageX + pageW < headerWidth || pageY + pageH < headerHeight) continue;
        if (pageX > width || pageY > height) continue;
        ctx.fillRect(pageX, pageY, pageW, pageH);
        ctx.strokeRect(pageX + 0.5, pageY + 0.5, pageW - 1, pageH - 1);
      }
    }
    ctx.restore();
  }

  private drawViewModeOverlay(ctx: CanvasRenderingContext2D, width: number, height: number) {
    const { headerHeight, headerWidth } = this.options;
    const metrics = getDefaultPageMetrics(this.zoom);
    const { rowBreaks, colBreaks } = this.getPageBreakPositions();

    ctx.save();
    ctx.setLineDash([6, 4]);
    ctx.strokeStyle = '#0070C0';
    ctx.lineWidth = 1;

    for (const row of rowBreaks) {
      const y = this.sheetYForRow(row) + this.sheet.getRowHeight(row) * this.zoom;
      if (y < headerHeight || y > height) continue;
      ctx.beginPath();
      ctx.moveTo(headerWidth, y + 0.5);
      ctx.lineTo(width, y + 0.5);
      ctx.stroke();
    }

    for (const col of colBreaks) {
      const x = this.sheetXForCol(col) + this.sheet.getColumnWidth(col) * this.zoom;
      if (x < headerWidth || x > width) continue;
      ctx.beginPath();
      ctx.moveTo(x + 0.5, headerHeight);
      ctx.lineTo(x + 0.5, height);
      ctx.stroke();
    }

    if (this.viewMode === 'pageLayout') {
      ctx.setLineDash([]);
      ctx.fillStyle = 'rgba(220, 220, 220, 0.55)';
      ctx.font = `10px ${this.theme.fontFamily}`;
      ctx.fillStyle = '#666666';

      const rowStarts = [this.getVisibleRows()[0] ?? 1, ...rowBreaks.map((r) => r + 1)];
      const colStarts = [1, ...colBreaks.map((c) => c + 1)];

      for (let pi = 0; pi < rowStarts.length; pi++) {
        const startRow = rowStarts[pi];
        const endRow = pi < rowBreaks.length ? rowBreaks[pi] : (this.getVisibleRows().at(-1) ?? startRow);
        for (let pj = 0; pj < colStarts.length; pj++) {
          const startCol = colStarts[pj];
          const endCol = pj < colBreaks.length ? colBreaks[pj] : this.sheet.colCount;
          const rect = this.rectForRange(startRow, startCol, endRow, endCol);
          if (!rect) continue;

          const pageX = rect.x - metrics.marginLeftPx;
          const pageY = rect.y - metrics.marginTopPx - metrics.headerPx;
          const pageW = rect.w + metrics.marginLeftPx + metrics.marginRightPx;
          const pageH = rect.h + metrics.marginTopPx + metrics.marginBottomPx + metrics.headerPx + metrics.footerPx;

          ctx.fillStyle = 'rgba(220, 220, 220, 0.55)';
          ctx.fillRect(pageX, pageY, metrics.marginLeftPx, pageH);
          ctx.fillRect(pageX + pageW - metrics.marginRightPx, pageY, metrics.marginRightPx, pageH);
          ctx.fillRect(pageX, pageY, pageW, metrics.marginTopPx + metrics.headerPx);
          ctx.fillRect(pageX, pageY + pageH - metrics.marginBottomPx - metrics.footerPx, pageW, metrics.marginBottomPx + metrics.footerPx);

          const headerFooter = this.sheet.getHeaderFooter?.() ?? {
            header: { left: '', center: '', right: '' },
            footer: { left: '', center: '', right: '' },
          };
          const pageNumber = pi * colStarts.length + pj + 1;
          const totalPages = rowStarts.length * colStarts.length;
          const expandCtx = {
            page: pageNumber,
            pages: totalPages,
            sheetName: this.sheet.name,
            fileName: 'Workbook',
          };

          const headerBandY = pageY + metrics.marginTopPx;
          const headerBandH = metrics.headerPx;
          const footerBandY = pageY + pageH - metrics.marginBottomPx - metrics.footerPx;
          const footerBandH = metrics.footerPx;

          this.drawHeaderFooterBand(
            ctx,
            pageX,
            pageW,
            headerBandY,
            headerBandH,
            headerFooter.header,
            expandCtx,
            'Header',
          );
          this.drawHeaderFooterBand(
            ctx,
            pageX,
            pageW,
            footerBandY,
            footerBandH,
            headerFooter.footer,
            expandCtx,
            'Footer',
          );
        }
      }
    }

    ctx.restore();
  }

  private drawHeaderFooterBand(
    ctx: CanvasRenderingContext2D,
    pageX: number,
    pageW: number,
    bandY: number,
    bandH: number,
    section: HeaderFooterSection,
    expandCtx: { page: number; pages: number; sheetName: string; fileName: string },
    placeholder: 'Header' | 'Footer',
  ): void {
    ctx.save();
    ctx.fillStyle = '#333333';
    ctx.font = `10px ${this.theme.fontFamily}`;
    ctx.textBaseline = 'middle';

    const padding = 8 * this.zoom;
    const centerY = bandY + bandH / 2;

    if (!sectionHasContent(section)) {
      ctx.textAlign = 'center';
      ctx.fillStyle = '#666666';
      ctx.fillText(placeholder, pageX + pageW / 2, centerY);
      ctx.restore();
      return;
    }

    const left = expandHeaderFooterText(section.left, expandCtx);
    const center = expandHeaderFooterText(section.center, expandCtx);
    const right = expandHeaderFooterText(section.right, expandCtx);

    ctx.fillStyle = '#333333';
    if (left) {
      ctx.textAlign = 'left';
      ctx.fillText(left, pageX + padding, centerY);
    }
    if (center) {
      ctx.textAlign = 'center';
      ctx.fillText(center, pageX + pageW / 2, centerY);
    }
    if (right) {
      ctx.textAlign = 'right';
      ctx.fillText(right, pageX + pageW - padding, centerY);
    }
    ctx.restore();
  }

  private rectForRange(r1: number, c1: number, r2: number, c2: number): { x: number; y: number; w: number; h: number } | null {
    this.ensureLayoutCache();
    const { headerHeight, headerWidth } = this.options;
    const firstRowIndex = this.lowerBound(this.visibleRowsCache, r1);
    const lastRowIndexExclusive = this.upperBound(this.visibleRowsCache, r2);
    if (firstRowIndex >= lastRowIndexExclusive) return null;

    const leftCol = Math.max(1, Math.min(c1, this.sheet.colCount));
    const rightCol = Math.max(leftCol, Math.min(c2, this.sheet.colCount));
    const vx = headerWidth - this.scrollX + this.colLeftsCache[leftCol - 1];
    const vy = headerHeight - this.scrollY + this.rowTopsCache[firstRowIndex];
    const w = this.colLeftsCache[rightCol] - this.colLeftsCache[leftCol - 1];
    const h = this.rowTopsCache[lastRowIndexExclusive] - this.rowTopsCache[firstRowIndex];
    if (vx + w < headerWidth || vy + h < headerHeight) return null; // off-screen
    return { x: vx, y: vy, w, h };
  }

  // Excel-like column label (1 -> A, 26 -> Z, 27 -> AA)
  private colLabel(n: number): string {
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
  }

  // Map viewport coordinates (client space relative to container) to a cell address.
  cellAt(clientX: number, clientY: number): Address | null {
    const rect = this.canvas.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    const { headerHeight, headerWidth } = this.options;
    if (x < headerWidth || y < headerHeight) return null;

    const col = this.columnIndexAt(x - headerWidth + this.scrollX);
    const rowIndex = this.visibleRowIndexAt(y - headerHeight + this.scrollY);
    const row = rowIndex == null ? undefined : this.visibleRowsCache[rowIndex];
    if (col == null || row == null) return null;
    return { row, col };
  }

  // Hit testing for header/cell areas
  hitTest(clientX: number, clientY: number):
    | { type: 'col-resize'; col: number }
    | { type: 'row-resize'; row: number }
    | { type: 'cell'; addr: Address }
    | { type: 'header-col'; col: number }
    | { type: 'header-row'; row: number }
    | { type: 'fill-handle'; rangeIndex: number }
    | { type: 'select-all' }
    | null {
    const rect = this.canvas.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    const { headerHeight, headerWidth } = this.options;
    const threshold = 4; // px proximity to border
    
    // Check select-all button (top-left corner)
    if (x <= headerWidth && y <= headerHeight) {
      return { type: 'select-all' };
    }
    
    // Check auto-fill handle for active selection (last one in multi-range)
    if (this.selections.length > 0) {
      const activeIdx = this.selections.length - 1;
      const sel = this.selections[activeIdx];
      const r1 = Math.min(sel.start.row, sel.end.row);
      const r2 = Math.max(sel.start.row, sel.end.row);
      const c1 = Math.min(sel.start.col, sel.end.col);
      const c2 = Math.max(sel.start.col, sel.end.col);
      const selRect = this.rectForRange(r1, c1, r2, c2);
      if (selRect) {
        const handleSize = 6;
        const hx = selRect.x + selRect.w - handleSize / 2;
        const hy = selRect.y + selRect.h - handleSize / 2;
        if (x >= hx && x <= hx + handleSize && y >= hy && y <= hy + handleSize) {
          return { type: 'fill-handle', rangeIndex: activeIdx };
        }
      }
    }
    
    if (y <= headerHeight && x >= headerWidth) {
      // Column header area - check resize handles first
      const contentX = x - headerWidth + this.scrollX;
      const col = this.columnIndexAt(contentX);
      if (col == null) return null;
      const left = this.colLeftsCache[col - 1];
      const right = this.colLeftsCache[col];
      if (col > 1 && Math.abs(contentX - left) <= threshold) return { type: 'col-resize', col: col - 1 };
      if (Math.abs(contentX - right) <= threshold) return { type: 'col-resize', col };
      return { type: 'header-col', col };
    }
    if (x <= headerWidth && y >= headerHeight) {
      // Row header area - check resize handles first
      const contentY = y - headerHeight + this.scrollY;
      const rowIndex = this.visibleRowIndexAt(contentY);
      if (rowIndex == null) return null;
      const row = this.visibleRowsCache[rowIndex];
      const top = this.rowTopsCache[rowIndex];
      const bottom = this.rowTopsCache[rowIndex + 1];
      if (rowIndex > 0 && Math.abs(contentY - top) <= threshold) return { type: 'row-resize', row: this.visibleRowsCache[rowIndex - 1] };
      if (Math.abs(contentY - bottom) <= threshold) return { type: 'row-resize', row };
      return { type: 'header-row', row };
    }
    const addr = this.cellAt(clientX, clientY);
    return addr ? { type: 'cell', addr } : null;
  }

  private drawLayers(stage: RenderStage, ctx: CanvasRenderingContext2D, width: number, height: number) {
    if (!this.layers.length) return;
    const api = {
      width, height, dpr: this.dpr, options: this.options, sheet: this.sheet,
      getScroll: () => this.getScroll(), rectForRange: (r1: number, c1: number, r2: number, c2: number) => this.rectForRange(r1, c1, r2, c2)
    };
    for (const l of this.layers) if (l.stage === stage) l.draw(this.ctx, api);
  }

  // ==================== Event System ====================

  private isAddressInRange(addr: Address, range: { start: Address; end: Address }): boolean {
    const r1 = Math.min(range.start.row, range.end.row);
    const r2 = Math.max(range.start.row, range.end.row);
    const c1 = Math.min(range.start.col, range.end.col);
    const c2 = Math.max(range.start.col, range.end.col);
    return addr.row >= r1 && addr.row <= r2 && addr.col >= c1 && addr.col <= c2;
  }

  private isAddressInAnySelection(addr: Address): boolean {
    const activeSelections = this.selections.length ? this.selections : (this.selection ? [this.selection] : []);
    return activeSelections.some((range) => this.isAddressInRange(addr, range));
  }

  private getFullSheetRange(): { start: Address; end: Address } {
    return {
      start: { row: 1, col: 1 },
      end: { row: this.sheet.rowCount, col: this.sheet.colCount },
    };
  }

  private handleMouseDown = (e: MouseEvent) => {
    const hit = this.hitTest(e.clientX, e.clientY);
    if (!hit) return;
    
    // Handle select-all button (top-left corner)
    if (hit.type === 'select-all') {
      this.setSelections([this.getFullSheetRange()]);
      this.isDragging = false;
      this.dragStartCell = null;
      this.clickStartCell = null;
      return;
    }
    
    // Handle column resize
    if (hit.type === 'col-resize') {
      const applyToAll = this.shouldResizeAllFromSelection();
      this.resizeState = {
        type: 'col',
        index: hit.col,
        startPos: e.clientX,
        startSize: this.sheet.getColumnWidth(hit.col),
        applyToAll
      };
      return;
    }
    
    // Handle row resize
    if (hit.type === 'row-resize') {
      const applyToAll = this.shouldResizeAllFromSelection();
      this.resizeState = {
        type: 'row',
        index: hit.row,
        startPos: e.clientY,
        startSize: this.sheet.getRowHeight(hit.row),
        applyToAll
      };
      return;
    }
    
    // Handle header clicks for column/row selection
    if (hit.type === 'header-col') {
      const col = hit.col;
      this.setSelections([{ start: { row: 1, col }, end: { row: this.sheet.rowCount, col } }]);
      return;
    }
    if (hit.type === 'header-row') {
      const row = hit.row;
      this.setSelections([{ start: { row, col: 1 }, end: { row, col: this.sheet.colCount } }]);
      return;
    }
    
    if (hit.type !== 'cell') return;
    const addr = hit.addr;

    // Right-click / Ctrl+click (Mac): preserve multi-cell selection when clicking inside it
    const isContextMenuTrigger = e.button === 2 || (e.button === 0 && e.ctrlKey);
    if (isContextMenuTrigger) {
      if (!this.isAddressInAnySelection(addr)) {
        this.setSelections([{ start: addr, end: addr }]);
      }
      this.isDragging = false;
      this.dragStartCell = null;
      this.clickStartCell = null;
      return;
    }
    
    // Check for double-click
    const now = Date.now();
    const isDoubleClick = this.clickStartCell && 
      addr.row === this.clickStartCell.row && 
      addr.col === this.clickStartCell.col && 
      now - this.clickStartTime < 300;
    
    if (isDoubleClick) {
      e.preventDefault(); // Prevent text selection on double-click
      const bounds = this.getCellBounds(addr);
      if (bounds) {
        const eventObj = {
          type: 'cell-double-click' as const,
          event: { address: addr, bounds, originalEvent: e },
        };
        this.sheet['events'].emit(eventObj);
      }
      this.clickStartCell = null;
      this.isDragging = false;
      return;
    }
    
    // Start drag selection
    this.isDragging = true;
    this.dragStartCell = addr;
    this.clickStartCell = addr;
    this.clickStartTime = now;
    
    // Set initial single-cell selection
    this.setSelections([{ start: addr, end: addr }]);
    
    const bounds = this.getCellBounds(addr);
    if (bounds) {
      this.sheet['events'].emit({
        type: 'cell-click',
        event: { address: addr, bounds, originalEvent: e },
      });
    }
  };

  private handleMouseMove = (e: MouseEvent) => {
    // Store mouse position without processing - will be handled in rAF
    this.pendingMouseX = e.clientX;
    this.pendingMouseY = e.clientY;
    
    // Schedule processing on next animation frame if not already scheduled
    if (!this.hoverProcessPending) {
      this.hoverProcessPending = true;
      requestAnimationFrame(() => this.processHover());
    }
  };

  // Process hover logic on rAF to throttle to max 60fps and reduce RAM/CPU
  // Optimizations applied:
  // 1. rAF throttling: max 60 calls/sec vs 1000+ mousemove events
  // 2. Primitive comparisons: number === number (no object allocations)
  // 3. Early exit: skip event emission if cell unchanged (99% of hovers)
  // 4. Visibility check: pause when tab hidden (saves CPU in background)
  // 5. No event references: originalEvent set to null (prevents memory leaks)
  private processHover = () => {
    this.hoverProcessPending = false;
    
    // Skip processing if tab is hidden (micro-optimization)
    if (!this.isVisible) {
      return;
    }
    
    const mouseX = this.pendingMouseX;
    const mouseY = this.pendingMouseY;
    
    const addr = this.cellAt(mouseX, mouseY);
    
    // Handle column/row resize dragging
    if (this.resizeState) {
      if (this.resizeState.type === 'col') {
        const delta = (mouseX - this.resizeState.startPos) / this.zoom;
        const newWidth = Math.max(10, this.resizeState.startSize + delta) | 0; // Integer hint
        if (this.resizeState.applyToAll) {
          for (let c = 1; c <= this.sheet.colCount; c++) {
            this.sheet.setColumnWidth(c, newWidth);
          }
          // All column boundaries changed, force full redraw.
          this.dirtyCells.clear();
        } else {
          this.sheet.setColumnWidth(this.resizeState.index, newWidth);
          // Mark all visible cells in this column dirty
          const visRows = this.getVisibleRows();
          for (const row of visRows) {
            this.markCellDirty(row, this.resizeState.index);
          }
        }
        this.gridLinesNeedRedraw = true; // Gridlines changed
        this.scheduleRedraw();
      } else if (this.resizeState.type === 'row') {
        const delta = (mouseY - this.resizeState.startPos) / this.zoom;
        const newHeight = Math.max(10, this.resizeState.startSize + delta) | 0; // Integer hint
        if (this.resizeState.applyToAll) {
          for (let r = 1; r <= this.sheet.rowCount; r++) {
            this.sheet.setRowHeight(r, newHeight);
          }
          // All row boundaries changed, force full redraw.
          this.dirtyCells.clear();
        } else {
          this.sheet.setRowHeight(this.resizeState.index, newHeight);
          // Mark all visible cells in this row dirty
          for (let c = 1; c <= this.sheet.colCount; c++) {
            this.markCellDirty(this.resizeState.index, c);
          }
        }
        this.gridLinesNeedRedraw = true; // Gridlines changed
        this.scheduleRedraw();
      }
      return;
    }
    
    // Handle drag selection
    if (this.isDragging && this.dragStartCell && addr) {
      // Extend selection from drag start to current cell
      this.setSelections([{ start: this.dragStartCell, end: addr }]);
      
      // Auto-scroll when near edges
      const rect = this.canvas.getBoundingClientRect();
      const relX = mouseX - rect.left;
      const relY = mouseY - rect.top;
      const viewport = this.getViewportSize();
      const margin = 24;
      const step = 10;
      
      if (relX < this.options.headerWidth + margin) {
        this.scrollBy(-step, 0);
      } else if (relX > this.options.headerWidth + viewport.width - margin) {
        this.scrollBy(step, 0);
      }
      
      if (relY < this.options.headerHeight + margin) {
        this.scrollBy(0, -step);
      } else if (relY > this.options.headerHeight + viewport.height - margin) {
        this.scrollBy(0, step);
      }
      
      return;
    }
    
    // Update cursor based on hover position
    const hit = this.hitTest(mouseX, mouseY);
    if (hit) {
      if (hit.type === 'col-resize') {
        this.canvas.style.cursor = 'col-resize';
      } else if (hit.type === 'row-resize') {
        this.canvas.style.cursor = 'row-resize';
      } else if (hit.type === 'select-all') {
        this.canvas.style.cursor = 'pointer';
      } else if (hit.type === 'cell' && this.sheet.getHyperlink?.(hit.addr)?.target) {
        this.canvas.style.cursor = 'pointer';
      } else {
        this.canvas.style.cursor = 'default';
      }
    } else {
      this.canvas.style.cursor = 'default';
    }
    
    // Handle hover - only fire events if cell changed (use primitives to avoid allocations)
    const newRow = addr ? addr.row : -1;
    const newCol = addr ? addr.col : -1;
    
    // Early exit if same cell (no event emission)
    if (newRow === this.lastHoveredRow && newCol === this.lastHoveredCol) {
      return;
    }
    
    // Mark previous and new cells dirty for hover highlight redraw
    if (this.lastHoveredRow !== -1 && this.lastHoveredCol !== -1) {
      this.markCellDirty(this.lastHoveredRow, this.lastHoveredCol);
    }
    if (newRow !== -1 && newCol !== -1) {
      this.markCellDirty(newRow, newCol);
    }
    
    // Fire hover-end for previous cell
    if (this.lastHoveredRow !== -1 && this.hoveredCell) {
      this.sheet['events'].emit({
        type: 'cell-hover-end',
        address: this.hoveredCell,
      });
    }
    
    // Fire hover for new cell
    if (addr) {
      const bounds = this.getCellBounds(addr);
      if (bounds) {
        this.sheet['events'].emit({
          type: 'cell-hover',
          event: { address: addr, bounds, originalEvent: null }, // No event ref to avoid memory retention
        });
      }
      this.hoveredCell = addr;
    } else {
      this.hoveredCell = null;
    }
    
    this.lastHoveredRow = newRow;
    this.lastHoveredCol = newCol;
  };

  private handleMouseUp = (e: MouseEvent) => {
    this.isDragging = false;
    this.resizeState = null;
  };

  private isFullSheetSelection(range: { start: Address; end: Address }): boolean {
    const minRow = Math.min(range.start.row, range.end.row);
    const maxRow = Math.max(range.start.row, range.end.row);
    const minCol = Math.min(range.start.col, range.end.col);
    const maxCol = Math.max(range.start.col, range.end.col);

    const coversRowsOneBased = minRow <= 1 && maxRow >= this.sheet.rowCount;
    const coversColsOneBased = minCol <= 1 && maxCol >= this.sheet.colCount;

    // Some legacy callers still use 0-based addresses when setting selection.
    const coversRowsZeroBased = minRow <= 0 && maxRow >= this.sheet.rowCount - 1;
    const coversColsZeroBased = minCol <= 0 && maxCol >= this.sheet.colCount - 1;

    return (coversRowsOneBased && coversColsOneBased) || (coversRowsZeroBased && coversColsZeroBased);
  }

  private shouldResizeAllFromSelection(): boolean {
    const activeSelections = this.selections.length ? this.selections : (this.selection ? [this.selection] : []);
    return activeSelections.some((range) => this.isFullSheetSelection(range));
  }

  private handleDblClick = (e: MouseEvent) => {
    const hit = this.hitTest(e.clientX, e.clientY);
    if (!hit) return;
    
    // Double-click on column resize handle - auto-size column
    if (hit.type === 'col-resize') {
      this.autoSizeColumn(hit.col);
      return;
    }
    
    // Double-click on row resize handle - auto-size row
    if (hit.type === 'row-resize') {
      this.autoSizeRow(hit.row);
      return;
    }
  };
  
  // Auto-size column to fit content
  private autoSizeColumn(col: number) {
    this.ctx.save();
    this.ctx.font = `${this.theme.fontSize}px ${this.theme.fontFamily}`;
    
    let maxWidth = 50; // minimum width
    
    // Scan visible rows for max content width
    const visRows = this.getVisibleRows();
    const scanRows = visRows.slice(0, Math.min(100, visRows.length)); // limit scan
    
    for (const row of scanRows) {
      const addr = { row, col };
      const cell = this.sheet.getCell(addr);
      if (!cell) continue;
      
      const value = cell.value;
      if (value != null) {
        const style = this.sheet.getCellStyle(addr);
        const formatted = this.formatCache.formatValue(value, style?.numberFormat);
        const text = formatted.text;
        const width = this.ctx.measureText(text).width;
        maxWidth = Math.max(maxWidth, width + 12); // 12px padding (6px each side)
      }
    }
    
    this.sheet.setColumnWidth(col, Math.min(maxWidth, 400)); // cap at 400px
    this.ctx.restore();
    this.scheduleRedraw();
  }
  
  // Auto-size row to fit content - calculates height needed to show all content
  private autoSizeRow(row: number) {
    this.ctx.save();
    this.ctx.font = `${this.theme.fontSize}px ${this.theme.fontFamily}`;
    
    let maxHeight = 20; // minimum height (default)
    const paddingTop = 2; // px
    const paddingBottom = 4; // px
    const totalVerticalPadding = paddingTop + paddingBottom; // 6px total
    
    // Scan all columns for max content height in this row
    for (let col = 1; col <= Math.min(this.sheet.colCount, 100); col++) {
      const addr = { row, col };
      const cell = this.sheet.getCell(addr);
      if (!cell) continue;
      
      const value = cell.value;
      if (value == null) continue;
      
      const style = this.sheet.getCellStyle(addr);
      const formatted = this.formatCache.formatValue(value, style?.numberFormat);
      const text = formatted.text;
      
      // Get cell-specific font size
      const cellFontSize = style?.fontSize ?? this.theme.fontSize;
      const lineHeight = cellFontSize * 1.2; // Line height includes spacing between lines
      
      // Get column width to calculate how many lines the content needs
      const colWidth = this.sheet.getColumnWidth(col);
      const horizontalPadding = 8; // 4px each side
      const maxTextWidth = colWidth - horizontalPadding;
      
      let lineCount = 1;
      
      // Check for explicit newlines first
      if (text.includes('\n')) {
        // Text has newlines - each line might also wrap
        const lines = text.split('\n');
        lineCount = 0;
        
        // Set font for this specific cell
        this.ctx.font = `${cellFontSize}px ${style?.fontFamily ?? this.theme.fontFamily}`;
        
        for (const line of lines) {
          if (line.length === 0) {
            // Empty line still takes up space
            lineCount += 1;
          } else {
            const lineWidth = this.ctx.measureText(line).width;
            if (lineWidth > maxTextWidth) {
              // This line wraps - calculate how many wrapped lines
              lineCount += Math.ceil(lineWidth / maxTextWidth);
            } else {
              lineCount += 1;
            }
          }
        }
      } else {
        // No newlines - calculate if text would wrap based on width
        this.ctx.font = `${cellFontSize}px ${style?.fontFamily ?? this.theme.fontFamily}`;
        const textWidth = this.ctx.measureText(text).width;
        
        if (textWidth > maxTextWidth) {
          // Text is wider than cell - calculate wrapped lines
          lineCount = Math.ceil(textWidth / maxTextWidth);
        } else {
          // Text fits in one line
          lineCount = 1;
        }
      }
      
      // Calculate total height needed: content + padding above and below
      const contentHeight = lineCount * lineHeight;
      const cellHeight = contentHeight + totalVerticalPadding;
      
      maxHeight = Math.max(maxHeight, cellHeight);
    }
    
    this.sheet.setRowHeight(row, Math.min(Math.ceil(maxHeight), 300)); // cap at 300px
    this.ctx.restore();
    this.scheduleRedraw();
  }

  private handleClick = (e: MouseEvent) => {
    // Click is now mostly handled by mousedown/mouseup
    // Keep this for compatibility but most logic moved to handleMouseDown
  };

  private handleContextMenu = (e: MouseEvent) => {
    e.preventDefault();
    const hit = this.hitTest(e.clientX, e.clientY);
    if (!hit) return;
    // Header column filter menu
    if (hit.type === 'header-col' && this.options.onRequestColumnFilterMenu) {
      const col = hit.col;
      const rect = this.canvas.getBoundingClientRect();
      const anchor = { x: e.clientX - rect.left, y: this.options.headerHeight };
  // Use 'filtered' domain so this column's own filter does not hide its possible values, but other column filters still narrow.
  const values = this.getColumnDistinctValues(col, { domain: 'filtered' });
      const apply = (selected: string[]) => {
        // If all values selected, clear filter (no-op visually) so icon/state resets
        if (selected.length === values.length) {
          this.sheet.clearColumnFilter(col);
        } else {
          this.sheet.setColumnFilter(col, { type: 'in', value: selected });
        }
        this.redraw();
      };
      const clear = () => { this.sheet.clearColumnFilter(col); this.redraw(); };
      this.options.onRequestColumnFilterMenu({ col, anchor, values, apply, clear });
      return;
    }
    const addr = hit.type === 'cell' ? hit.addr : null;
    if (!addr) return;
    
    const bounds = this.getCellBounds(addr);
    if (!bounds) return;
    
    this.sheet['events'].emit({
      type: 'cell-right-click',
      event: { address: addr, bounds, originalEvent: e },
    });
  };

  private handleVisibilityChange = () => {
    this.isVisible = !document.hidden;
    // Cancel pending hover processing when tab becomes hidden
    if (!this.isVisible) {
      this.hoverProcessPending = false;
    }
  };

  private setupEventListeners() {
    this.canvas.addEventListener('mousedown', this.handleMouseDown);
    this.canvas.addEventListener('mousemove', this.handleMouseMove);
    this.canvas.addEventListener('mouseup', this.handleMouseUp);
    this.canvas.addEventListener('click', this.handleClick);
    this.canvas.addEventListener('dblclick', this.handleDblClick);
    this.canvas.addEventListener('contextmenu', this.handleContextMenu);
    // Global mouseup to handle drag ending outside canvas
    window.addEventListener('mouseup', this.handleMouseUp);
    // Pause processing when tab is hidden to save CPU/RAM
    document.addEventListener('visibilitychange', this.handleVisibilityChange);
  }

  private removeEventListeners() {
    this.canvas.removeEventListener('mousedown', this.handleMouseDown);
    this.canvas.removeEventListener('mousemove', this.handleMouseMove);
    this.canvas.removeEventListener('mouseup', this.handleMouseUp);
    this.canvas.removeEventListener('click', this.handleClick);
    this.canvas.removeEventListener('dblclick', this.handleDblClick);
    this.canvas.removeEventListener('contextmenu', this.handleContextMenu);
    window.removeEventListener('mouseup', this.handleMouseUp);
    document.removeEventListener('visibilitychange', this.handleVisibilityChange);
  }

  // ==================== Navigation API ====================

  /**
   * Viewport rectangle for a cell range (null when fully off-screen).
   */
  getRangeRect(r1: number, c1: number, r2: number, c2: number): { x: number; y: number; w: number; h: number } | null {
    return this.rectForRange(r1, c1, r2, c2);
  }

  /**
   * Get cell bounds in viewport coordinates (includes zoom).
   */
  getCellBounds(addr: Address): { x: number; y: number; width: number; height: number } | null {
    const rect = this.rectForRange(addr.row, addr.col, addr.row, addr.col);
    if (!rect) return null;
    return { x: rect.x, y: rect.y, width: rect.w, height: rect.h };
  }

  /**
   * Scroll to make a cell visible
   */
  scrollToCell(addr: Address, align: 'start' | 'center' | 'end' | 'nearest' = 'nearest'): void {
    this.ensureLayoutCache();
    // Calculate cell position in content space
    const col = Math.max(1, Math.min(addr.col, this.sheet.colCount));
    const rowIndex = this.lowerBound(this.visibleRowsCache, addr.row);
    if (rowIndex >= this.visibleRowsCache.length || this.visibleRowsCache[rowIndex] !== addr.row) return;

    const cellX = this.colLeftsCache[col - 1];
    const cellY = this.rowTopsCache[rowIndex];
    const cellWidth = this.colLeftsCache[col] - this.colLeftsCache[col - 1];
    const cellHeight = this.rowTopsCache[rowIndex + 1] - this.rowTopsCache[rowIndex];
    const viewport = this.getViewportSize();
    
    let newScrollX = this.scrollX;
    let newScrollY = this.scrollY;
    
    // Horizontal scrolling
    if (align === 'start') {
      newScrollX = cellX;
    } else if (align === 'center') {
      newScrollX = cellX - viewport.width / 2 + cellWidth / 2;
    } else if (align === 'end') {
      newScrollX = cellX - viewport.width + cellWidth;
    } else {
      // 'nearest' - only scroll if not visible
      if (cellX < this.scrollX) {
        newScrollX = cellX;
      } else if (cellX + cellWidth > this.scrollX + viewport.width) {
        newScrollX = cellX + cellWidth - viewport.width;
      }
    }
    
    // Vertical scrolling
    if (align === 'start') {
      newScrollY = cellY;
    } else if (align === 'center') {
      newScrollY = cellY - viewport.height / 2 + cellHeight / 2;
    } else if (align === 'end') {
      newScrollY = cellY - viewport.height + cellHeight;
    } else {
      // 'nearest' - only scroll if not visible
      if (cellY < this.scrollY) {
        newScrollY = cellY;
      } else if (cellY + cellHeight > this.scrollY + viewport.height) {
        newScrollY = cellY + cellHeight - viewport.height;
      }
    }
    
    this.setScroll(newScrollX, newScrollY);
  }

  /**
   * Get currently visible cell range
   */
  getVisibleRange(): { start: Address; end: Address } {
    const { firstRow, firstCol } = this.visibleRange();
    const viewport = this.getViewportSize();
    this.ensureLayoutCache();
    const lastRowIndex = this.visibleRowIndexAt(this.scrollY + viewport.height) ?? Math.max(0, this.visibleRowsCache.length - 1);
    const lastRow = this.visibleRowsCache[Math.min(lastRowIndex, this.visibleRowsCache.length - 1)] ?? firstRow;
    const lastCol = this.columnIndexAt(this.scrollX + viewport.width) ?? this.sheet.colCount;
    
    return {
      start: { row: firstRow, col: firstCol },
      end: { row: lastRow, col: lastCol },
    };
  }
}

function stageOrder(s: RenderStage): number {
  const order: RenderStage[] = ['background', 'grid', 'headers', 'cells', 'selection', 'overlays', 'after'];
  return order.indexOf(s);
}

function unionRect(a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }) {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  const r = Math.max(a.x + a.w, b.x + b.w);
  const btm = Math.max(a.y + a.h, b.y + b.h);
  return { x, y, w: r - x, h: btm - y };
}
