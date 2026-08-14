import React, { forwardRef, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { Workbook, Worksheet, autoFill, FormulaWorkerProxy, type Address, type Command, type CommandManager, type CustomCellComponent, type EventAPI, SetCellComponentCommand } from '@cyber-sheet/core';
import { ComponentRegistry, CustomCellOverlay, type CustomCellComponentRenderProps, type RegisteredCellComponent } from './customComponents';
// Import locally to ensure dev picks up latest CanvasRenderer implementation
import { CanvasRenderer, CanvasRendererOptions, type ViewMode } from '../../renderer-canvas/src';
import { useCyberSheetConfig } from './config/CyberSheetConfigContext';
import { isArrowKey } from './utils/keyboardLayout';
import { registerRendererLazyBridges, reactivateRendererLazyBridges } from './rendererEventBridge';

export type CyberSheetProps = {
  workbook: Workbook;
  sheetName?: string;
  commandManager?: CommandManager;
  rendererOptions?: CanvasRendererOptions;
  style?: any;
  physicsOptions?: PhysicsOptions;
  zoom?: number; // external zoom control (1 = 100%)
  viewMode?: ViewMode;
  fontFamily?: string; // global font family configuration
  fontSize?: number; // global font size configuration (in pixels)
  onRendererReady?: (renderer: CanvasRenderer) => void;
  onSelectionChange?: (selection: { start: { row: number; col: number }; end: { row: number; col: number } }) => void;
  enableCustomComponents?: boolean;
};

export interface CyberSheetHandle extends EventAPI {
  getWorkbook: () => Workbook;
  getActiveSheet: () => Worksheet | undefined;
  executeCommand: (command: Command) => void;
  registerComponent: (id: string, component: RegisteredCellComponent) => void;
  unregisterComponent: (id: string) => void;
  listRegisteredComponents: () => string[];
  setCellComponent: (address: Address, component: CustomCellComponent | null) => void;
  getCellComponent: (address: Address) => CustomCellComponent | undefined;
  clearCellComponent: (address: Address) => void;
}

export type { CustomCellComponentRenderProps, RegisteredCellComponent };

export type PhysicsOptions = {
  // Toggles
  kineticScroll?: boolean;           // wheel/touch momentum
  selectionInertia?: boolean;        // selection glide after mouseup/touchend
  snapToGrid?: boolean;              // snap scroll to nearest cell boundaries when idle
  allowTouch?: boolean;              // enable touch handlers
  touchPanWithTwoFingers?: boolean;  // use two fingers to pan/scroll; single finger selects
  // Tuning
  friction?: number;                 // 0..1, velocity decay per frame (default 0.92)
  minSpeed?: number;                 // stop threshold (default 0.05)
  resistanceScale?: number;          // bigger = softer edges (default 40)
  edgeMargin?: number;               // px near edge to trigger autoscroll (default 24)
  autoScrollStep?: number;           // px step while autoscrolling (default 10)
  snapIntervalMs?: number;           // how often to check snap (default 200)
};

export const CyberSheet = forwardRef<CyberSheetHandle, CyberSheetProps>(function CyberSheet({
  workbook,
  sheetName,
  commandManager,
  rendererOptions,
  style,
  physicsOptions,
  zoom,
  viewMode = 'normal',
  fontFamily,
  fontSize,
  onRendererReady,
  onSelectionChange,
  enableCustomComponents = true,
}, ref) {
  const cyberSheetConfig = useCyberSheetConfig();
  const effectiveFontFamily = fontFamily ?? cyberSheetConfig.fonts.defaultFamily;
  const effectiveFontSize = fontSize ?? cyberSheetConfig.fonts.defaultSize;
  const containerRef = useRef(null as any);
  const rendererRef = useRef(null as any);
  const sheetRef = useRef(undefined as any);
  const overlayHostRef = useRef<HTMLElement | null>(null);
  const [overlayHost, setOverlayHost] = useState<HTMLElement | null>(null);
  const onSelectionChangeRef = useRef(onSelectionChange);
  onSelectionChangeRef.current = onSelectionChange;
  const [overlayTick, setOverlayTick] = useState(0);
  const commandManagerRef = useRef(commandManager);
  commandManagerRef.current = commandManager;
  const componentRegistryRef = useRef(new ComponentRegistry());
  const [selections, setSelections] = useState<Array<{ start: { row: number; col: number }; end: { row: number; col: number } }>>([]);
  const [horizontalScroll, setHorizontalScroll] = useState({ x: 0, maxX: 0, contentWidth: 0, viewportWidth: 0 });
  const selectionsRef = useRef(selections as any);
  const horizontalScrollbarRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => { selectionsRef.current = selections; }, [selections]);

  const syncHorizontalScrollState = useCallback((renderer = rendererRef.current as CanvasRenderer | null) => {
    if (!renderer) return;
    const scroll = renderer.getScroll();
    const max = renderer.getMaxScroll();
    const content = renderer.getContentSize();
    const viewport = renderer.getViewportSize();
    const next = {
      x: scroll.x,
      maxX: max.x,
      contentWidth: content.width,
      viewportWidth: viewport.width,
    };
    setHorizontalScroll((prev) =>
      prev.x === next.x &&
      prev.maxX === next.maxX &&
      prev.contentWidth === next.contentWidth &&
      prev.viewportWidth === next.viewportWidth
        ? prev
        : next
    );
  }, []);

  useEffect(() => {
    registerRendererLazyBridges(workbook.eventBus, () => rendererRef.current);
  }, [workbook]);

  useImperativeHandle(ref, () => ({
    on: workbook.eventBus.on.bind(workbook.eventBus),
    once: workbook.eventBus.once.bind(workbook.eventBus),
    off: workbook.eventBus.off.bind(workbook.eventBus),
    getWorkbook: () => workbook,
    getActiveSheet: () => sheetRef.current ?? workbook.activeSheet,
    executeCommand: (command: Command) => {
      if (!commandManagerRef.current) {
        throw new Error('CyberSheet executeCommand requires a commandManager prop');
      }
      commandManagerRef.current.execute(command);
    },
    registerComponent: (id: string, component: RegisteredCellComponent) => {
      componentRegistryRef.current.register(id, component);
    },
    unregisterComponent: (id: string) => {
      componentRegistryRef.current.unregister(id);
    },
    listRegisteredComponents: () => componentRegistryRef.current.list(),
    setCellComponent: (address: Address, component: CustomCellComponent | null) => {
      const sheet = sheetRef.current ?? workbook.activeSheet;
      if (!sheet) return;
      const apply = () => sheet.setCellComponent(address, component ?? undefined);
      if (commandManagerRef.current) {
        commandManagerRef.current.execute(new SetCellComponentCommand(sheet, address, component ?? undefined));
      } else {
        apply();
      }
      rendererRef.current?.invalidateRange(address.row, address.col, address.row, address.col);
      rendererRef.current?.scheduleRedraw?.();
    },
    getCellComponent: (address: Address) => {
      const sheet = sheetRef.current ?? workbook.activeSheet;
      return sheet?.getCellComponent(address);
    },
    clearCellComponent: (address: Address) => {
      const sheet = sheetRef.current ?? workbook.activeSheet;
      if (!sheet) return;
      if (commandManagerRef.current) {
        commandManagerRef.current.execute(new SetCellComponentCommand(sheet, address, undefined));
      } else {
        sheet.clearCellComponent(address);
      }
      rendererRef.current?.invalidateRange(address.row, address.col, address.row, address.col);
      rendererRef.current?.scheduleRedraw?.();
    },
  }), [workbook]);

  useEffect(() => {
    if (!onSelectionChangeRef.current || selections.length === 0) return;
    const active = selections[selections.length - 1];
    onSelectionChangeRef.current(active);
  }, [selections]);

  const setContainerRef = useCallback((node: HTMLDivElement | null) => {
    containerRef.current = node;
    if (overlayHostRef.current === node) return;
    overlayHostRef.current = node;
    setOverlayHost(node);
  }, []);

  // Physics state
  const velRef = useRef({ vx: 0, vy: 0 } as any);
  const lastWheelTimeRef = useRef(0 as any);
  const rafRef = useRef(null as any);
  const inertialSelectingRef = useRef(null as any);
  const draggingRef = useRef(null as any);
  const lastMouseRef = useRef({ x: 0, y: 0, t: 0 } as any);
  const touchRef = useRef(null as any);

  // Defaults for physics
  const p = {
    kineticScroll: false, // Disabled by default - let renderer handle scrolling
    selectionInertia: false, // Disabled by default
    snapToGrid: false, // Disabled by default
    allowTouch: true,
    touchPanWithTwoFingers: true,
    friction: 0.92,
    minSpeed: 0.05,
    resistanceScale: 40,
    edgeMargin: 24,
    autoScrollStep: 10,
    snapIntervalMs: 200,
    ...(physicsOptions || {}),
  } as PhysicsOptions & Required<Pick<PhysicsOptions, 'friction'|'minSpeed'|'resistanceScale'|'edgeMargin'|'autoScrollStep'|'snapIntervalMs'>>;

  useEffect(() => {
    sheetRef.current = sheetName ? workbook.getSheet(sheetName) : workbook.activeSheet;
    if (!sheetRef.current) return;
    const el = containerRef.current!;
    const effectiveRendererOptions: CanvasRendererOptions = {
      ...(rendererOptions ?? {}),
      formulaWorkerFactory: rendererOptions?.formulaWorker || rendererOptions?.formulaWorkerFactory
        ? rendererOptions.formulaWorkerFactory
        : () => new FormulaWorkerProxy(
            new Worker(new URL('../../core/src/worker/formula-worker.ts', import.meta.url), { type: 'module' })
          ),
    };
    const r = new CanvasRenderer(el, sheetRef.current, effectiveRendererOptions);
    rendererRef.current = r;
    syncHorizontalScrollState(r);
    setOverlayTick((value) => value + 1);
    if (typeof zoom === 'number' && typeof (r as any).setZoom === 'function') {
      try { (r as any).setZoom(zoom); } catch {}
    }
    if (typeof (r as any).setViewMode === 'function') {
      try { (r as any).setViewMode(viewMode); } catch {}
    }
    try { onRendererReady?.(r); } catch {}
    reactivateRendererLazyBridges(workbook.eventBus);
    return () => { r.dispose(); rendererRef.current = null; };
  }, [workbook, sheetName, syncHorizontalScrollState]);

  // Subscribe to renderer selection changes
  useEffect(() => {
    const r = rendererRef.current as CanvasRenderer | null;
    if (!r || typeof r.onSelectionChange !== 'function') return;
    
    const unsubscribe = r.onSelectionChange((nextSelections) => {
      setSelections(nextSelections);
    });
    
    return () => unsubscribe.dispose();
  }, [rendererRef.current]);

  useEffect(() => {
    const r = rendererRef.current as CanvasRenderer | null;
    if (!r || typeof r.onScrollChange !== 'function') return;

    syncHorizontalScrollState(r);
    const unsubscribe = r.onScrollChange((event) => {
      const content = r.getContentSize();
      const viewport = r.getViewportSize();
      const next = {
        x: event.x,
        maxX: event.maxX,
        contentWidth: content.width,
        viewportWidth: viewport.width,
      };
      setHorizontalScroll((prev) =>
        prev.x === next.x &&
        prev.maxX === next.maxX &&
        prev.contentWidth === next.contentWidth &&
        prev.viewportWidth === next.viewportWidth
          ? prev
          : next
      );
    });

    return () => unsubscribe.dispose();
  }, [rendererRef.current, syncHorizontalScrollState]);

  useEffect(() => {
    const scroller = horizontalScrollbarRef.current;
    if (!scroller) return;
    if (Math.abs(scroller.scrollLeft - horizontalScroll.x) > 1) {
      scroller.scrollLeft = horizontalScroll.x;
    }
  }, [horizontalScroll.x]);

  useEffect(() => {
    const el = containerRef.current as HTMLElement | null;
    if (!el || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(() => syncHorizontalScrollState());
    observer.observe(el);
    return () => observer.disconnect();
  }, [syncHorizontalScrollState]);

  // React to zoom prop changes
  useEffect(() => {
    const r = rendererRef.current as CanvasRenderer | null;
    if (!r) return;
    if (typeof zoom === 'number' && typeof (r as any).setZoom === 'function') {
      (r as any).setZoom(zoom);
      syncHorizontalScrollState(r);
    }
  }, [zoom, rendererRef.current, syncHorizontalScrollState]);

  useEffect(() => {
    const r = rendererRef.current as CanvasRenderer | null;
    if (!r || typeof (r as any).setViewMode !== 'function') return;
    (r as any).setViewMode(viewMode);
    syncHorizontalScrollState(r);
  }, [viewMode, rendererRef.current, syncHorizontalScrollState]);

  // React to font configuration changes
  useEffect(() => {
    const r = rendererRef.current as CanvasRenderer | null;
    if (!r || typeof r.setTheme !== 'function') return;
    
    const themeUpdate: any = {};
    if (effectiveFontFamily !== undefined) themeUpdate.fontFamily = effectiveFontFamily;
    if (effectiveFontSize !== undefined) themeUpdate.fontSize = effectiveFontSize;
    
    if (Object.keys(themeUpdate).length > 0) {
      r.setTheme(themeUpdate);
    }
  }, [effectiveFontFamily, effectiveFontSize, rendererRef.current]);

  // Basic wheel scroll handler
  useEffect(() => {
    const el = containerRef.current;
    const r = rendererRef.current;
    if (!el || !r) return;
    
    const onWheel = (e: WheelEvent) => {
      e.preventDefault();
      
      if (p.kineticScroll) {
        // Advanced: Kinetic scrolling with momentum (opt-in feature)
        const now = performance.now();
        const dt = Math.max(1, now - (lastWheelTimeRef.current || now));
        lastWheelTimeRef.current = now;
        velRef.current.vx += e.deltaX;
        velRef.current.vy += e.deltaY;
        stepScroll(r, velRef.current.vx, velRef.current.vy, true);
        if (rafRef.current == null) rafRef.current = requestAnimationFrame(() => animate(r));
      } else {
        // Simple: Direct scroll - let renderer handle all logic
        r.scrollBy(e.deltaX, e.deltaY);
      }
    };
    
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [rendererRef.current]);

  // Animation loop for kinetic scrolling and selection inertia
  const animate = (r: any) => {
    rafRef.current = null;
    const friction = p.friction || 0.92; // velocity decay per frame
    const minSpeed = p.minSpeed || 0.05; // stop threshold
    let { vx, vy } = velRef.current;
    let active = false;
    if (p.kineticScroll) {
      if (Math.abs(vx) > minSpeed || Math.abs(vy) > minSpeed) {
        stepScroll(r, vx, vy, false);
        vx *= friction; vy *= friction;
        velRef.current.vx = vx; velRef.current.vy = vy;
        active = true;
      }
    } else {
      velRef.current.vx = 0; velRef.current.vy = 0;
    }
    // Selection inertia updates
    if (p.selectionInertia && inertialSelectingRef.current) {
      const s = inertialSelectingRef.current;
      const currentSelections = selectionsRef.current as any[];
      const sel = currentSelections[s.index];
      if (sel) {
        const moved = stepSelection(r, sel, s, setSelections);
        if (moved) active = true; else inertialSelectingRef.current = null;
      } else {
        inertialSelectingRef.current = null;
      }
    }
    if (active) rafRef.current = requestAnimationFrame(() => animate(r));
  };

  // Apply scroll delta (only used when kineticScroll is enabled)
  const stepScroll = (r: any, dx: number, dy: number, immediate: boolean) => {
    const { x: sx, y: sy } = r.getScroll();
    // Simple: Just apply delta, renderer will handle clamping
    r.setScroll(sx + dx, sy + dy);
  };

  // Update selection position with inertia; also autoscroll when near edges
  const stepSelection = (r: any, sel: { start: any; end: any }, s: { index: number; dirX: number; dirY: number; vx: number; vy: number }, update: any) => {
    const sheet: Worksheet = sheetRef.current;
    const vp = r.getViewportSize?.();
    if (!vp) return false;
    // Decay velocity
    s.vx *= 0.9; s.vy *= 0.9;
    const speed = Math.hypot(s.vx, s.vy);
    if (speed < 0.1) return false;
    // Convert velocity to cell step based on approximate cell sizes
    const avgColW = Math.max(20, Math.round(vp.width / 10));
    const avgRowH = Math.max(12, Math.round(vp.height / 20));
    const dc = Math.round(s.vx / avgColW);
    const dr = Math.round(s.vy / avgRowH);
    if (dc === 0 && dr === 0) return true;
    const next = selections.slice();
    const a = sel.start;
    const b = sel.end;
    const newEnd = {
      row: Math.min(Math.max(1, b.row + dr), sheet.rowCount),
      col: Math.min(Math.max(1, b.col + dc), sheet.colCount),
    };
    next[s.index] = { start: a, end: newEnd };
    r.setSelections(next); update(next);
    // Edge autoscroll when selection moves to viewport edge
    const rect = r.rectForRange?.(newEnd.row, newEnd.col, newEnd.row, newEnd.col);
    if (rect) {
      const margin = 24;
      if (rect.x < r.options?.headerWidth + margin) r.scrollBy(-avgColW / 2, 0);
      if (rect.y < r.options?.headerHeight + margin) r.scrollBy(0, -avgRowH / 2);
      const vpW = vp.width, vpH = vp.height;
      if (rect.x + rect.w > r.options?.headerWidth + vpW - margin) r.scrollBy(avgColW / 2, 0);
      if (rect.y + rect.h > r.options?.headerHeight + vpH - margin) r.scrollBy(0, avgRowH / 2);
    }
    return true;
  };

  // Selection (single + multi-range) & resizing interactions
  // DISABLED: CanvasRenderer handles mouse events directly
  useEffect(() => {
    const el = containerRef.current as any;
    const r = rendererRef.current as any;
    const sheet = sheetRef.current as any;
  if (!el || !r || !sheet) return;
    
    return; // Skip all mouse handling - let CanvasRenderer do it

    draggingRef.current = null;

    const onMouseDown = (e: MouseEvent) => {
      const hit = r.hitTest(e.clientX, e.clientY);
      if (!hit) return;
      if (hit.type === 'fill-handle') {
        // Start auto-fill drag
        const idx = (hit as any).rangeIndex;
        const sel = selections[idx];
        if (!sel) return; // Safety check for undefined selection
        const r1 = Math.min(sel.start.row, sel.end.row);
        const r2 = Math.max(sel.start.row, sel.end.row);
        const c1 = Math.min(sel.start.col, sel.end.col);
        const c2 = Math.max(sel.start.col, sel.end.col);
        draggingRef.current = { type: 'auto-fill', rangeIndex: idx, sourceRange: { r1, c1, r2, c2 }, fillStart: { row: r2, col: c2 } } as any;
        e.preventDefault();
      } else if (hit.type === 'col-resize') {
        draggingRef.current = { type: 'col-resize', col: (hit as any).col, startX: e.clientX, startW: sheet.getColumnWidth((hit as any).col) } as any;
        e.preventDefault();
      } else if (hit.type === 'row-resize') {
        draggingRef.current = { type: 'row-resize', row: (hit as any).row, startY: e.clientY, startH: sheet.getRowHeight((hit as any).row) } as any;
        e.preventDefault();
      } else if (hit.type === 'cell') {
        const additive = e.ctrlKey || e.metaKey;
        const addr = (hit as any).addr;
        if (additive) {
          const next = selections.slice();
          next.push({ start: addr, end: addr });
          draggingRef.current = { type: 'select', start: addr, index: next.length - 1 } as any;
          r.setSelections(next);
          setSelections(next);
        } else {
          const sel = { start: addr, end: addr };
          draggingRef.current = { type: 'select', start: addr, index: 0 } as any;
          r.setSelections([sel]);
          setSelections([sel]);
        }
      }
    };

    const onMouseMove = (e: MouseEvent) => {
      const dragging = draggingRef.current as any;
      if (!dragging) return;
      if (dragging.type === 'auto-fill') {
        // Update fill preview range
        const addr = r.cellAt(e.clientX, e.clientY);
        if (addr) {
          const src = dragging.sourceRange;
          let fillR2 = addr.row, fillC2 = addr.col;
          // Clamp to sheet bounds
          fillR2 = Math.min(Math.max(src.r2, fillR2), sheet.rowCount);
          fillC2 = Math.min(Math.max(src.c2, fillC2), sheet.colCount);
          dragging.fillEnd = { row: fillR2, col: fillC2 };
          // Show preview by extending active selection visually
          const preview = { start: { row: src.r1, col: src.c1 }, end: { row: fillR2, col: fillC2 } };
          const next = selections.slice();
          next[dragging.rangeIndex] = preview;
          r.setSelections(next);
          setSelections(next);
        }
      } else if (dragging.type === 'col-resize') {
        const dx = e.clientX - dragging.startX;
        const newW = Math.max(20, dragging.startW + dx);
        sheet.setColumnWidth(dragging.col, newW);
        r.redraw?.();
      } else if (dragging.type === 'row-resize') {
        const dy = e.clientY - dragging.startY;
        const newH = Math.max(12, dragging.startH + dy);
        sheet.setRowHeight(dragging.row, newH);
        r.redraw?.();
      } else if (dragging.type === 'select') {
        // track velocity of mouse
        const now = performance.now();
        if (lastMouseRef.current.t) {
          const dt = Math.max(1, now - lastMouseRef.current.t);
          velRef.current.vx = (e.clientX - lastMouseRef.current.x) / dt * 16; // normalize per ~frame
          velRef.current.vy = (e.clientY - lastMouseRef.current.y) / dt * 16;
        }
        lastMouseRef.current = { x: e.clientX, y: e.clientY, t: now };
        const addr = r.cellAt(e.clientX, e.clientY);
        if (addr) {
          const next = selections.slice();
          const idx = dragging.index;
          next[idx] = { start: dragging.start, end: addr };
          r.setSelections(next);
          setSelections(next);
          // If near edges, autoscroll and prepare inertia vector based on mouse movement
          const rect = r.rectForRange?.(addr.row, addr.col, addr.row, addr.col);
          if (rect) {
            const margin = p.edgeMargin || 24;
            const step = p.autoScrollStep || 10;
            const opts = r.optionsReadonly ? r.optionsReadonly : { headerWidth: 48, headerHeight: 24 };
            if (rect.x < opts.headerWidth + margin) r.scrollBy(-step, 0);
            if (rect.y < opts.headerHeight + margin) r.scrollBy(0, -step);
            const vp = r.getViewportSize?.();
            if (vp) {
              if (rect.x + rect.w > opts.headerWidth + vp.width - margin) r.scrollBy(step, 0);
              if (rect.y + rect.h > opts.headerHeight + vp.height - margin) r.scrollBy(0, step);
            }
          }
        }
      }
    };

    const onMouseUp = () => {
      const dragging = draggingRef.current as any;
      if (dragging && dragging.type === 'auto-fill') {
        // Apply auto-fill to sheet
        if (dragging.fillEnd) {
          const src = dragging.sourceRange;
          const targetR2 = dragging.fillEnd.row;
          const targetC2 = dragging.fillEnd.col;
          // Only fill if we extended beyond source
          if (targetR2 > src.r2 || targetC2 > src.c2) {
            const targetRange = { r1: src.r2 + 1, c1: src.c1, r2: targetR2, c2: targetC2 };
            const fillResult = autoFill(src, targetRange, (addr) => sheet.getCellValue(addr));
            fillResult.forEach((value, key) => {
              const [row, col] = key.split(':').map(Number);
              sheet.setCellValue({ row, col }, value);
            });
            // Reset selection to filled range
            const filled = { start: { row: src.r1, col: src.c1 }, end: { row: targetR2, col: targetC2 } };
            r.setSelections([filled]);
            setSelections([filled]);
          }
        }
      } else if (dragging && dragging.type === 'select') {
        if (p.selectionInertia) {
          // Fire a brief selection inertia towards last move direction
          const idx = dragging.index;
          const vx = Math.max(-60, Math.min(60, velRef.current.vx * 8));
          const vy = Math.max(-60, Math.min(60, velRef.current.vy * 8));
          inertialSelectingRef.current = { index: idx, dirX: 1, dirY: 1, vx, vy };
          if (rafRef.current == null) rafRef.current = requestAnimationFrame(() => animate(r));
        }
      }
      draggingRef.current = null;
    };

    el.addEventListener('mousedown', onMouseDown);
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
    return () => {
      el.removeEventListener('mousedown', onMouseDown);
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };
  }, [rendererRef.current, sheetRef.current]);

  // Arrow-key navigation only — clipboard shortcuts are handled by ExcelApp via PasteCommand
  // so undo/redo stays on the command stack.
  useEffect(() => {
    const el = containerRef.current as any;
    const r = rendererRef.current as any;
    const sheet = sheetRef.current as Worksheet | undefined;
    if (!el || !r || !sheet) return;
    const onKeyDown = (e: KeyboardEvent) => {
      const ranges = selections;
      if (!ranges.length) return;
      const activeIndex = ranges.length - 1;
      const sel = ranges[activeIndex];
      const anchor = sel.end; // use end as active cell in the active range
      const extend = e.shiftKey;
      const move = (dr: number, dc: number) => {
        const nr = Math.min(Math.max(1, anchor.row + dr), sheet.rowCount);
        const nc = Math.min(Math.max(1, anchor.col + dc), sheet.colCount);
        const newEnd = { row: nr, col: nc };
        const newSel = extend ? { start: sel.start, end: newEnd } : { start: newEnd, end: newEnd };
        const next = ranges.slice();
        next[activeIndex] = newSel;
        r.setSelections(next);
        setSelections(next);
      };
      if (isArrowKey(e)) {
        e.preventDefault();
        switch (e.code) {
          case 'ArrowUp': move(-1, 0); break;
          case 'ArrowDown': move(1, 0); break;
          case 'ArrowLeft': move(0, -1); break;
          case 'ArrowRight': move(0, 1); break;
        }
      }
    };
    el.tabIndex = 0;
    el.addEventListener('keydown', onKeyDown);
    return () => { el.removeEventListener('keydown', onKeyDown); };
  }, [selections, rendererRef.current, sheetRef.current]);

  // Optional: snap to nearest cell boundary when scroll settles
  useEffect(() => {
    const r = rendererRef.current as any;
    if (!r) return;
    if (!p.snapToGrid) return;
    const snap = () => {
      const max = r.getMaxScroll?.(); if (!max) return;
      if (Math.abs(velRef.current.vx) < 0.1 && Math.abs(velRef.current.vy) < 0.1) {
        const { x, y } = r.getScroll();
        const snapX = findSnapX(r, x);
        const snapY = findSnapY(r, y);
        r.setScroll(snapX, snapY);
      }
    };
    const id = setInterval(snap, p.snapIntervalMs || 200);
    return () => clearInterval(id);
  }, [rendererRef.current, p.snapToGrid, p.snapIntervalMs]);

  const findSnapX = (r: any, sx: number) => {
    let acc = 0;
    const sheet = sheetRef.current as Worksheet;
    for (let c = 1; c <= sheet.colCount; c++) {
      const w = sheet.getColumnWidth(c);
      if (acc + w > sx) {
        // choose closer edge
        const distPrev = sx - acc;
        const distNext = acc + w - sx;
        return distPrev < distNext ? acc : acc + w;
      }
      acc += w;
    }
    return sx;
  };
  const findSnapY = (r: any, sy: number) => {
    let acc = 0;
    const sheet = sheetRef.current as Worksheet;
    for (let rr = 1; rr <= sheet.rowCount; rr++) {
      const h = sheet.getRowHeight(rr);
      if (acc + h > sy) {
        const distPrev = sy - acc;
        const distNext = acc + h - sy;
        return distPrev < distNext ? acc : acc + h;
      }
      acc += h;
    }
    return sy;
  };

  // Touch support: two-finger pan or single-finger select based on options
  useEffect(() => {
    const el = containerRef.current as any;
    const r = rendererRef.current as any;
    const sheet = sheetRef.current as Worksheet | undefined;
    if (!el || !r || !sheet || !p.allowTouch) return;

    const onTouchStart = (e: TouchEvent) => {
      if (!p.allowTouch) return;
      if (p.touchPanWithTwoFingers && e.touches.length === 2) {
        // Start panning
        const t1 = e.touches[0], t2 = e.touches[1];
        const cx = (t1.clientX + t2.clientX) / 2;
        const cy = (t1.clientY + t2.clientY) / 2;
        touchRef.current = { mode: 'pan', x: cx, y: cy, t: performance.now() };
        e.preventDefault();
      } else if (!p.touchPanWithTwoFingers && e.touches.length === 1) {
        // One-finger pan if configured
        const t = e.touches[0];
        touchRef.current = { mode: 'pan', x: t.clientX, y: t.clientY, t: performance.now() };
        e.preventDefault();
      } else if (e.touches.length === 1) {
        // Single finger select
        const t = e.touches[0];
        const hit = r.hitTest(t.clientX, t.clientY);
        if (hit && hit.type === 'cell') {
          const addr = (hit as any).addr;
          const sel = { start: addr, end: addr };
          r.setSelections([sel]);
          setSelections([sel]);
          touchRef.current = { mode: 'select', index: 0 };
        } else {
          touchRef.current = { mode: 'none' };
        }
      }
    };

    const onTouchMove = (e: TouchEvent) => {
      const st = touchRef.current;
      if (!st) return;
      if (st.mode === 'pan') {
        // Determine center point for 2-finger pan, or single finger
        const now = performance.now();
        let cx = 0, cy = 0;
        if (p.touchPanWithTwoFingers && e.touches.length >= 2) {
          const t1 = e.touches[0], t2 = e.touches[1]; cx = (t1.clientX + t2.clientX) / 2; cy = (t1.clientY + t2.clientY) / 2;
        } else if (!p.touchPanWithTwoFingers && e.touches.length >= 1) {
          const t = e.touches[0]; cx = t.clientX; cy = t.clientY;
        } else {
          return;
        }
        const dx = st.x - cx; // invert to match scroll direction
        const dy = st.y - cy;
        // Update velocity
        const dt = Math.max(1, now - (st.t || now));
        velRef.current.vx = dx / dt * 16;
        velRef.current.vy = dy / dt * 16;
        stepScroll(r, dx, dy, true);
        touchRef.current = { ...st, x: cx, y: cy, t: now };
        e.preventDefault();
      } else if (st.mode === 'select') {
        const t = e.touches[0]; if (!t) return;
        const now = performance.now();
        if (lastMouseRef.current.t) {
          const dt = Math.max(1, now - lastMouseRef.current.t);
          velRef.current.vx = (t.clientX - lastMouseRef.current.x) / dt * 16;
          velRef.current.vy = (t.clientY - lastMouseRef.current.y) / dt * 16;
        }
        lastMouseRef.current = { x: t.clientX, y: t.clientY, t: now };
        const addr = r.cellAt(t.clientX, t.clientY);
        if (addr) {
          const next = selectionsRef.current.slice();
          next[0] = { start: next[0]?.start || addr, end: addr };
          r.setSelections(next);
          setSelections(next);
          const rect = r.rectForRange?.(addr.row, addr.col, addr.row, addr.col);
          if (rect) {
            const margin = p.edgeMargin || 24;
            const step = p.autoScrollStep || 10;
            const opts = r.optionsReadonly ? r.optionsReadonly : { headerWidth: 48, headerHeight: 24 };
            if (rect.x < opts.headerWidth + margin) r.scrollBy(-step, 0);
            if (rect.y < opts.headerHeight + margin) r.scrollBy(0, -step);
            const vp = r.getViewportSize?.();
            if (vp) {
              if (rect.x + rect.w > opts.headerWidth + vp.width - margin) r.scrollBy(step, 0);
              if (rect.y + rect.h > opts.headerHeight + vp.height - margin) r.scrollBy(0, step);
            }
          }
        }
        e.preventDefault();
      }
    };

    const onTouchEnd = (e: TouchEvent) => {
      const st = touchRef.current;
      if (!st) return;
      if (st.mode === 'pan') {
        if (p.kineticScroll) {
          if (rafRef.current == null) rafRef.current = requestAnimationFrame(() => animate(r));
        } else {
          velRef.current.vx = 0; velRef.current.vy = 0;
        }
      } else if (st.mode === 'select') {
        if (p.selectionInertia) {
          const vx = Math.max(-60, Math.min(60, velRef.current.vx * 8));
          const vy = Math.max(-60, Math.min(60, velRef.current.vy * 8));
          inertialSelectingRef.current = { index: 0, dirX: 1, dirY: 1, vx, vy };
          if (rafRef.current == null) rafRef.current = requestAnimationFrame(() => animate(r));
        }
      }
      touchRef.current = null;
    };

    el.addEventListener('touchstart', onTouchStart, { passive: false });
    el.addEventListener('touchmove', onTouchMove, { passive: false });
    el.addEventListener('touchend', onTouchEnd, { passive: false });
    el.addEventListener('touchcancel', onTouchEnd, { passive: false });
    return () => {
      el.removeEventListener('touchstart', onTouchStart as any);
      el.removeEventListener('touchmove', onTouchMove as any);
      el.removeEventListener('touchend', onTouchEnd as any);
      el.removeEventListener('touchcancel', onTouchEnd as any);
    };
  }, [rendererRef.current, sheetRef.current, p.allowTouch, p.touchPanWithTwoFingers, p.kineticScroll, p.selectionInertia]);

  return (
    <div style={{ position: 'relative', width: '100%', height: '100%', ...style }}>
      <div
        ref={setContainerRef}
        style={{
          position: 'absolute',
          inset: horizontalScroll.maxX > 0 ? '0 0 16px 0' : 0,
        }}
      />
      {horizontalScroll.maxX > 0 && (
        <div
          ref={horizontalScrollbarRef}
          onScroll={(event) => {
            const nextX = event.currentTarget.scrollLeft;
            const renderer = rendererRef.current as CanvasRenderer | null;
            if (renderer && Math.abs(renderer.getScroll().x - nextX) > 1) {
              const current = renderer.getScroll();
              renderer.setScroll(nextX, current.y);
            }
          }}
          style={{
            position: 'absolute',
            left: 0,
            right: 0,
            bottom: 0,
            height: 16,
            overflowX: 'auto',
            overflowY: 'hidden',
            background: '#f3f3f3',
            borderTop: '1px solid #d9d9d9',
          }}
        >
          <div
            style={{
              width: Math.max(horizontalScroll.contentWidth, horizontalScroll.viewportWidth + horizontalScroll.maxX),
              height: 1,
            }}
          />
        </div>
      )}
      <CustomCellOverlay
        key={overlayTick}
        container={overlayHost}
        sheet={sheetRef.current ?? workbook.activeSheet}
        renderer={rendererRef.current}
        registry={componentRegistryRef.current}
        zoom={typeof zoom === 'number' ? zoom : 1}
        enabled={enableCustomComponents}
      />
    </div>
  );
});
