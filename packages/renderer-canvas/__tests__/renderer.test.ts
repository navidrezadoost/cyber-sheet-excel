import { mockCanvas } from '@cyber-sheet/test-utils';
import { Workbook } from '@cyber-sheet/core';
import { CanvasRenderer } from '../src/CanvasRenderer';

describe('CanvasRenderer', () => {
  beforeAll(() => {
    mockCanvas();
  });

  function setViewport(container: HTMLElement, width = 1000, height = 600) {
    Object.defineProperty(container, 'clientWidth', { value: width, configurable: true });
    Object.defineProperty(container, 'clientHeight', { value: height, configurable: true });
  }

  it('creates renderer instance without crashing', () => {
    const container = document.createElement('div');
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1');
    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    expect(renderer).toBeTruthy();
  });

  it('resizes all columns when full sheet is selected', () => {
    const container = document.createElement('div');
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1');
    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    const anyRenderer = renderer as any;

    const canvas = anyRenderer.canvas as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => ({
      left: 0, top: 0, right: 1000, bottom: 600, width: 1000, height: 600, x: 0, y: 0, toJSON: () => ({})
    } as DOMRect);

    renderer.setSelections([{
      start: { row: 1, col: 1 },
      end: { row: sheet.rowCount, col: sheet.colCount },
    }]);

    const startWidth = sheet.getColumnWidth(1);
    const resizeHandleX = 48 + startWidth;

    anyRenderer.handleMouseDown(new MouseEvent('mousedown', { clientX: resizeHandleX, clientY: 10 }));
    anyRenderer.pendingMouseX = resizeHandleX + 30;
    anyRenderer.pendingMouseY = 10;
    anyRenderer.processHover();
    anyRenderer.handleMouseUp(new MouseEvent('mouseup', { clientX: resizeHandleX + 30, clientY: 10 }));

    expect(sheet.getColumnWidth(1)).toBe(startWidth + 30);
    expect(sheet.getColumnWidth(2)).toBe(startWidth + 30);
    expect(sheet.getColumnWidth(sheet.colCount)).toBe(startWidth + 30);
  });

  it('resizes all rows when full sheet is selected', () => {
    const container = document.createElement('div');
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1');
    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    const anyRenderer = renderer as any;

    const canvas = anyRenderer.canvas as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => ({
      left: 0, top: 0, right: 1000, bottom: 600, width: 1000, height: 600, x: 0, y: 0, toJSON: () => ({})
    } as DOMRect);

    renderer.setSelections([{
      start: { row: 1, col: 1 },
      end: { row: sheet.rowCount, col: sheet.colCount },
    }]);

    const startHeight = sheet.getRowHeight(1);
    const resizeHandleY = 24 + startHeight;

    anyRenderer.handleMouseDown(new MouseEvent('mousedown', { clientX: 10, clientY: resizeHandleY }));
    anyRenderer.pendingMouseX = 10;
    anyRenderer.pendingMouseY = resizeHandleY + 20;
    anyRenderer.processHover();
    anyRenderer.handleMouseUp(new MouseEvent('mouseup', { clientX: 10, clientY: resizeHandleY + 20 }));

    expect(sheet.getRowHeight(1)).toBe(startHeight + 20);
    expect(sheet.getRowHeight(2)).toBe(startHeight + 20);
    expect(sheet.getRowHeight(sheet.rowCount)).toBe(startHeight + 20);
  });

  it('maps pixels to cells with cached cumulative row and column offsets', () => {
    const container = document.createElement('div');
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1');
    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    const anyRenderer = renderer as any;

    const canvas = anyRenderer.canvas as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => ({
      left: 0, top: 0, right: 1000, bottom: 600, width: 1000, height: 600, x: 0, y: 0, toJSON: () => ({})
    } as DOMRect);

    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 120);
    sheet.setRowHeight(1, 20);
    sheet.setRowHeight(2, 40);
    renderer.setScroll(50, 20);

    expect(renderer.cellAt(58, 34)).toEqual({ row: 2, col: 2 });

    sheet.setColumnWidth(1, 80);
    renderer.setScroll(80, 20);

    expect(renderer.cellAt(58, 34)).toEqual({ row: 2, col: 2 });
    renderer.dispose();
  });

  it('skips hidden rows and columns in hit testing and range geometry', () => {
    const container = document.createElement('div');
    setViewport(container);
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1', 10, 10);
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 70);
    sheet.setRowHeight(1, 20);
    sheet.setRowHeight(2, 40);
    sheet.setRowHeight(3, 30);
    sheet.hideCol(2);
    sheet.hideRow(2);

    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    const anyRenderer = renderer as any;
    const canvas = anyRenderer.canvas as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => ({
      left: 0, top: 0, right: 1000, bottom: 600, width: 1000, height: 600, x: 0, y: 0, toJSON: () => ({})
    } as DOMRect);

    expect(renderer.cellAt(48 + 50 + 1, 24 + 20 + 1)).toEqual({ row: 3, col: 3 });
    expect(anyRenderer.rectForRange(1, 1, 3, 3)).toEqual({
      x: 48,
      y: 24,
      w: 120,
      h: 50,
    });

    renderer.dispose();
  });

  it('uses visible tracks for merged-cell rectangles', () => {
    const container = document.createElement('div');
    setViewport(container);
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1', 10, 10);
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 70);
    sheet.setRowHeight(1, 20);
    sheet.setRowHeight(2, 40);
    sheet.setRowHeight(3, 30);
    sheet.mergeCells({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    sheet.hideCol(2);
    sheet.hideRow(2);

    const renderer = new CanvasRenderer(container, sheet, { debug: false });
    const rect = (renderer as any).rectForRange(1, 1, 3, 3);

    expect(rect).toEqual({
      x: 48,
      y: 24,
      w: 120,
      h: 50,
    });

    renderer.dispose();
  });

  it('caches canvas image loads and redraws after the image resolves', () => {
    const container = document.createElement('div');
    setViewport(container);
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1', 10, 10);
    const originalImage = (globalThis as any).Image;
    const created: any[] = [];
    class FakeImage {
      crossOrigin = '';
      complete = false;
      naturalWidth = 0;
      naturalHeight = 0;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      private _src = '';
      set src(value: string) { this._src = value; }
      get src() { return this._src; }
    }
    (globalThis as any).Image = function Image() {
      const image = new FakeImage();
      created.push(image);
      return image;
    };

    const renderer = new CanvasRenderer(container, sheet, {
      debug: false,
      images: [{
        source: 'data:image/png;base64,abc',
        anchor: { row: 1, col: 1 },
      }],
    });

    renderer.redraw();
    expect(created).toHaveLength(1);
    expect((renderer as any).imageCache.get('data:image/png;base64,abc').status).toBe('loading');

    created[0].complete = true;
    created[0].naturalWidth = 16;
    created[0].naturalHeight = 16;
    created[0].onload();
    renderer.redraw();

    expect(created).toHaveLength(1);
    expect((renderer as any).imageCache.get('data:image/png;base64,abc').status).toBe('loaded');

    renderer.dispose();
    (globalThis as any).Image = originalImage;
  });

  it('uses formula worker batches for large visible dirty sets', async () => {
    const container = document.createElement('div');
    setViewport(container);
    const workbook = new Workbook();
    const sheet = workbook.addSheet('Sheet1');
    sheet.setCellValue({ row: 1, col: 1 }, 5);
    sheet.setCellFormula({ row: 1, col: 2 }, '=A1*2');
    sheet.registerDependencies({ row: 1, col: 2 }, [{ row: 1, col: 1 }]);

    const worker = {
      setCellValue: jest.fn().mockResolvedValue(undefined),
      setCellFormula: jest.fn().mockResolvedValue(undefined),
      evaluateBatch: jest.fn().mockResolvedValue({
        values: new Float64Array([10]),
        evaluated: 1,
        hasCycles: false,
      }),
      terminate: jest.fn(),
    };

    const renderer = new CanvasRenderer(container, sheet, {
      debug: false,
      formulaWorkerThreshold: 1,
      formulaWorkerFactory: () => worker,
    });

    (renderer as any).evaluateVisibleFormulaCells();
    await Promise.resolve();
    await Promise.resolve();

    expect(worker.evaluateBatch).toHaveBeenCalledWith([{ row: 1, col: 2 }]);
    expect(worker.setCellValue).toHaveBeenCalledWith(1, 1, 5);
    expect(worker.setCellFormula).toHaveBeenCalledWith(1, 2, '=A1*2', null);
    expect(worker.setCellValue.mock.invocationCallOrder[0]).toBeLessThan(worker.evaluateBatch.mock.invocationCallOrder[0]);
    expect(worker.setCellFormula.mock.invocationCallOrder[0]).toBeLessThan(worker.evaluateBatch.mock.invocationCallOrder[0]);
    expect(sheet.getCellValue({ row: 1, col: 2 })).toBe(10);

    renderer.dispose();
    expect(worker.terminate).toHaveBeenCalled();
  });
});
