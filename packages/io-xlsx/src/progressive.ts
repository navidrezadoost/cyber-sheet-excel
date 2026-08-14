import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Workbook, type CellStyle, type CellValue } from '@cyber-sheet/core';
import { streamXLSX, type ImportOptions } from './import';
import type { XLSXMetadata } from './LightweightParser';

export type ProgressiveXLSXCell = {
  ref: string;
  value: CellValue;
  formula?: string;
  style?: CellStyle;
};

export type ProgressiveXLSXChunk = {
  sheetName: string;
  startRow: number;
  endRow: number;
  totalRows: number;
  cells: ProgressiveXLSXCell[];
};

type WorkerMessage =
  | { type: 'metadata'; metadata: SerializedMetadata }
  | { type: 'chunk'; chunk: ProgressiveXLSXChunk }
  | { type: 'done' }
  | { type: 'error'; error: string };

type SerializedMetadata = {
  sheetNames: string[];
  sheetDimensions: Array<[string, { rows: number; cols: number }]>;
  styleCount: number;
  sharedStringCount: number;
  fileSize: number;
};

export type ProgressiveXLSXOptions = ImportOptions & {
  chunkRows?: number;
  workerFactory?: () => Worker;
  onChunk?: (chunk: ProgressiveXLSXChunk) => void;
  onProgress?: (progress: { sheetName: string; loadedRows: number; totalRows: number; done: boolean }) => void;
};

export type ProgressiveXLSXSession = {
  workbook: Workbook;
  metadata: XLSXMetadata;
  done: Promise<void>;
  cancel: () => void;
};

function sheetDisplayDimensions(rows: number, cols: number): { rows: number; cols: number } {
  return {
    rows: Math.max(DEFAULT_WORKSHEET_ROWS, rows),
    cols: Math.max(DEFAULT_WORKSHEET_COLS, cols),
  };
}

export async function loadXlsxProgressivelyFromArrayBuffer(
  buffer: ArrayBuffer | Uint8Array,
  options: ProgressiveXLSXOptions = {},
): Promise<ProgressiveXLSXSession> {
  const source = buffer instanceof Uint8Array
    ? buffer.slice().buffer as ArrayBuffer
    : buffer;

  if (typeof Worker === 'undefined' && !options.workerFactory) {
    return loadProgressivelyOnMainThread(source, options);
  }

  const worker = options.workerFactory?.() ?? new Worker(
    new URL('./worker/progressive-xlsx-worker.ts', import.meta.url),
    { type: 'module' },
  );

  let workbook: Workbook | null = null;
  let metadata: XLSXMetadata | null = null;
  let settled = false;
  let resolveReady!: (session: ProgressiveXLSXSession) => void;
  let rejectReady!: (error: Error) => void;
  let resolveDone!: () => void;
  let rejectDone!: (error: Error) => void;

  const ready = new Promise<ProgressiveXLSXSession>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });
  const done = new Promise<void>((resolve, reject) => {
    resolveDone = resolve;
    rejectDone = reject;
  });

  const cancel = () => {
    if (settled) return;
    settled = true;
    worker.terminate();
    rejectReady(new Error('Progressive XLSX load cancelled.'));
    rejectDone(new Error('Progressive XLSX load cancelled.'));
  };

  worker.addEventListener('message', (ev: MessageEvent<WorkerMessage>) => {
    const msg = ev.data;
    if (msg.type === 'metadata') {
      metadata = deserializeMetadata(msg.metadata);
      workbook = createProgressiveWorkbook(metadata);
      resolveReady({ workbook, metadata, done, cancel });
      return;
    }

    if (msg.type === 'chunk') {
      if (!workbook) return;
      applyChunk(workbook, msg.chunk);
      options.onChunk?.(msg.chunk);
      options.onProgress?.({
        sheetName: msg.chunk.sheetName,
        loadedRows: msg.chunk.endRow,
        totalRows: msg.chunk.totalRows,
        done: false,
      });
      return;
    }

    if (msg.type === 'done') {
      settled = true;
      finishProgressiveWorkbook(workbook);
      worker.terminate();
      resolveDone();
      return;
    }

    const error = new Error(msg.error);
    settled = true;
    worker.terminate();
    rejectReady(error);
    rejectDone(error);
  });

  worker.postMessage({
    type: 'start',
    buffer: source,
    options: {
      ...options,
      workerFactory: undefined,
      onChunk: undefined,
      onProgress: undefined,
    },
  }, [source]);

  return ready;
}

async function loadProgressivelyOnMainThread(
  buffer: ArrayBuffer,
  options: ProgressiveXLSXOptions,
): Promise<ProgressiveXLSXSession> {
  const iterator = streamXLSX(buffer, { ...options, chunkRows: options.chunkRows ?? 500 } as ImportOptions);
  const first = await iterator.next();
  if (first.done) {
    const workbook = new Workbook();
    return { workbook, metadata: emptyMetadata(buffer.byteLength), done: Promise.resolve(), cancel: () => {} };
  }

  const dims = new Map([[first.value.sheetName, {
    rows: first.value.progress.total,
    cols: Math.max(DEFAULT_WORKSHEET_COLS, maxChunkCol(first.value.cells)),
  }]]);
  const metadata = {
    sheetNames: [first.value.sheetName],
    sheetDimensions: dims,
    styleCount: 0,
    sharedStringCount: 0,
    fileSize: buffer.byteLength,
  };
  const workbook = createProgressiveWorkbook(metadata);

  const done = (async () => {
    applyChunk(workbook, mapStreamChunk(first.value, options.chunkRows ?? 500));
    for await (const chunk of iterator) {
      const mapped = mapStreamChunk(chunk, options.chunkRows ?? 500);
      if (!workbook.getSheet(mapped.sheetName)) {
        const displayDims = sheetDisplayDimensions(mapped.totalRows, maxCellCol(mapped.cells));
        workbook.addSheet(mapped.sheetName, displayDims.rows, displayDims.cols).beginProgressiveLoad(displayDims.rows);
      }
      applyChunk(workbook, mapped);
      options.onChunk?.(mapped);
      options.onProgress?.({ sheetName: mapped.sheetName, loadedRows: mapped.endRow, totalRows: mapped.totalRows, done: false });
    }
    finishProgressiveWorkbook(workbook);
  })();

  return { workbook, metadata, done, cancel: () => {} };
}

function createProgressiveWorkbook(metadata: XLSXMetadata): Workbook {
  const workbook = new Workbook();
  metadata.sheetNames.forEach((sheetName) => {
    const dims = metadata.sheetDimensions.get(sheetName) ?? { rows: DEFAULT_WORKSHEET_ROWS, cols: DEFAULT_WORKSHEET_COLS };
    const displayDims = sheetDisplayDimensions(dims.rows, dims.cols);
    const sheet = workbook.addSheet(sheetName, displayDims.rows, displayDims.cols);
    sheet.beginProgressiveLoad(displayDims.rows);
  });
  return workbook;
}

function finishProgressiveWorkbook(workbook: Workbook | null): void {
  if (!workbook) return;
  for (const sheetName of workbook.getSheetNames()) {
    workbook.getSheet(sheetName)?.finishProgressiveLoad();
  }
}

function applyChunk(workbook: Workbook, chunk: ProgressiveXLSXChunk): void {
  let sheet = workbook.getSheet(chunk.sheetName);
  if (!sheet) {
    const displayDims = sheetDisplayDimensions(chunk.totalRows, maxCellCol(chunk.cells));
    sheet = workbook.addSheet(chunk.sheetName, displayDims.rows, displayDims.cols);
    sheet.beginProgressiveLoad(displayDims.rows);
  }

  for (const cell of chunk.cells) {
    const addr = a1ToAddress(cell.ref);
    if (!addr) continue;
    if (cell.formula) {
      sheet.setCellFormula(addr, cell.formula, cell.value);
    } else if (cell.value !== null && cell.value !== undefined) {
      sheet.setCellValue(addr, cell.value);
    }
    if (cell.style) sheet.setCellStyle(addr, cell.style);
  }

  sheet.markRowsLoaded(chunk.startRow, chunk.endRow);
}

function mapStreamChunk(chunk: { sheetName: string; cells: Map<string, { value: CellValue; style?: CellStyle; formula?: string }>; progress: { row: number; total: number } }, chunkRows: number): ProgressiveXLSXChunk {
  const cells = Array.from(chunk.cells, ([ref, cell]) => ({ ref, ...cell }));
  const endRow = chunk.progress.row;
  const startRow = Math.max(1, endRow - chunkRows + 1);
  return { sheetName: chunk.sheetName, startRow, endRow, totalRows: chunk.progress.total, cells };
}

function deserializeMetadata(metadata: SerializedMetadata): XLSXMetadata {
  return {
    ...metadata,
    sheetDimensions: new Map(metadata.sheetDimensions),
  };
}

function emptyMetadata(fileSize: number): XLSXMetadata {
  return {
    sheetNames: [],
    sheetDimensions: new Map(),
    styleCount: 0,
    sharedStringCount: 0,
    fileSize,
  };
}

function a1ToAddress(ref: string): { row: number; col: number } | null {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  let col = 0;
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64);
  }
  return { row: Number(match[2]), col };
}

function maxChunkCol(cells: Map<string, unknown>): number {
  return maxCellCol(Array.from(cells.keys()).map(ref => ({ ref })));
}

function maxCellCol(cells: Array<{ ref: string }>): number {
  let max = 1;
  for (const cell of cells) {
    max = Math.max(max, a1ToAddress(cell.ref)?.col ?? 1);
  }
  return max;
}
