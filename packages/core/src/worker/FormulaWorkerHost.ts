import { FormulaEngine } from '../FormulaEngine';
import type { CellValue } from '../types';
import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Worksheet } from '../worksheet';
import {
  type FormulaWorkerOpName,
  type FormulaWorkerPayload,
  type FormulaWorkerRequest,
  type FormulaWorkerResponse,
  getFormulaWorkerResponseTransferList,
} from './FormulaWorkerProtocol';

export type FormulaWorkerSendFn = (response: FormulaWorkerResponse, transferList: Transferable[]) => void;

export class FormulaWorkerHost {
  private readonly engine = new FormulaEngine();
  private readonly worksheet: Worksheet;

  constructor(sheetName = 'Sheet1') {
    this.worksheet = new Worksheet(sheetName, DEFAULT_WORKSHEET_ROWS, DEFAULT_WORKSHEET_COLS, this.engine as any);
  }

  handleMessage(msg: FormulaWorkerRequest): { response: FormulaWorkerResponse; transferList: Transferable[] } {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const result = this.dispatch(msg) as any;
      const response: FormulaWorkerResponse = { id: msg.id, type: msg.type as FormulaWorkerOpName, ok: true, result };
      return { response, transferList: getFormulaWorkerResponseTransferList(response) };
    } catch (err) {
      const error = err instanceof Error ? err.message : String(err);
      return {
        response: { id: msg.id, type: msg.type as FormulaWorkerOpName, ok: false, error },
        transferList: [],
      };
    }
  }

  install(sendFn?: FormulaWorkerSendFn): void {
    const send: FormulaWorkerSendFn = sendFn ?? ((resp, transferList) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).postMessage(resp, transferList);
    });

    globalThis.addEventListener('message', (ev: Event | MessageEvent) => {
      const { response, transferList } = this.handleMessage((ev as MessageEvent).data as FormulaWorkerRequest);
      send(response, transferList);
    });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private dispatch(msg: FormulaWorkerRequest): any {
    switch (msg.type) {
      case 'ping':
        return 'pong';

      case 'setCellValue': {
        const { row, col, value } = msg.payload as FormulaWorkerPayload<'setCellValue'>;
        this.worksheet.setCellValue({ row, col }, value);
        return;
      }

      case 'setCellFormula': {
        const { row, col, formula, displayValue } = msg.payload as FormulaWorkerPayload<'setCellFormula'>;
        this.worksheet.setCellFormula({ row, col }, formula, displayValue);
        return;
      }

      case 'registerDeps': {
        const { row, col, deps } = msg.payload as FormulaWorkerPayload<'registerDeps'>;
        this.worksheet.registerDependencies({ row, col }, deps);
        return;
      }

      case 'evaluateBatch': {
        const { addresses } = msg.payload as FormulaWorkerPayload<'evaluateBatch'>;
        const recalc = this.worksheet.evaluateBatch(addresses);
        const values = new Float64Array(addresses.length);

        for (let i = 0; i < addresses.length; i++) {
          const value = this.worksheet.getCellValue(addresses[i]);
          values[i] = typeof value === 'number' ? value : Number.NaN;
        }

        return {
          values: values.buffer,
          evaluated: recalc.evaluated,
          hasCycles: recalc.cycles.length > 0,
        };
      }

      default:
        throw new Error(`FormulaWorkerHost: unknown message type "${(msg as FormulaWorkerRequest).type}".`);
    }
  }
}
