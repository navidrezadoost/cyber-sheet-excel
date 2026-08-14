/**
 * KeyboardShortcutManager.ts — Phase 22: Keyboard Shortcuts
 *
 * Framework-agnostic keyboard shortcut manager for @cyber-sheet/core SDK.
 *
 * Features:
 *   - Maintains its own selection cursor (activeRow, activeCol, anchorRow, anchorCol).
 *   - Maps normalised key combos to ShortcutHandler functions.
 *   - ~40 built-in Excel-parity shortcuts wired to SpreadsheetSDK methods.
 *   - Fully open for customisation: bind() / unbind() / resetToDefaults().
 *   - Optional callbacks for UI-level actions (Find, Replace, GoTo, Edit,
 *     Format) that the headless SDK cannot implement directly.
 *
 * Combo format:  modifiers in fixed order ctrl+meta+shift+alt, then key name
 *   (all lowercase), joined by +.  Examples: ctrl+z, ctrl+shift+end, f2,
 *   delete, arrowleft.
 */

import type { SpreadsheetSDK } from './SpreadsheetSDK';
import type { Range } from '../types';
import { DisposedError } from './errors';

export type KeyLike = {
  key: string;
  ctrlKey?:  boolean;
  shiftKey?: boolean;
  altKey?:   boolean;
  metaKey?:  boolean;
};

export type ShortcutHandler = (
  sdk: SpreadsheetSDK,
  manager: KeyboardShortcutManager,
  event: KeyLike,
) => boolean | void;

export interface KeyboardShortcutManagerOptions {
  initialRow?: number;
  initialCol?: number;
  pageSize?: number;
  onSelectionChange?: (activeRow: number, activeCol: number, anchorRow: number | null, anchorCol: number | null) => void;
  onFindRequest?: () => boolean | void;
  onReplaceRequest?: () => boolean | void;
  onGoToRequest?: () => boolean | void;
  onEditRequest?: (row: number, col: number) => boolean | void;
  onEscapeRequest?: (row: number, col: number) => boolean | void;
  onFormatRequest?: (action: 'bold' | 'italic' | 'underline' | 'format-cells', row: number, col: number) => boolean | void;
}

export interface KeyboardShortcutManager {
  readonly activeRow: number;
  readonly activeCol: number;
  readonly anchorRow: number | null;
  readonly anchorCol: number | null;
  readonly selectionRange: Range | null;
  setActiveCell(row: number, col: number, keepAnchor?: boolean): void;
  handleKeyEvent(event: KeyLike): boolean;
  bind(combo: string | string[], handler: ShortcutHandler): void;
  unbind(combo: string | string[]): void;
  resetToDefaults(): void;
  getBoundCombos(): string[];
  dispose(): void;
}

export function normaliseCombo(comboOrEvent: string | KeyLike): string {
  if (typeof comboOrEvent === 'string') {
    const parts  = comboOrEvent.toLowerCase().split(/\s*\+\s*/).filter(Boolean);
    const modSet = new Set(['ctrl', 'meta', 'cmd', 'command', 'shift', 'alt']);
    const ctrl   = parts.includes('ctrl');
    const meta   = parts.includes('meta') || parts.includes('cmd') || parts.includes('command');
    const shift  = parts.includes('shift');
    const alt    = parts.includes('alt');
    const key    = parts.filter(p => !modSet.has(p))[0] ?? '';
    return _mkCombo(ctrl, meta, shift, alt, key);
  }
  return _mkCombo(
    comboOrEvent.ctrlKey  ?? false,
    comboOrEvent.metaKey  ?? false,
    comboOrEvent.shiftKey ?? false,
    comboOrEvent.altKey   ?? false,
    _normaliseKeyName(comboOrEvent.key),
  );
}

function _normaliseKeyName(key: string): string {
  if (key === ' ') return 'space';
  return key.toLowerCase();
}

function _mkCombo(ctrl: boolean, meta: boolean, shift: boolean, alt: boolean, key: string): string {
  const parts: string[] = [];
  if (ctrl)  parts.push('ctrl');
  if (meta)  parts.push('meta');
  if (shift) parts.push('shift');
  if (alt)   parts.push('alt');
  parts.push(key);
  return parts.join('+');
}

class KeyboardShortcutManagerImpl implements KeyboardShortcutManager {
  private readonly _sdk:       SpreadsheetSDK;
  private readonly _opts:      KeyboardShortcutManagerOptions;
  private readonly _pageSize:  number;
  private          _handlers:  Map<string, ShortcutHandler> = new Map();
  private          _activeRow: number;
  private          _activeCol: number;
  private          _anchorRow: number | null = null;
  private          _anchorCol: number | null = null;
  private          _disposed   = false;

  constructor(sdk: SpreadsheetSDK, opts: KeyboardShortcutManagerOptions) {
    this._sdk      = sdk;
    this._opts     = opts;
    this._pageSize = opts.pageSize ?? 10;
    this._activeRow = Math.max(1, Math.min(opts.initialRow ?? 1, sdk.rowCount));
    this._activeCol = Math.max(1, Math.min(opts.initialCol ?? 1, sdk.colCount));
    this._installDefaults();
  }

  get activeRow(): number { return this._activeRow; }
  get activeCol(): number { return this._activeCol; }
  get anchorRow(): number | null { return this._anchorRow; }
  get anchorCol(): number | null { return this._anchorCol; }

  get selectionRange(): Range | null {
    if (this._anchorRow == null || this._anchorCol == null) return null;
    const r1 = Math.min(this._anchorRow, this._activeRow);
    const r2 = Math.max(this._anchorRow, this._activeRow);
    const c1 = Math.min(this._anchorCol, this._activeCol);
    const c2 = Math.max(this._anchorCol, this._activeCol);
    return { start: { row: r1, col: c1 }, end: { row: r2, col: c2 } };
  }

  setActiveCell(row: number, col: number, keepAnchor = false): void {
    this._guard('setActiveCell');
    const r = Math.max(1, Math.min(row, this._sdk.rowCount));
    const c = Math.max(1, Math.min(col, this._sdk.colCount));
    this._setActive(r, c, keepAnchor);
  }

  handleKeyEvent(event: KeyLike): boolean {
    this._guard('handleKeyEvent');
    const handler = this._handlers.get(normaliseCombo(event));
    if (!handler) return false;
    return handler(this._sdk, this, event) !== false;
  }

  bind(combo: string | string[], handler: ShortcutHandler): void {
    this._guard('bind');
    for (const c of (Array.isArray(combo) ? combo : [combo])) {
      this._handlers.set(normaliseCombo(c), handler);
    }
  }

  unbind(combo: string | string[]): void {
    this._guard('unbind');
    for (const c of (Array.isArray(combo) ? combo : [combo])) {
      this._handlers.delete(normaliseCombo(c));
    }
  }

  resetToDefaults(): void {
    this._guard('resetToDefaults');
    this._handlers.clear();
    this._installDefaults();
  }

  getBoundCombos(): string[] {
    this._guard('getBoundCombos');
    return [...this._handlers.keys()].sort();
  }

  dispose(): void {
    this._disposed = true;
    this._handlers.clear();
  }

  // ── internal navigation API (used by built-in handlers) ────────────────────

  _move(dr: number, dc: number): { row: number; col: number } {
    const row = Math.max(1, Math.min(this._activeRow + dr, this._sdk.rowCount));
    const col = Math.max(1, Math.min(this._activeCol + dc, this._sdk.colCount));
    this._setActive(row, col, false);
    return { row, col };
  }

  _extend(dr: number, dc: number): void {
    if (this._anchorRow == null) {
      this._anchorRow = this._activeRow;
      this._anchorCol = this._activeCol;
    }
    const row = Math.max(1, Math.min(this._activeRow + dr, this._sdk.rowCount));
    const col = Math.max(1, Math.min(this._activeCol + dc, this._sdk.colCount));
    this._setActive(row, col, true);
  }

  _jumpToEdge(dr: number, dc: number, extend: boolean): void {
    const sdk = this._sdk;
    let r = this._activeRow;
    let c = this._activeCol;
    const maxR = sdk.rowCount;
    const maxC = sdk.colCount;

    if (sdk.getCellValue(r, c) !== null) {
      while (true) {
        const nr = r + dr, nc = c + dc;
        if (nr < 1 || nr > maxR || nc < 1 || nc > maxC) break;
        if (sdk.getCellValue(nr, nc) === null) break;
        r = nr; c = nc;
      }
    } else {
      while (true) {
        const nr = r + dr, nc = c + dc;
        if (nr < 1 || nr > maxR || nc < 1 || nc > maxC) {
          r = dr < 0 ? 1 : dr > 0 ? maxR : r;
          c = dc < 0 ? 1 : dc > 0 ? maxC : c;
          break;
        }
        r = nr; c = nc;
        if (sdk.getCellValue(r, c) !== null) break;
      }
    }

    if (extend) {
      if (this._anchorRow == null) {
        this._anchorRow = this._activeRow;
        this._anchorCol = this._activeCol;
      }
      this._setActive(r, c, true);
    } else {
      this._setActive(r, c, false);
    }
  }

  // ── private ────────────────────────────────────────────────────────────────

  private _guard(method: string): void {
    if (this._disposed)
      throw new DisposedError(method, 'KeyboardShortcutManager');
  }

  private _setActive(row: number, col: number, keepAnchor: boolean): void {
    if (!keepAnchor) { this._anchorRow = null; this._anchorCol = null; }
    this._activeRow = row;
    this._activeCol = col;
    this._opts.onSelectionChange?.(row, col, this._anchorRow, this._anchorCol);
  }

  private _installDefaults(): void {
    const $ = (combos: string | string[], h: ShortcutHandler): void => {
      for (const c of (Array.isArray(combos) ? combos : [combos])) {
        this._handlers.set(normaliseCombo(c), h);
      }
    };
    const opts = this._opts;

    // ── Navigation ─────────────────────────────────────────────────────────
    $('arrowleft',  (s, m) => { (m as KeyboardShortcutManagerImpl)._move(0, -1);  return true; });
    $('arrowright', (s, m) => { (m as KeyboardShortcutManagerImpl)._move(0,  1);  return true; });
    $('arrowup',    (s, m) => { (m as KeyboardShortcutManagerImpl)._move(-1, 0);  return true; });
    $('arrowdown',  (s, m) => { (m as KeyboardShortcutManagerImpl)._move(1,  0);  return true; });

    $('tab', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      if (mi._activeCol >= s.colCount) mi._move(1, 1 - mi._activeCol);
      else                             mi._move(0, 1);
      return true;
    });
    $('shift+tab', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      if (mi._activeCol <= 1) mi._move(-1, s.colCount - mi._activeCol);
      else                    mi._move(0, -1);
      return true;
    });

    $('enter',       (s, m) => { (m as KeyboardShortcutManagerImpl)._move(1,  0); return true; });
    $('shift+enter', (s, m) => { (m as KeyboardShortcutManagerImpl)._move(-1, 0); return true; });

    $('home',  (s, m) => { (m as KeyboardShortcutManagerImpl)._setActive(m.activeRow, 1,           false); return true; });
    $('end',   (s, m) => { (m as KeyboardShortcutManagerImpl)._setActive(m.activeRow, s.colCount,  false); return true; });
    $('ctrl+home', (s, m) => { (m as KeyboardShortcutManagerImpl)._setActive(1, 1, false); return true; });
    $('ctrl+end',  (s, m) => {
      let maxR = 1, maxC = 1;
      outer: for (let r = s.rowCount; r >= 1; r--) {
        for (let c = s.colCount; c >= 1; c--) {
          if (s.getCellValue(r, c) !== null) { maxR = r; if (c > maxC) maxC = c; break outer; }
        }
      }
      (m as KeyboardShortcutManagerImpl)._setActive(maxR, maxC, false);
      return true;
    });

    $('pageup',   (s, m) => { (m as KeyboardShortcutManagerImpl)._move(-(m as KeyboardShortcutManagerImpl)._pageSize, 0); return true; });
    $('pagedown', (s, m) => { (m as KeyboardShortcutManagerImpl)._move( (m as KeyboardShortcutManagerImpl)._pageSize, 0); return true; });

    $('ctrl+arrowleft',  (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(0, -1, false); return true; });
    $('ctrl+arrowright', (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(0,  1, false); return true; });
    $('ctrl+arrowup',    (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(-1, 0, false); return true; });
    $('ctrl+arrowdown',  (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(1,  0, false); return true; });

    // ── Selection extension ────────────────────────────────────────────────
    $('shift+arrowleft',  (s, m) => { (m as KeyboardShortcutManagerImpl)._extend(0, -1); return true; });
    $('shift+arrowright', (s, m) => { (m as KeyboardShortcutManagerImpl)._extend(0,  1); return true; });
    $('shift+arrowup',    (s, m) => { (m as KeyboardShortcutManagerImpl)._extend(-1, 0); return true; });
    $('shift+arrowdown',  (s, m) => { (m as KeyboardShortcutManagerImpl)._extend(1,  0); return true; });

    $('ctrl+shift+arrowleft',  (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(0, -1, true); return true; });
    $('ctrl+shift+arrowright', (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(0,  1, true); return true; });
    $('ctrl+shift+arrowup',    (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(-1, 0, true); return true; });
    $('ctrl+shift+arrowdown',  (s, m) => { (m as KeyboardShortcutManagerImpl)._jumpToEdge(1,  0, true); return true; });

    $('shift+home', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      if (mi._anchorRow == null) { mi._anchorRow = mi._activeRow; mi._anchorCol = mi._activeCol; }
      mi._setActive(mi._activeRow, 1, true); return true;
    });
    $('ctrl+shift+home', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      if (mi._anchorRow == null) { mi._anchorRow = mi._activeRow; mi._anchorCol = mi._activeCol; }
      mi._setActive(1, 1, true); return true;
    });

    $('ctrl+a', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      const range = mi._getUsedRange();
      mi._anchorRow = range.start.row; mi._anchorCol = range.start.col;
      mi._setActive(range.end.row, range.end.col, true); return true;
    });
    $('ctrl+space', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      mi._anchorRow = 1; mi._anchorCol = mi._activeCol;
      mi._setActive(s.rowCount, mi._activeCol, true); return true;
    });
    $('shift+space', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      mi._anchorRow = mi._activeRow; mi._anchorCol = 1;
      mi._setActive(mi._activeRow, s.colCount, true); return true;
    });

    // ── Edit ──────────────────────────────────────────────────────────────
    $(['delete', 'backspace'], (s, m) => {
      try { s.setCell(m.activeRow, m.activeCol, null); } catch { /* protected */ }
      return true;
    });

    $('ctrl+;', (s, m) => {
      const today = new Date().toISOString().slice(0, 10);
      try { s.setCell(m.activeRow, m.activeCol, today); } catch { /* */ }
      return true;
    });

    $('ctrl+shift+:', (s, m) => {
      const d = new Date();
      const time = `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
      try { s.setCell(m.activeRow, m.activeCol, time); } catch { /* */ }
      return true;
    });

    $('ctrl+d', (s, m) => {
      const sel = m.selectionRange;
      if (!sel || sel.start.row === sel.end.row) return false;
      const srcR = sel.start.row;
      for (let c = sel.start.col; c <= sel.end.col; c++) {
        const val = s.getCellValue(srcR, c);
        for (let r = srcR + 1; r <= sel.end.row; r++) {
          try { s.setCell(r, c, val); } catch { /* */ }
        }
      }
      return true;
    });

    $('ctrl+r', (s, m) => {
      const sel = m.selectionRange;
      if (!sel || sel.start.col === sel.end.col) return false;
      const srcC = sel.start.col;
      for (let r = sel.start.row; r <= sel.end.row; r++) {
        const val = s.getCellValue(r, srcC);
        for (let c = srcC + 1; c <= sel.end.col; c++) {
          try { s.setCell(r, c, val); } catch { /* */ }
        }
      }
      return true;
    });

    // ── Undo / Redo ────────────────────────────────────────────────────────
    $('ctrl+z', (s) => s.undo());
    $(['ctrl+y', 'ctrl+shift+z', 'f4'], (s) => s.redo());

    // ── Dialog callbacks ───────────────────────────────────────────────────
    $('ctrl+f',         () => { if (!opts.onFindRequest)    return false; return opts.onFindRequest()    !== false; });
    $('ctrl+h',         () => { if (!opts.onReplaceRequest) return false; return opts.onReplaceRequest() !== false; });
    $(['ctrl+g', 'f5'], () => { if (!opts.onGoToRequest)    return false; return opts.onGoToRequest()    !== false; });
    $('f2', (s, m) => {
      if (!opts.onEditRequest) return false;
      return opts.onEditRequest(m.activeRow, m.activeCol) !== false;
    });

    $('escape', (s, m) => {
      const mi = m as KeyboardShortcutManagerImpl;
      if (opts.onEscapeRequest?.(m.activeRow, m.activeCol) === true) return true;
      if (mi._anchorRow !== null) {
        mi._anchorRow = null; mi._anchorCol = null;
        opts.onSelectionChange?.(mi._activeRow, mi._activeCol, null, null);
        return true;
      }
      return false;
    });

    // ── Format callbacks ───────────────────────────────────────────────────
    const _fmt = (action: 'bold' | 'italic' | 'underline' | 'format-cells'): ShortcutHandler =>
      (s, m) => {
        if (!opts.onFormatRequest) return false;
        return opts.onFormatRequest(action, m.activeRow, m.activeCol) !== false;
      };
    $('ctrl+b', _fmt('bold'));
    $('ctrl+i', _fmt('italic'));
    $('ctrl+u', _fmt('underline'));
    $('ctrl+1', _fmt('format-cells'));

    // ── Filter ─────────────────────────────────────────────────────────────
    $('ctrl+shift+l', (s, m) => {
      if (s.getAutoFilterRange()) {
        s.clearAutoFilterRange();
      } else {
        const selection = m.selectionRange;
        if (selection) {
          s.setAutoFilterRange(selection.start.row, selection.start.col, selection.end.col);
        } else {
          s.setAutoFilterRange(m.activeRow, 1, s.colCount);
        }
      }
      return true;
    });
  }

  private _getUsedRange(): Range {
    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxRow = 0;
    let maxCol = 0;

    for (let row = 1; row <= this._sdk.rowCount; row++) {
      for (let col = 1; col <= this._sdk.colCount; col++) {
        const cell = this._sdk.getCell(row, col);
        const hasValue = cell?.value !== undefined && cell.value !== null && cell.value !== '';
        if (!hasValue && !cell?.formula) continue;

        minRow = Math.min(minRow, row);
        minCol = Math.min(minCol, col);
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      }
    }

    if (maxRow === 0 || maxCol === 0) {
      return {
        start: { row: this._activeRow, col: this._activeCol },
        end: { row: this._activeRow, col: this._activeCol },
      };
    }

    return {
      start: { row: minRow, col: minCol },
      end: { row: maxRow, col: maxCol },
    };
  }
}

export function createKeyboardManager(
  sdk: SpreadsheetSDK,
  options?: KeyboardShortcutManagerOptions,
): KeyboardShortcutManager {
  return new KeyboardShortcutManagerImpl(sdk, options ?? {});
}
