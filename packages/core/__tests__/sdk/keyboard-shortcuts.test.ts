/**
 * @group sdk
 *
 * Phase 22: Keyboard Shortcuts — SDK Test Suite
 *
 *  §1   normaliseCombo() — string & KeyLike parsing, modifier ordering
 *  §2   Construction — initial state, custom position, bounds clamping
 *  §3   rowCount / colCount on SDK
 *  §4   Navigation — Arrow keys (basic move, bounds clamping, clears anchor)
 *  §5   Navigation — Tab / Shift+Tab (wrap around rows)
 *  §6   Navigation — Enter / Shift+Enter
 *  §7   Navigation — Home / End / Ctrl+Home / Ctrl+End
 *  §8   Navigation — PageUp / PageDown (clamped, custom pageSize)
 *  §9   Navigation — Ctrl+Arrow (data-edge jump)
 *  §10  Selection — Shift+Arrow (sets anchor, extends, selectionRange)
 *  §11  Selection — Ctrl+Shift+Arrow (extend to data edge)
 *  §12  Selection — Shift+Home / Ctrl+Shift+Home
 *  §13  Selection — Ctrl+A / Ctrl+Space / Shift+Space
 *  §14  Escape — clears anchor; no-op when no anchor/no callback
 *  §15  Edit — Delete / Backspace (clear active cell)
 *  §16  Edit — Ctrl+; (today's date ISO)
 *  §17  Edit — Ctrl+D (fill down) / Ctrl+R (fill right)
 *  §18  Undo / Redo
 *  §19  Custom bind() / unbind() / getBoundCombos()
 *  §20  resetToDefaults()
 *  §21  Format callbacks (Ctrl+B/I/U/1)
 *  §22  Dialog callbacks (Ctrl+F / Ctrl+H / Ctrl+G / F5 / F2)
 *  §23  AutoFilter toggle (Ctrl+Shift+L)
 *  §24  handleKeyEvent return values
 *  §25  setActiveCell (clamp, keepAnchor)
 *  §26  Dispose safety
 */

import { createSpreadsheet } from '../../src/sdk/index';
import {
  createKeyboardManager,
  normaliseCombo,
} from '../../src/sdk/KeyboardShortcutManager';
import type {
  KeyboardShortcutManager,
  KeyboardShortcutManagerOptions,
  KeyLike,
} from '../../src/sdk/KeyboardShortcutManager';
import type { SpreadsheetSDK } from '../../src/sdk/SpreadsheetSDK';

// ── helpers ───────────────────────────────────────────────────────────────

function makeSheet(rows = 20, cols = 10): SpreadsheetSDK {
  return createSpreadsheet('KBD', { rows, cols });
}

function key(k: string, mods: Partial<Omit<KeyLike, 'key'>> = {}): KeyLike {
  return { key: k, ...mods };
}

const ctrl  = (k: string) => key(k, { ctrlKey: true });
const shift = (k: string) => key(k, { shiftKey: true });
const ctrlShift = (k: string) => key(k, { ctrlKey: true, shiftKey: true });

// ── §1 normaliseCombo() ───────────────────────────────────────────────────

describe('§1 normaliseCombo', () => {
  test('plain key lowercased', () => {
    expect(normaliseCombo('ArrowLeft')).toBe('arrowleft');
  });

  test('single modifier', () => {
    expect(normaliseCombo('ctrl+z')).toBe('ctrl+z');
    expect(normaliseCombo('Ctrl+Z')).toBe('ctrl+z');
  });

  test('modifiers reordered to canonical ctrl+meta+shift+alt+key', () => {
    expect(normaliseCombo('shift+ctrl+z')).toBe('ctrl+shift+z');
    expect(normaliseCombo('alt+ctrl+shift+f')).toBe('ctrl+shift+alt+f');
  });

  test('KeyLike object — plain key', () => {
    expect(normaliseCombo({ key: 'ArrowRight' })).toBe('arrowright');
  });

  test('KeyLike object — ctrl+shift', () => {
    expect(normaliseCombo({ key: 'Home', ctrlKey: true, shiftKey: true })).toBe('ctrl+shift+home');
  });

  test('KeyLike object — all modifiers', () => {
    expect(normaliseCombo({ key: 'X', ctrlKey: true, metaKey: true, shiftKey: true, altKey: true }))
      .toBe('ctrl+meta+shift+alt+x');
  });

  test('cmd alias becomes meta', () => {
    expect(normaliseCombo('cmd+z')).toBe('meta+z');
  });

  test('F-key preserved', () => {
    expect(normaliseCombo('F2')).toBe('f2');
    expect(normaliseCombo('f5')).toBe('f5');
  });
});

// ── §2 Construction ────────────────────────────────────────────────────────

describe('§2 Construction', () => {
  test('default cursor is (1,1)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.activeRow).toBe(1);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('no anchor on fresh manager', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.anchorRow).toBeNull();
    expect(m.anchorCol).toBeNull();
    expect(m.selectionRange).toBeNull();
    s.dispose();
  });

  test('custom initialRow / initialCol', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 3 });
    expect(m.activeRow).toBe(5);
    expect(m.activeCol).toBe(3);
    s.dispose();
  });

  test('initialRow / initialCol clamped to sheet bounds', () => {
    const s = makeSheet(10, 5);
    const m = createKeyboardManager(s, { initialRow: 99, initialCol: 99 });
    expect(m.activeRow).toBe(10);
    expect(m.activeCol).toBe(5);
    s.dispose();
  });

  test('initialRow / initialCol clamped to at least 1', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 0, initialCol: -3 });
    expect(m.activeRow).toBe(1);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });
});

// ── §3 rowCount / colCount on SDK ─────────────────────────────────────────

describe('§3 rowCount / colCount', () => {
  test('sheet created with rows=15 cols=8', () => {
    const s = createSpreadsheet('RC', { rows: 15, cols: 8 });
    expect(s.rowCount).toBe(15);
    expect(s.colCount).toBe(8);
    s.dispose();
  });

  test('default sheet size', () => {
    const s = makeSheet();
    expect(s.rowCount).toBeGreaterThan(0);
    expect(s.colCount).toBeGreaterThan(0);
    s.dispose();
  });
});

// ── §4 Navigation — Arrow keys ────────────────────────────────────────────

describe('§4 Arrow navigation', () => {
  test('ArrowRight moves +1 col', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(key('ArrowRight'));
    expect(m.activeRow).toBe(3);
    expect(m.activeCol).toBe(4);
    s.dispose();
  });

  test('ArrowLeft moves -1 col', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 5 });
    m.handleKeyEvent(key('ArrowLeft'));
    expect(m.activeCol).toBe(4);
    s.dispose();
  });

  test('ArrowUp moves -1 row', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 4, initialCol: 2 });
    m.handleKeyEvent(key('ArrowUp'));
    expect(m.activeRow).toBe(3);
    s.dispose();
  });

  test('ArrowDown moves +1 row', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 4, initialCol: 2 });
    m.handleKeyEvent(key('ArrowDown'));
    expect(m.activeRow).toBe(5);
    s.dispose();
  });

  test('ArrowLeft clamped at col 1', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialCol: 1 });
    m.handleKeyEvent(key('ArrowLeft'));
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('ArrowUp clamped at row 1', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 1 });
    m.handleKeyEvent(key('ArrowUp'));
    expect(m.activeRow).toBe(1);
    s.dispose();
  });

  test('Arrow clears anchor', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    // establish anchor via Shift+Arrow
    m.handleKeyEvent(shift('ArrowDown'));
    expect(m.anchorRow).not.toBeNull();
    m.handleKeyEvent(key('ArrowRight'));
    expect(m.anchorRow).toBeNull();
    s.dispose();
  });
});

// ── §5 Navigation — Tab / Shift+Tab ─────────────────────────────────────

describe('§5 Tab / Shift+Tab', () => {
  test('Tab moves +1 col within row', () => {
    const s = makeSheet(10, 10);
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3 });
    m.handleKeyEvent(key('Tab'));
    expect(m.activeRow).toBe(2);
    expect(m.activeCol).toBe(4);
    s.dispose();
  });

  test('Tab wraps to col 1 of next row at last col', () => {
    const s = makeSheet(10, 10);
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 10 });
    m.handleKeyEvent(key('Tab'));
    expect(m.activeRow).toBe(3);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('Shift+Tab moves -1 col', () => {
    const s = makeSheet(10, 10);
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 5 });
    m.handleKeyEvent(shift('Tab'));
    expect(m.activeRow).toBe(2);
    expect(m.activeCol).toBe(4);
    s.dispose();
  });

  test('Shift+Tab wraps to last col of previous row at col 1', () => {
    const s = makeSheet(10, 10);
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 1 });
    m.handleKeyEvent(shift('Tab'));
    expect(m.activeRow).toBe(2);
    expect(m.activeCol).toBe(10);
    s.dispose();
  });
});

// ── §6 Navigation — Enter / Shift+Enter ──────────────────────────────────

describe('§6 Enter / Shift+Enter', () => {
  test('Enter moves +1 row', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 2 });
    m.handleKeyEvent(key('Enter'));
    expect(m.activeRow).toBe(4);
    expect(m.activeCol).toBe(2);
    s.dispose();
  });

  test('Shift+Enter moves -1 row', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 2 });
    m.handleKeyEvent(shift('Enter'));
    expect(m.activeRow).toBe(2);
    s.dispose();
  });
});

// ── §7 Navigation — Home / End / Ctrl+Home / Ctrl+End ────────────────────

describe('§7 Home / End / Ctrl+Home / Ctrl+End', () => {
  test('Home goes to col 1 of current row', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 7 });
    m.handleKeyEvent(key('Home'));
    expect(m.activeRow).toBe(5);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('End goes to last col', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 1 });
    m.handleKeyEvent(key('End'));
    expect(m.activeRow).toBe(5);
    expect(m.activeCol).toBe(10);
    s.dispose();
  });

  test('Ctrl+Home goes to (1,1)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 10, initialCol: 8 });
    m.handleKeyEvent(ctrl('Home'));
    expect(m.activeRow).toBe(1);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('Ctrl+End goes to last used cell', () => {
    const s = makeSheet(20, 10);
    s.setCell(5, 7, 'hello');
    s.setCell(3, 2, 'world');
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    m.handleKeyEvent(ctrl('End'));
    // Last used cell should be (5, 7)
    expect(m.activeRow).toBe(5);
    expect(m.activeCol).toBe(7);
    s.dispose();
  });

  test('Ctrl+End on blank sheet stays at (1,1)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.handleKeyEvent(ctrl('End'));
    expect(m.activeRow).toBe(1);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });
});

// ── §8 Navigation — PageUp / PageDown ─────────────────────────────────────

describe('§8 PageUp / PageDown', () => {
  test('PageDown moves +pageSize rows (default 10)', () => {
    const s = makeSheet(50, 10);
    const m = createKeyboardManager(s, { initialRow: 5 });
    m.handleKeyEvent(key('PageDown'));
    expect(m.activeRow).toBe(15);
    s.dispose();
  });

  test('PageUp moves -pageSize rows (default 10)', () => {
    const s = makeSheet(50, 10);
    const m = createKeyboardManager(s, { initialRow: 20 });
    m.handleKeyEvent(key('PageUp'));
    expect(m.activeRow).toBe(10);
    s.dispose();
  });

  test('PageDown clamped at rowCount', () => {
    const s = makeSheet(15, 10);
    const m = createKeyboardManager(s, { initialRow: 12 });
    m.handleKeyEvent(key('PageDown'));
    expect(m.activeRow).toBe(15);
    s.dispose();
  });

  test('PageUp clamped at row 1', () => {
    const s = makeSheet(50, 10);
    const m = createKeyboardManager(s, { initialRow: 3 });
    m.handleKeyEvent(key('PageUp'));
    expect(m.activeRow).toBe(1);
    s.dispose();
  });

  test('custom pageSize respected', () => {
    const s = makeSheet(50, 10);
    const m = createKeyboardManager(s, { initialRow: 5, pageSize: 5 });
    m.handleKeyEvent(key('PageDown'));
    expect(m.activeRow).toBe(10);
    s.dispose();
  });
});

// ── §9 Navigation — Ctrl+Arrow ────────────────────────────────────────────

describe('§9 Ctrl+Arrow', () => {
  test('Ctrl+ArrowRight from filled cell jumps to last consecutive filled', () => {
    const s = makeSheet(20, 10);
    s.setCell(3, 2, 'a');
    s.setCell(3, 3, 'b');
    s.setCell(3, 4, 'c');
    // col 5 is empty
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 2 });
    m.handleKeyEvent(ctrl('ArrowRight'));
    expect(m.activeRow).toBe(3);
    expect(m.activeCol).toBe(4);
    s.dispose();
  });

  test('Ctrl+ArrowRight from empty cell jumps to first non-empty or boundary', () => {
    const s = makeSheet(20, 10);
    s.setCell(3, 5, 'data');
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 2 });
    m.handleKeyEvent(ctrl('ArrowRight'));
    expect(m.activeRow).toBe(3);
    expect(m.activeCol).toBe(5);
    s.dispose();
  });

  test('Ctrl+ArrowDown from empty column jumps to last row', () => {
    const s = makeSheet(10, 5);
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    m.handleKeyEvent(ctrl('ArrowDown'));
    expect(m.activeRow).toBe(10);
    s.dispose();
  });
});

// ── §10 Selection — Shift+Arrow ───────────────────────────────────────────

describe('§10 Shift+Arrow', () => {
  test('first Shift+ArrowDown sets anchor then extends down', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(shift('ArrowDown'));
    expect(m.anchorRow).toBe(3);
    expect(m.anchorCol).toBe(3);
    expect(m.activeRow).toBe(4);
    expect(m.activeCol).toBe(3);
    s.dispose();
  });

  test('selectionRange bounding box is correct', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(shift('ArrowRight'));
    const r = m.selectionRange!;
    expect(r.start).toEqual({ row: 3, col: 3 });
    expect(r.end).toEqual({ row: 4, col: 4 });
    s.dispose();
  });

  test('extending above anchor swaps bounding box', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 5 });
    m.handleKeyEvent(shift('ArrowUp'));
    m.handleKeyEvent(shift('ArrowUp'));
    const r = m.selectionRange!;
    expect(r.start.row).toBe(3);
    expect(r.end.row).toBe(5);
    s.dispose();
  });
});

// ── §11 Selection — Ctrl+Shift+Arrow ─────────────────────────────────────

describe('§11 Ctrl+Shift+Arrow', () => {
  test('Ctrl+Shift+ArrowRight extends to data edge', () => {
    const s = makeSheet(20, 10);
    s.setCell(2, 3, 'a');
    s.setCell(2, 4, 'b');
    s.setCell(2, 5, 'c');
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3 });
    m.handleKeyEvent(ctrlShift('ArrowRight'));
    expect(m.anchorRow).toBe(2);
    expect(m.activeCol).toBe(5);
    s.dispose();
  });
});

// ── §12 Selection — Shift+Home / Ctrl+Shift+Home ─────────────────────────

describe('§12 Shift+Home / Ctrl+Shift+Home', () => {
  test('Shift+Home extends selection to col 1', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 4, initialCol: 6 });
    m.handleKeyEvent(shift('Home'));
    expect(m.anchorRow).toBe(4);
    expect(m.anchorCol).toBe(6);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });

  test('Ctrl+Shift+Home extends selection to (1,1)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 5 });
    m.handleKeyEvent(ctrlShift('Home'));
    expect(m.anchorRow).toBe(5);
    expect(m.anchorCol).toBe(5);
    expect(m.activeRow).toBe(1);
    expect(m.activeCol).toBe(1);
    s.dispose();
  });
});

// ── §13 Selection — Ctrl+A / Ctrl+Space / Shift+Space ────────────────────

describe('§13 Ctrl+A / Ctrl+Space / Shift+Space', () => {
  test('Ctrl+A selects the used data range', () => {
    const s = makeSheet(20, 10);
    s.setCell(3, 2, 'northwest');
    s.setCell(8, 6, 'southeast');
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 5 });
    m.handleKeyEvent(ctrl('a'));
    const r = m.selectionRange!;
    expect(r.start).toEqual({ row: 3, col: 2 });
    expect(r.end).toEqual({ row: 8, col: 6 });
    s.dispose();
  });

  test('Ctrl+A selects the active cell when the sheet is empty', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 5 });
    m.handleKeyEvent(ctrl('a'));
    const r = m.selectionRange!;
    expect(r.start).toEqual({ row: 5, col: 5 });
    expect(r.end).toEqual({ row: 5, col: 5 });
    s.dispose();
  });

  test('Ctrl+Space selects entire column', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 5, initialCol: 3 });
    m.handleKeyEvent(ctrl(' '));
    const r = m.selectionRange!;
    expect(r.start).toEqual({ row: 1, col: 3 });
    expect(r.end).toEqual({ row: 20, col: 3 });
    s.dispose();
  });

  test('Shift+Space selects entire row', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 7, initialCol: 4 });
    m.handleKeyEvent(shift(' '));
    const r = m.selectionRange!;
    expect(r.start).toEqual({ row: 7, col: 1 });
    expect(r.end).toEqual({ row: 7, col: 10 });
    s.dispose();
  });
});

// ── §14 Escape ────────────────────────────────────────────────────────────

describe('§14 Escape', () => {
  test('Escape clears anchor when set', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(shift('ArrowDown'));
    expect(m.anchorRow).not.toBeNull();
    const result = m.handleKeyEvent(key('Escape'));
    expect(result).toBe(true);
    expect(m.anchorRow).toBeNull();
    s.dispose();
  });

  test('Escape returns false when no anchor and no callback', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(key('Escape'))).toBe(false);
    s.dispose();
  });

  test('Escape delegates to onEscapeRequest callback', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onEscapeRequest: cb });
    const result = m.handleKeyEvent(key('Escape'));
    expect(cb).toHaveBeenCalledWith(1, 1);
    expect(result).toBe(true);
    s.dispose();
  });
});

// ── §15 Edit — Delete / Backspace ─────────────────────────────────────────

describe('§15 Delete / Backspace clear active cell', () => {
  test('Delete clears active cell value', () => {
    const s = makeSheet();
    s.setCell(2, 3, 'hello');
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3 });
    m.handleKeyEvent(key('Delete'));
    expect(s.getCellValue(2, 3)).toBeNull();
    s.dispose();
  });

  test('Backspace clears active cell value', () => {
    const s = makeSheet();
    s.setCell(4, 5, 42);
    const m = createKeyboardManager(s, { initialRow: 4, initialCol: 5 });
    m.handleKeyEvent(key('Backspace'));
    expect(s.getCellValue(4, 5)).toBeNull();
    s.dispose();
  });

  test('Delete on already-empty cell is a no-op (returns true)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    expect(m.handleKeyEvent(key('Delete'))).toBe(true);
    s.dispose();
  });
});

// ── §16 Edit — Ctrl+; ────────────────────────────────────────────────────

describe('§16 Ctrl+; inserts today\'s date', () => {
  test('inserts a YYYY-MM-DD string', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    m.handleKeyEvent(ctrl(';'));
    const val = s.getCellValue(1, 1) as string;
    expect(val).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    s.dispose();
  });
});

// ── §17 Edit — Ctrl+D / Ctrl+R ───────────────────────────────────────────

describe('§17 Ctrl+D fill-down / Ctrl+R fill-right', () => {
  test('Ctrl+D fills top row\'s values down the selection', () => {
    const s = makeSheet(20, 10);
    s.setCell(2, 3, 'source');
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3 });
    // extend selection down to row 4
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(ctrl('d'));
    expect(s.getCellValue(3, 3)).toBe('source');
    expect(s.getCellValue(4, 3)).toBe('source');
    s.dispose();
  });

  test('Ctrl+D returns false with single-row selection', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    expect(m.handleKeyEvent(ctrl('d'))).toBe(false);
    s.dispose();
  });

  test('Ctrl+R fills left column\'s values right across the selection', () => {
    const s = makeSheet(20, 10);
    s.setCell(3, 2, 'src');
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 2 });
    m.handleKeyEvent(shift('ArrowRight'));
    m.handleKeyEvent(shift('ArrowRight'));
    m.handleKeyEvent(ctrl('r'));
    expect(s.getCellValue(3, 3)).toBe('src');
    expect(s.getCellValue(3, 4)).toBe('src');
    s.dispose();
  });

  test('Ctrl+R returns false with single-col selection', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    expect(m.handleKeyEvent(ctrl('r'))).toBe(false);
    s.dispose();
  });
});

// ── §18 Undo / Redo ───────────────────────────────────────────────────────

describe('§18 Undo / Redo', () => {
  test('Ctrl+Z calls sdk.undo()', () => {
    const s = makeSheet();
    s.setCell(1, 1, 'val');
    const m = createKeyboardManager(s);
    m.handleKeyEvent(ctrl('z'));
    expect(s.getCellValue(1, 1)).toBeNull();
    s.dispose();
  });

  test('Ctrl+Y calls sdk.redo()', () => {
    const s = makeSheet();
    s.setCell(1, 1, 'val');
    const m = createKeyboardManager(s);
    m.handleKeyEvent(ctrl('z'));   // undo
    m.handleKeyEvent(ctrl('y'));   // redo
    expect(s.getCellValue(1, 1)).toBe('val');
    s.dispose();
  });

  test('Ctrl+Shift+Z is an alias for redo', () => {
    const s = makeSheet();
    s.setCell(1, 1, 'x');
    const m = createKeyboardManager(s);
    m.handleKeyEvent(ctrl('z'));
    m.handleKeyEvent(ctrlShift('z'));
    expect(s.getCellValue(1, 1)).toBe('x');
    s.dispose();
  });

  test('F4 is also redo', () => {
    const s = makeSheet();
    s.setCell(1, 1, 99);
    const m = createKeyboardManager(s);
    m.handleKeyEvent(ctrl('z'));
    m.handleKeyEvent(key('F4'));
    expect(s.getCellValue(1, 1)).toBe(99);
    s.dispose();
  });

  test('Ctrl+Z on empty stack returns false', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(ctrl('z'))).toBe(false);
    s.dispose();
  });
});

// ── §19 Custom bind() / unbind() / getBoundCombos() ──────────────────────

describe('§19 bind / unbind / getBoundCombos', () => {
  test('bind overrides existing combo', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    const spy = jest.fn(() => true as const);
    m.bind('arrowleft', spy);
    m.handleKeyEvent(key('ArrowLeft'));
    expect(spy).toHaveBeenCalledTimes(1);
    s.dispose();
  });

  test('bind registers new combo', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    const spy = jest.fn(() => true as const);
    m.bind('ctrl+alt+x', spy);
    m.handleKeyEvent(key('x', { ctrlKey: true, altKey: true }));
    expect(spy).toHaveBeenCalledTimes(1);
    s.dispose();
  });

  test('bind accepts array of combos', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    const spy = jest.fn(() => true as const);
    m.bind(['ctrl+1', 'ctrl+2'], spy);
    m.handleKeyEvent(ctrl('1'));
    m.handleKeyEvent(ctrl('2'));
    expect(spy).toHaveBeenCalledTimes(2);
    s.dispose();
  });

  test('unbind removes handler, key returns false', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.unbind('arrowleft');
    expect(m.handleKeyEvent(key('ArrowLeft'))).toBe(false);
    s.dispose();
  });

  test('getBoundCombos returns sorted list including custom combo', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.bind('ctrl+alt+q', () => true);
    const combos = m.getBoundCombos();
    expect(combos).toContain('ctrl+alt+q');
    expect(combos).toEqual([...combos].sort());
    s.dispose();
  });
});

// ── §20 resetToDefaults() ─────────────────────────────────────────────────

describe('§20 resetToDefaults', () => {
  test('discards custom bindings and restores defaults', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.bind('ctrl+alt+q', () => true);
    m.resetToDefaults();
    expect(m.handleKeyEvent(key('x', { ctrlKey: true, altKey: true }))).toBe(false);
    // default ArrowLeft still works
    const m2 = createKeyboardManager(s, { initialCol: 5 });
    m2.resetToDefaults();
    m2.handleKeyEvent(key('ArrowLeft'));
    expect(m2.activeCol).toBe(4);
    s.dispose();
  });
});

// ── §21 Format callbacks ───────────────────────────────────────────────────

describe('§21 Format callbacks', () => {
  test('Ctrl+B fires onFormatRequest with "bold"', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3, onFormatRequest: cb });
    m.handleKeyEvent(ctrl('b'));
    expect(cb).toHaveBeenCalledWith('bold', 2, 3);
    s.dispose();
  });

  test('Ctrl+I fires "italic"', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onFormatRequest: cb });
    m.handleKeyEvent(ctrl('i'));
    expect(cb).toHaveBeenCalledWith('italic', 1, 1);
    s.dispose();
  });

  test('Ctrl+U fires "underline"', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onFormatRequest: cb });
    m.handleKeyEvent(ctrl('u'));
    expect(cb).toHaveBeenCalledWith('underline', 1, 1);
    s.dispose();
  });

  test('Ctrl+1 fires "format-cells"', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onFormatRequest: cb });
    m.handleKeyEvent(ctrl('1'));
    expect(cb).toHaveBeenCalledWith('format-cells', 1, 1);
    s.dispose();
  });

  test('format shortcut returns false when no callback', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(ctrl('b'))).toBe(false);
    s.dispose();
  });
});

// ── §22 Dialog callbacks ───────────────────────────────────────────────────

describe('§22 Dialog callbacks', () => {
  test('Ctrl+F fires onFindRequest', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onFindRequest: cb });
    const result = m.handleKeyEvent(ctrl('f'));
    expect(cb).toHaveBeenCalledTimes(1);
    expect(result).toBe(true);
    s.dispose();
  });

  test('Ctrl+F returns false when no callback', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(ctrl('f'))).toBe(false);
    s.dispose();
  });

  test('Ctrl+H fires onReplaceRequest', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onReplaceRequest: cb });
    m.handleKeyEvent(ctrl('h'));
    expect(cb).toHaveBeenCalledTimes(1);
    s.dispose();
  });

  test('Ctrl+G fires onGoToRequest', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onGoToRequest: cb });
    m.handleKeyEvent(ctrl('g'));
    expect(cb).toHaveBeenCalledTimes(1);
    s.dispose();
  });

  test('F5 also fires onGoToRequest', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { onGoToRequest: cb });
    m.handleKeyEvent(key('F5'));
    expect(cb).toHaveBeenCalledTimes(1);
    s.dispose();
  });

  test('F2 fires onEditRequest with current position', () => {
    const s = makeSheet();
    const cb = jest.fn(() => true as const);
    const m = createKeyboardManager(s, { initialRow: 4, initialCol: 6, onEditRequest: cb });
    m.handleKeyEvent(key('F2'));
    expect(cb).toHaveBeenCalledWith(4, 6);
    s.dispose();
  });

  test('F2 returns false when no callback', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(key('F2'))).toBe(false);
    s.dispose();
  });
});

// ── §23 AutoFilter toggle ─────────────────────────────────────────────────

describe('§23 AutoFilter toggle (Ctrl+Shift+L)', () => {
  test('enables autofilter when none active', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    m.handleKeyEvent(ctrlShift('l'));
    const range = s.getAutoFilterRange();
    expect(range).toEqual({ headerRow: 1, startCol: 1, endCol: 10 });
    s.dispose();
  });

  test('enables autofilter for selected column range', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 2 });
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(ctrlShift('l'));
    expect(s.getAutoFilterRange()).toEqual({ headerRow: 1, startCol: 2, endCol: 2 });
    s.dispose();
  });

  test('enables autofilter for selected multi-column range', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 2 });
    m.handleKeyEvent(shift('ArrowRight'));
    m.handleKeyEvent(shift('ArrowDown'));
    m.handleKeyEvent(ctrlShift('l'));
    expect(s.getAutoFilterRange()).toEqual({ headerRow: 1, startCol: 2, endCol: 3 });
    s.dispose();
  });

  test('enables autofilter for whole-sheet selection', () => {
    const s = makeSheet(20, 10);
    s.setCell(1, 1, 'first');
    s.setCell(20, 10, 'last');
    const m = createKeyboardManager(s, { initialRow: 2, initialCol: 3 });
    m.handleKeyEvent(ctrl('a'));
    m.handleKeyEvent(ctrlShift('l'));
    expect(s.getAutoFilterRange()).toEqual({ headerRow: 1, startCol: 1, endCol: 10 });
    s.dispose();
  });

  test('disables autofilter when already active', () => {
    const s = makeSheet(20, 10);
    s.setAutoFilterRange(1, 1, 10);
    const m = createKeyboardManager(s, { initialRow: 1, initialCol: 1 });
    m.handleKeyEvent(ctrlShift('l'));
    expect(s.getAutoFilterRange()).toBeNull();
    s.dispose();
  });
});

// ── §24 handleKeyEvent return values ──────────────────────────────────────

describe('§24 handleKeyEvent return values', () => {
  test('unrecognised key returns false', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(key('F99'))).toBe(false);
    expect(m.handleKeyEvent(key('q', { ctrlKey: true, altKey: true }))).toBe(false);
    s.dispose();
  });

  test('recognised key returns true', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    expect(m.handleKeyEvent(key('ArrowRight'))).toBe(true);
    s.dispose();
  });
});

// ── §25 setActiveCell ─────────────────────────────────────────────────────

describe('§25 setActiveCell', () => {
  test('moves cursor directly', () => {
    const s = makeSheet(20, 10);
    const m = createKeyboardManager(s);
    m.setActiveCell(7, 5);
    expect(m.activeRow).toBe(7);
    expect(m.activeCol).toBe(5);
    s.dispose();
  });

  test('clamps to sheet bounds', () => {
    const s = makeSheet(10, 5);
    const m = createKeyboardManager(s);
    m.setActiveCell(100, 100);
    expect(m.activeRow).toBe(10);
    expect(m.activeCol).toBe(5);
    s.dispose();
  });

  test('clears anchor by default', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(shift('ArrowDown'));
    expect(m.anchorRow).not.toBeNull();
    m.setActiveCell(5, 5);
    expect(m.anchorRow).toBeNull();
    s.dispose();
  });

  test('keepAnchor=true preserves existing anchor', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s, { initialRow: 3, initialCol: 3 });
    m.handleKeyEvent(shift('ArrowDown'));
    const savedAnchorRow = m.anchorRow;
    m.setActiveCell(6, 6, true);
    expect(m.anchorRow).toBe(savedAnchorRow);
    s.dispose();
  });

  test('onSelectionChange fired with new position', () => {
    const s = makeSheet(20, 10);
    const cb = jest.fn();
    const m = createKeyboardManager(s, { onSelectionChange: cb });
    m.setActiveCell(8, 4);
    expect(cb).toHaveBeenCalledWith(8, 4, null, null);
    s.dispose();
  });
});

// ── §26 Dispose safety ────────────────────────────────────────────────────

describe('§26 Dispose safety', () => {
  test('handleKeyEvent throws after dispose', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.dispose();
    expect(() => m.handleKeyEvent(key('ArrowLeft'))).toThrow(/dispose/i);
    s.dispose();
  });

  test('bind throws after dispose', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.dispose();
    expect(() => m.bind('ctrl+z', () => true)).toThrow(/dispose/i);
    s.dispose();
  });

  test('getBoundCombos throws after dispose', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.dispose();
    expect(() => m.getBoundCombos()).toThrow(/dispose/i);
    s.dispose();
  });

  test('double dispose is safe (no throw)', () => {
    const s = makeSheet();
    const m = createKeyboardManager(s);
    m.dispose();
    expect(() => m.dispose()).not.toThrow();
    s.dispose();
  });
});
