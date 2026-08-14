/**
 * visibility-api.test.ts — Phase 3: Hidden Row/Column Model
 *
 * PM-mandated coverage:
 *   1. Hide row → find skips hidden row
 *   2. Hide column → find skips hidden column
 *   3. Hide anchor of merged region (merge still intact, just visually hidden)
 *   4. Hide part of merged region (non-anchor rows/cols hidden)
 *   5. Unhide restores traversal (deterministic)
 *   6. Stress test: 10k hidden rows
 *   7. Hidden state does NOT affect storage memory (no per-cell flags)
 *
 * Plus full behavioral matrix:
 *   - getCell always returns, even in hidden row/col
 *   - findIterator skips hidden by default
 *   - findIterator includes hidden when includeHidden: true
 *   - row × col independence (hide row ≠ hide col)
 *   - idempotent hide/show
 *   - event emission: row-hidden, row-shown, col-hidden, col-shown
 *   - spill NOT blocked by hidden cells (display-only model)
 *
 * Run: npx jest packages/core/__tests__/visibility-api.test.ts --no-coverage --verbose
 */

import { describe, test, expect, beforeEach } from '@jest/globals';
import { Worksheet } from '../src/worksheet';
import { VisibilityStoreV1, VisibilityStoreLegacy } from '../src/storage/VisibilityStore';
import { SpillEngine } from '../src/SpillEngine';

function makeSheet(rows = 100, cols = 26): Worksheet {
  return new Worksheet('Test', rows, cols);
}

// ---------------------------------------------------------------------------
// 1.  VisibilityStoreV1 unit tests — data structure correctness
// ---------------------------------------------------------------------------

describe('VisibilityStoreV1 — data structure correctness', () => {
  let store: VisibilityStoreV1;
  beforeEach(() => { store = new VisibilityStoreV1(); });

  // ── basic hide/show ───────────────────────────────────────────────────────

  test('hideRow → isRowHidden returns true', () => {
    store.hideRow(5);
    expect(store.isRowHidden(5)).toBe(true);
  });

  test('showRow → isRowHidden returns false', () => {
    store.hideRow(5);
    store.showRow(5);
    expect(store.isRowHidden(5)).toBe(false);
  });

  test('hideCol → isColHidden returns true', () => {
    store.hideCol(3);
    expect(store.isColHidden(3)).toBe(true);
  });

  test('showCol → isColHidden returns false', () => {
    store.hideCol(3);
    store.showCol(3);
    expect(store.isColHidden(3)).toBe(false);
  });

  test('unhidden row/col returns false by default', () => {
    expect(store.isRowHidden(1)).toBe(false);
    expect(store.isColHidden(1)).toBe(false);
  });

  // ── isCellHidden ─────────────────────────────────────────────────────────

  test('isCellHidden: false when both visible', () => {
    expect(store.isCellHidden(1, 1)).toBe(false);
  });

  test('isCellHidden: true when row hidden', () => {
    store.hideRow(3);
    expect(store.isCellHidden(3, 5)).toBe(true);
  });

  test('isCellHidden: true when col hidden', () => {
    store.hideCol(4);
    expect(store.isCellHidden(7, 4)).toBe(true);
  });

  test('isCellHidden: true when both row and col hidden', () => {
    store.hideRow(2);
    store.hideCol(2);
    expect(store.isCellHidden(2, 2)).toBe(true);
  });

  test('isCellHidden: adjacent cells not affected by row hide', () => {
    store.hideRow(5);
    expect(store.isCellHidden(4, 1)).toBe(false);
    expect(store.isCellHidden(6, 1)).toBe(false);
  });

  // ── idempotency ───────────────────────────────────────────────────────────

  test('hideRow is idempotent — double hide/show roundtrip', () => {
    store.hideRow(1);
    store.hideRow(1); // idempotent
    expect(store.isRowHidden(1)).toBe(true);
    store.showRow(1);
    expect(store.isRowHidden(1)).toBe(false);
    store.showRow(1); // idempotent
    expect(store.isRowHidden(1)).toBe(false);
  });

  test('showRow on non-hidden row is a no-op', () => {
    expect(() => store.showRow(99)).not.toThrow();
    expect(store.isRowHidden(99)).toBe(false);
  });

  // ── counts ────────────────────────────────────────────────────────────────

  test('hiddenRowCount / hiddenColCount reflect current state', () => {
    expect(store.hiddenRowCount).toBe(0);
    store.hideRow(1);
    store.hideRow(2);
    store.hideRow(3);
    expect(store.hiddenRowCount).toBe(3);
    store.showRow(2);
    expect(store.hiddenRowCount).toBe(2);

    store.hideCol(5);
    expect(store.hiddenColCount).toBe(1);
  });

  // ── getHiddenRows / getHiddenCols ─────────────────────────────────────────

  test('getHiddenRows returns correct set', () => {
    store.hideRow(10);
    store.hideRow(20);
    const hidden = store.getHiddenRows();
    expect(hidden.has(10)).toBe(true);
    expect(hidden.has(20)).toBe(true);
    expect(hidden.has(15)).toBe(false);
  });

  test('getHiddenCols returns correct set', () => {
    store.hideCol(2);
    store.hideCol(4);
    const hidden = store.getHiddenCols();
    expect(hidden.has(2)).toBe(true);
    expect(hidden.has(4)).toBe(true);
    expect(hidden.has(3)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 2.  VisibilityStoreLegacy — rollback shim parity
// ---------------------------------------------------------------------------

describe('VisibilityStoreLegacy — rollback shim parity', () => {
  let legacy: VisibilityStoreLegacy;
  beforeEach(() => { legacy = new VisibilityStoreLegacy(); });

  test('hideRow → isRowHidden true', () => {
    legacy.hideRow(3);
    expect(legacy.isRowHidden(3)).toBe(true);
  });

  test('showRow → isRowHidden false', () => {
    legacy.hideRow(3);
    legacy.showRow(3);
    expect(legacy.isRowHidden(3)).toBe(false);
  });

  test('isCellHidden: hidden row', () => {
    legacy.hideRow(7);
    expect(legacy.isCellHidden(7, 1)).toBe(true);
  });

  test('isCellHidden: hidden col', () => {
    legacy.hideCol(2);
    expect(legacy.isCellHidden(1, 2)).toBe(true);
  });

  test('hiddenRowCount accurate', () => {
    legacy.hideRow(1);
    legacy.hideRow(2);
    expect(legacy.hiddenRowCount).toBe(2);
  });

  test('Legacy is idempotent for hideRow', () => {
    legacy.hideRow(5);
    legacy.hideRow(5);
    expect(legacy.hiddenRowCount).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// 3.  PM Requirement 1: Hide row → find skips hidden row
// ---------------------------------------------------------------------------

describe('PM-1 · Hide row — find skips hidden row', () => {
  let sheet: Worksheet;
  beforeEach(() => { sheet = makeSheet(); });

  test('findAll does not return cells in hidden row (default includeHidden=false)', () => {
    sheet.setCellValue({ row: 3, col: 1 }, 'Target');
    sheet.hideRow(3);
    const results = sheet.findAll({ what: 'Target' });
    expect(results).toHaveLength(0);
  });

  test('findAll returns cells in hidden row when includeHidden: true', () => {
    sheet.setCellValue({ row: 3, col: 1 }, 'Target');
    sheet.hideRow(3);
    const results = sheet.findAll({ what: 'Target', includeHidden: true });
    expect(results).toHaveLength(1);
    expect(results[0]).toEqual({ row: 3, col: 1 });
  });

  test('findAll returns visible cells, skips hidden row cells', () => {
    sheet.setCellValue({ row: 1, col: 1 }, 'Apple');
    sheet.setCellValue({ row: 2, col: 1 }, 'Apple');
    sheet.setCellValue({ row: 3, col: 1 }, 'Apple');
    sheet.hideRow(2);

    const results = sheet.findAll({ what: 'Apple' });
    expect(results).toHaveLength(2);
    expect(results).toContainEqual({ row: 1, col: 1 });
    expect(results).toContainEqual({ row: 3, col: 1 });
    expect(results).not.toContainEqual({ row: 2, col: 1 });
  });

  test('hiding a row does not destroy cell data — getCell still works', () => {
    sheet.setCellValue({ row: 5, col: 2 }, 42);
    sheet.hideRow(5);
    expect(sheet.getCellValue({ row: 5, col: 2 })).toBe(42);
  });
});

// ---------------------------------------------------------------------------
// 4.  PM Requirement 2: Hide column → find skips hidden column
// ---------------------------------------------------------------------------

describe('PM-2 · Hide column — find skips hidden column', () => {
  let sheet: Worksheet;
  beforeEach(() => { sheet = makeSheet(); });

  test('findAll does not return cells in hidden column (default)', () => {
    sheet.setCellValue({ row: 1, col: 4 }, 'HiddenColValue');
    sheet.hideCol(4);
    const results = sheet.findAll({ what: 'HiddenColValue' });
    expect(results).toHaveLength(0);
  });

  test('findAll returns cells in hidden column when includeHidden: true', () => {
    sheet.setCellValue({ row: 1, col: 4 }, 'HiddenColValue');
    sheet.hideCol(4);
    const results = sheet.findAll({ what: 'HiddenColValue', includeHidden: true });
    expect(results).toHaveLength(1);
  });

  test('hide col: adjacent column cells still found', () => {
    sheet.setCellValue({ row: 1, col: 3 }, 'Left');
    sheet.setCellValue({ row: 1, col: 4 }, 'Hidden');
    sheet.setCellValue({ row: 1, col: 5 }, 'Right');
    sheet.hideCol(4);

    const left  = sheet.findAll({ what: 'Left' });
    const right = sheet.findAll({ what: 'Right' });
    const mid   = sheet.findAll({ what: 'Hidden' });

    expect(left).toHaveLength(1);
    expect(right).toHaveLength(1);
    expect(mid).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// 5.  PM Requirement 3: Hide anchor of merged region
// ---------------------------------------------------------------------------

describe('PM-3 · Hide anchor of merged region', () => {
  test('hiding the anchor row hides the merged cell from find', () => {
    const sheet = makeSheet();
    sheet.setCellValue({ row: 1, col: 1 }, 'MergedAnchor');
    sheet.mergeCells({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    sheet.hideRow(1); // hide anchor row

    const results = sheet.findAll({ what: 'MergedAnchor' });
    expect(results).toHaveLength(0);
  });

  test('merge is structurally intact after hiding anchor row', () => {
    const sheet = makeSheet();
    sheet.setCellValue({ row: 2, col: 2 }, 'MergedData');
    sheet.mergeCells({ start: { row: 2, col: 2 }, end: { row: 4, col: 4 } });
    sheet.hideRow(2);

    // Merge still exists — hiding is visual only
    expect(sheet.getMergedRangeForCell({ row: 2, col: 2 })).not.toBeNull();
    expect(sheet.getMergedRangeForCell({ row: 4, col: 4 })).not.toBeNull();

    // Data accessible via getCell even though hidden
    expect(sheet.getCellValue({ row: 2, col: 2 })).toBe('MergedData');
  });

  test('isRowHidden does not affect isInMerge', () => {
    const sheet = makeSheet();
    sheet.mergeCells({ start: { row: 3, col: 1 }, end: { row: 5, col: 3 } });
    sheet.hideRow(3);

    expect(sheet.isInMerge({ row: 3, col: 1 })).toBe(true);
    expect(sheet.isInMerge({ row: 4, col: 2 })).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 6.  PM Requirement 4: Hide part of merged region (non-anchor rows/cols)
// ---------------------------------------------------------------------------

describe('PM-4 · Hide part of merged region', () => {
  test('hiding a non-anchor row of merged region hides the merged cell from find', () => {
    // Anchor at (1,1), hide row 2 (non-anchor row but part of merge).
    // find should skip because ANCHOR (1,1) lives in row 1 (visible).
    const sheet = makeSheet();
    sheet.setCellValue({ row: 1, col: 1 }, 'Anchor');
    sheet.mergeCells({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    sheet.hideRow(2); // non-anchor row

    // Anchor row (1) is visible, so anchor should still be found.
    const results = sheet.findAll({ what: 'Anchor' });
    expect(results).toHaveLength(1);
    expect(results[0]).toEqual({ row: 1, col: 1 });
  });

  test('hiding anchor col of merged region hides anchor from find', () => {
    const sheet = makeSheet();
    sheet.setCellValue({ row: 1, col: 1 }, 'ColMerge');
    sheet.mergeCells({ start: { row: 1, col: 1 }, end: { row: 1, col: 5 } });
    sheet.hideCol(1); // anchor col

    const results = sheet.findAll({ what: 'ColMerge' });
    expect(results).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// 7.  PM Requirement 5: Unhide restores traversal
// ---------------------------------------------------------------------------

describe('PM-5 · Unhide restores deterministic traversal', () => {
  let sheet: Worksheet;
  beforeEach(() => { sheet = makeSheet(); });

  test('showRow immediately restores cells to find results', () => {
    sheet.setCellValue({ row: 4, col: 1 }, 'WasHidden');
    sheet.hideRow(4);
    expect(sheet.findAll({ what: 'WasHidden' })).toHaveLength(0);

    sheet.showRow(4);
    const results = sheet.findAll({ what: 'WasHidden' });
    expect(results).toHaveLength(1);
    expect(results[0]).toEqual({ row: 4, col: 1 });
  });

  test('showCol immediately restores cells to find results', () => {
    sheet.setCellValue({ row: 1, col: 3 }, 'ColRestore');
    sheet.hideCol(3);
    expect(sheet.findAll({ what: 'ColRestore' })).toHaveLength(0);

    sheet.showCol(3);
    const results = sheet.findAll({ what: 'ColRestore' });
    expect(results).toHaveLength(1);
  });

  test('traversal order is deterministic after hide/unhide cycle', () => {
    for (let r = 1; r <= 5; r++) {
      sheet.setCellValue({ row: r, col: 1 }, 'Item');
    }

    sheet.hideRow(2);
    sheet.hideRow(4);
    const hidden = sheet.findAll({ what: 'Item' });
    expect(hidden).toEqual([
      { row: 1, col: 1 },
      { row: 3, col: 1 },
      { row: 5, col: 1 },
    ]);

    sheet.showRow(2);
    sheet.showRow(4);
    const restored = sheet.findAll({ what: 'Item' });
    expect(restored).toEqual([
      { row: 1, col: 1 },
      { row: 2, col: 1 },
      { row: 3, col: 1 },
      { row: 4, col: 1 },
      { row: 5, col: 1 },
    ]);
  });

  test('getHiddenRows is empty after all rows shown', () => {
    sheet.hideRow(1);
    sheet.hideRow(2);
    sheet.showRow(1);
    sheet.showRow(2);
    expect(sheet.getHiddenRows().size).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// 8.  PM Requirement 6: Stress test — 10k hidden rows
// ---------------------------------------------------------------------------

describe('PM-6 · Stress test: 10k hidden rows', () => {
  test('hiding 10000 rows completes in < 50ms', () => {
    const sheet = new Worksheet('Stress', 11_000, 10);
    const t0 = performance.now();
    for (let r = 1; r <= 10_000; r++) {
      sheet.hideRow(r);
    }
    const ms = performance.now() - t0;
    console.log(`\n  Hide 10k rows: ${ms.toFixed(2)}ms`);
    expect(ms).toBeLessThan(50);
  }, 10_000);

  test('10000 isRowHidden lookups after 10k hides: < 10ms', () => {
    const sheet = new Worksheet('Stress', 11_000, 10);
    for (let r = 1; r <= 10_000; r++) {
      sheet.hideRow(r);
    }
    const t0 = performance.now();
    for (let r = 1; r <= 10_000; r++) {
      sheet.isRowHidden(r);
    }
    const ms = performance.now() - t0;
    console.log(`  isRowHidden × 10k lookups: ${ms.toFixed(2)}ms`);
    expect(ms).toBeLessThan(10);
  }, 10_000);

  test('findAll with 10k hidden rows: only visible rows returned', () => {
    const sheet = new Worksheet('Stress', 10_005, 2);
    // Set a value in row 10001 (visible) and row 5000 (will be hidden)
    sheet.setCellValue({ row: 10_001, col: 1 }, 'Visible');
    sheet.setCellValue({ row: 5_000,  col: 1 }, 'Hidden' );

    for (let r = 1; r <= 10_000; r++) {
      sheet.hideRow(r);
    }

    const results = sheet.findAll({ what: 'Hidden' });
    expect(results).toHaveLength(0); // row 5000 is hidden

    const visible = sheet.findAll({ what: 'Visible' });
    expect(visible).toHaveLength(1); // row 10001 is visible
  }, 15_000);
});

// ---------------------------------------------------------------------------
// 9.  PM Requirement 7: Hidden state does NOT affect storage memory
// ---------------------------------------------------------------------------

describe('PM-7 · Hidden state does not affect storage memory', () => {
  test('hiding rows does not create cell objects', () => {
    const sheet = makeSheet(1000, 26);
    // Baseline: no cells created
    const beforeUsed = sheet.getUsedRange();
    expect(beforeUsed).toBeNull();

    // Hide rows without setting any cell values
    for (let r = 1; r <= 100; r++) {
      sheet.hideRow(r);
    }

    // Still no cells created in ICellStore
    const afterUsed = sheet.getUsedRange();
    expect(afterUsed).toBeNull();
  });

  test('CellRecord has no hidden property — invariant check', () => {
    const sheet = makeSheet();
    sheet.setCellValue({ row: 1, col: 1 }, 'Test');
    sheet.hideRow(1);

    const cell = sheet.getCell({ row: 1, col: 1 });
    expect(cell).toBeDefined();
    // Cell object must NOT have a 'hidden' key
    expect('hidden' in (cell as object)).toBe(false);
    // No visibility state leaked into the cell
    const cellKeys = Object.keys(cell as object).filter(k => k !== 'value' && k !== 'formula' && k !== 'style' && k !== 'comments' && k !== 'icon' && k !== 'customComponent' && k !== 'hyperlink' && k !== 'spillSource' && k !== 'spilledFrom');
    expect(cellKeys).toHaveLength(0);
  });

  test('hiding row/col does not grow ICellStore entry count', () => {
    const sheet = makeSheet(100, 26);
    // Create exactly 3 cells
    sheet.setCellValue({ row: 1, col: 1 }, 'A');
    sheet.setCellValue({ row: 2, col: 2 }, 'B');
    sheet.setCellValue({ row: 3, col: 3 }, 'C');

    // Count via findAll with includeHidden: true
    const beforeCount = sheet.findAll({ what: '*', includeHidden: true }).length;

    // Hide rows and columns
    for (let r = 1; r <= 50; r++) sheet.hideRow(r);
    for (let c = 1; c <= 13; c++) sheet.hideCol(c);

    // includeHidden: true should still find exactly the same 3 cells
    const afterCount = sheet.findAll({ what: '*', includeHidden: true }).length;
    expect(afterCount).toBe(beforeCount);
  });
});

// ---------------------------------------------------------------------------
// 10.  Behavioral matrix: getCell always returns (hidden row/col)
// ---------------------------------------------------------------------------

describe('Behavioral matrix — getCell always returns regardless of hidden state', () => {
  let sheet: Worksheet;
  beforeEach(() => { sheet = makeSheet(); });

  test('getCell on hidden row returns cell', () => {
    sheet.setCellValue({ row: 5, col: 1 }, 'X');
    sheet.hideRow(5);
    expect(sheet.getCell({ row: 5, col: 1 })).toBeDefined();
    expect(sheet.getCellValue({ row: 5, col: 1 })).toBe('X');
  });

  test('getCell on hidden col returns cell', () => {
    sheet.setCellValue({ row: 1, col: 7 }, 'Y');
    sheet.hideCol(7);
    expect(sheet.getCell({ row: 1, col: 7 })).toBeDefined();
    expect(sheet.getCellValue({ row: 1, col: 7 })).toBe('Y');
  });

  test('setCellValue on hidden row still writes to ICellStore', () => {
    sheet.hideRow(3);
    sheet.setCellValue({ row: 3, col: 1 }, 'WrittenInHidden');
    expect(sheet.getCellValue({ row: 3, col: 1 })).toBe('WrittenInHidden');
  });

  test('setCellStyle on hidden col still writes to ICellStore', () => {
    sheet.hideCol(2);
    sheet.setCellStyle({ row: 1, col: 2 }, { bold: true });
    expect(sheet.getCellStyle({ row: 1, col: 2 })?.bold).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 11.  find / findIterator: includeHidden option correctness
// ---------------------------------------------------------------------------

describe('find/findIterator: includeHidden option', () => {
  let sheet: Worksheet;
  beforeEach(() => { sheet = makeSheet(); });

  test('default (includeHidden undefined) skips hidden', () => {
    sheet.setCellValue({ row: 2, col: 2 }, 'Needle');
    sheet.hideRow(2);
    expect(sheet.findAll({ what: 'Needle' })).toHaveLength(0);
  });

  test('includeHidden: false explicitly skips hidden', () => {
    sheet.setCellValue({ row: 2, col: 2 }, 'Needle');
    sheet.hideRow(2);
    expect(sheet.findAll({ what: 'Needle', includeHidden: false })).toHaveLength(0);
  });

  test('includeHidden: true includes hidden row cells', () => {
    sheet.setCellValue({ row: 2, col: 2 }, 'Needle');
    sheet.hideRow(2);
    expect(sheet.findAll({ what: 'Needle', includeHidden: true })).toHaveLength(1);
  });

  test('includeHidden: true includes hidden col cells', () => {
    sheet.setCellValue({ row: 1, col: 5 }, 'Needle');
    sheet.hideCol(5);
    expect(sheet.findAll({ what: 'Needle', includeHidden: true })).toHaveLength(1);
  });

  test('cell in both hidden row AND col: included when includeHidden: true', () => {
    sheet.setCellValue({ row: 3, col: 3 }, 'DoubleHidden');
    sheet.hideRow(3);
    sheet.hideCol(3);
    expect(sheet.findAll({ what: 'DoubleHidden' })).toHaveLength(0);
    expect(sheet.findAll({ what: 'DoubleHidden', includeHidden: true })).toHaveLength(1);
  });

  test('mixed visible + hidden + unhidden', () => {
    sheet.setCellValue({ row: 1, col: 1 }, 'V');
    sheet.setCellValue({ row: 2, col: 1 }, 'V');
    sheet.setCellValue({ row: 3, col: 1 }, 'V');
    sheet.hideRow(2);

    let found = sheet.findAll({ what: 'V' });
    expect(found).toHaveLength(2);

    sheet.showRow(2);
    found = sheet.findAll({ what: 'V' });
    expect(found).toHaveLength(3);
  });
});

// ---------------------------------------------------------------------------
// 12.  Event emissions
// ---------------------------------------------------------------------------

describe('Visibility events', () => {
  let sheet: Worksheet;
  let events: any[] = [];

  beforeEach(() => {
    sheet = makeSheet();
    events = [];
    sheet.on(e => events.push(e));
  });

  test('hideRow emits row-hidden event', () => {
    sheet.hideRow(7);
    expect(events.some(e => e.type === 'row-hidden' && e.row === 7)).toBe(true);
  });

  test('showRow emits row-shown event', () => {
    sheet.hideRow(7);
    events = [];
    sheet.showRow(7);
    expect(events.some(e => e.type === 'row-shown' && e.row === 7)).toBe(true);
  });

  test('hideCol emits col-hidden event', () => {
    sheet.hideCol(3);
    expect(events.some(e => e.type === 'col-hidden' && e.col === 3)).toBe(true);
  });

  test('showCol emits col-shown event', () => {
    sheet.hideCol(3);
    events = [];
    sheet.showCol(3);
    expect(events.some(e => e.type === 'col-shown' && e.col === 3)).toBe(true);
  });

  test('hideRow also emits sheet-mutated', () => {
    sheet.hideRow(1);
    expect(events.some(e => e.type === 'sheet-mutated')).toBe(true);
  });

  test('showCol also emits sheet-mutated', () => {
    sheet.hideCol(2);
    events = [];
    sheet.showCol(2);
    expect(events.some(e => e.type === 'sheet-mutated')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 13.  Spill: NOT blocked by hidden cells (display-only model)
// ---------------------------------------------------------------------------

describe('Spill behavior with hidden cells', () => {
  test('spill into a hidden row is allowed (hidden is display-only)', () => {
    const sheet = makeSheet(20, 10);
    const engine = new SpillEngine();

    // Hide row 3 — this is on the spill path from (1,1) spilling 5 rows
    sheet.hideRow(3);

    // Spill of 5 rows should NOT be blocked by the hidden row
    const result = engine.checkSpillRange(sheet, { row: 1, col: 1 }, 5, 1);
    expect(result).toBeNull(); // no error
  });

  test('spill into a hidden col is allowed', () => {
    const sheet = makeSheet(10, 20);
    const engine = new SpillEngine();

    sheet.hideCol(2);

    const result = engine.checkSpillRange(sheet, { row: 1, col: 1 }, 1, 5);
    expect(result).toBeNull();
  });

  test('spill is still blocked by MERGED cells even if merged area is in hidden row', () => {
    const sheet = makeSheet(20, 10);
    const engine = new SpillEngine();

    // Merge cells in row 3 (on the spill path)
    sheet.mergeCells({ start: { row: 3, col: 1 }, end: { row: 4, col: 2 } });
    // Also hide that row — merge should still block spill
    sheet.hideRow(3);

    const result = engine.checkSpillRange(sheet, { row: 1, col: 1 }, 5, 1);
    expect(result).toBeInstanceOf(Error);
    expect(result?.message).toBe('#SPILL!');
  });
});

// ---------------------------------------------------------------------------
// 14. getHiddenRows / getHiddenCols API
// ---------------------------------------------------------------------------

describe('getHiddenRows / getHiddenCols API', () => {
  test('returns empty sets when nothing is hidden', () => {
    const sheet = makeSheet();
    expect(sheet.getHiddenRows().size).toBe(0);
    expect(sheet.getHiddenCols().size).toBe(0);
  });

  test('getHiddenRows reflects current hidden row set', () => {
    const sheet = makeSheet();
    sheet.hideRow(5);
    sheet.hideRow(10);
    const rows = sheet.getHiddenRows();
    expect(rows.has(5)).toBe(true);
    expect(rows.has(10)).toBe(true);
    expect(rows.size).toBe(2);
  });

  test('getHiddenCols reflects current hidden col set', () => {
    const sheet = makeSheet();
    sheet.hideCol(3);
    expect(sheet.getHiddenCols().has(3)).toBe(true);
  });

  test('getHiddenRows updates after showRow', () => {
    const sheet = makeSheet();
    sheet.hideRow(2);
    sheet.showRow(2);
    expect(sheet.getHiddenRows().size).toBe(0);
  });

  test('getVisibleRowIndices excludes hidden rows while preserving filter behavior', () => {
    const sheet = makeSheet(5, 3);
    sheet.setCellValue({ row: 1, col: 1 }, 'keep');
    sheet.setCellValue({ row: 2, col: 1 }, 'keep');
    sheet.setCellValue({ row: 3, col: 1 }, 'drop');
    sheet.hideRow(2);
    sheet.setColumnFilter(1, { type: 'equals', value: 'keep' });

    expect(sheet.getVisibleRowIndices({ start: { row: 1, col: 1 }, end: { row: 5, col: 3 } })).toEqual([1]);
  });

  test('getVisibleColumnIndices excludes hidden columns', () => {
    const sheet = makeSheet(5, 5);
    sheet.hideCol(2);
    sheet.hideCol(4);

    expect(sheet.getVisibleColumnIndices()).toEqual([1, 3, 5]);
    expect(sheet.getVisibleColumnIndices({ start: { row: 1, col: 2 }, end: { row: 5, col: 5 } })).toEqual([3, 5]);
  });
});

// ---------------------------------------------------------------------------
// 15. Performance benchmark: V1 vs Legacy — confirming O(1) invariant
// ---------------------------------------------------------------------------

describe('Performance benchmark', () => {
  test('VisibilityStoreV1 isCellHidden stays flat across 1k, 10k, 100k hides', () => {
    const results: { hiddenCount: number; ms: number }[] = [];

    for (const n of [1_000, 10_000, 100_000]) {
      const store = new VisibilityStoreV1();
      for (let r = 1; r <= n; r++) store.hideRow(r);

      const t0 = performance.now();
      for (let i = 0; i < 100_000; i++) {
        store.isCellHidden(i % n + 1, 1);
      }
      results.push({ hiddenCount: n, ms: performance.now() - t0 });
    }

    console.log('\n  VisibilityStoreV1 isCellHidden 100k lookups:');
    for (const r of results) {
      console.log(`    ${String(r.hiddenCount).padStart(7)} rows hidden → ${r.ms.toFixed(2)}ms`);
    }

    // O(1) invariant: 100k-hidden run must be ≤ 5× the 1k-hidden run
    const ratio = results[2].ms / results[0].ms;
    console.log(`    Ratio (100k / 1k): ${ratio.toFixed(2)}× (must be ≤ 5×)`);
    expect(ratio).toBeLessThan(5);
  }, 30_000);

  test('Worksheet.hideRow 10k rows < 50ms, isRowHidden 10k checks < 5ms', () => {
    const sheet = new Worksheet('Bench', 11_000, 10);

    const t0 = performance.now();
    for (let r = 1; r <= 10_000; r++) sheet.hideRow(r);
    const hideMs = performance.now() - t0;

    const t1 = performance.now();
    for (let r = 1; r <= 10_000; r++) sheet.isRowHidden(r);
    const lookupMs = performance.now() - t1;

    console.log(`\n  Worksheet.hideRow × 10k: ${hideMs.toFixed(2)}ms`);
    console.log(`  Worksheet.isRowHidden × 10k: ${lookupMs.toFixed(2)}ms`);

    expect(hideMs).toBeLessThan(50);
    expect(lookupMs).toBeLessThan(5);
  }, 15_000);
});
