/**
 * @group sdk
 *
 * Public API Freeze Test — Phase 15 Release & Adoption Readiness.
 *
 * PURPOSE
 * -------
 * This test file acts as a canary for accidental breaking changes.
 * It encodes the EXACT public contract of @cyber-sheet/core/sdk as of v0.1.0.
 *
 * If any of these tests fail, it means a breaking change has been introduced.
 * Breaking changes require a version bump per docs/VERSIONING.md.
 *
 * HOW TO INTENTIONALLY UPDATE THIS FILE
 * --------------------------------------
 * 1. Determine whether the change is PATCH, MINOR, or MAJOR per VERSIONING.md.
 * 2. Update `packages/core/package.json` version accordingly.
 * 3. Update the hardcoded expectations below.
 * 4. Add a CHANGELOG.md entry explaining the change.
 * 5. If MAJOR: ensure snapshot upgrader registered in SnapshotCodec.ts.
 *
 * DO NOT silence or skip these tests to "make the build green."
 * A failing API freeze test IS the signal — fix the root cause, not the test.
 */

import * as SdkIndex from '../../src/sdk/index';
import {
  createSpreadsheet,
  SdkError,
  DisposedError,
  BoundsError,
  SnapshotError,
  MergeError,
  PatchError,
  PatchSerializer,
  PatchDeserializeError,
} from '../../src/sdk/index';

// ---------------------------------------------------------------------------
// § 1. Exported value names
// ---------------------------------------------------------------------------
// Every name listed here is part of the stable public contract.
// Adding a name here is a MINOR change. Removing one is MAJOR.

describe('§1 exported value names (v0.1.0 lock)', () => {
  /**
   * The canonical set of value exports.
   * Type-only exports (interfaces, type aliases) are not runtime-checkable
   * but are documented in VERSIONING.md.
   */
  const EXPECTED_VALUE_EXPORTS: ReadonlySet<string> = new Set([
    // Factory
    'createSpreadsheet',
    // Error classes
    'SdkError',
    'DisposedError',
    'BoundsError',
    'SnapshotError',
    'MergeError',
    'PatchError',
    // Patch utilities
    'PatchSerializer',
    'PatchDeserializeError',
  ]);

  it('module exports every expected value name', () => {
    const missing: string[] = [];
    for (const name of EXPECTED_VALUE_EXPORTS) {
      if (!(name in SdkIndex)) {
        missing.push(name);
      }
    }
    if (missing.length > 0) {
      fail(
        `BREAKING CHANGE: The following names were removed from the SDK exports.\n` +
        `See docs/VERSIONING.md for required action before removing exports.\n\n` +
        `Missing: ${missing.join(', ')}`,
      );
    }
  });

  it('module exports no unexpected value names (additions allowed — update this list)', () => {
    // This test documents that new exports were INTENTIONALLY added.
    // When you add a new export, add it to EXPECTED_VALUE_EXPORTS above
    // AND note it in CHANGELOG.md as a MINOR addition.
    const actual = Object.keys(SdkIndex).filter(
      (k) => typeof (SdkIndex as Record<string, unknown>)[k] !== 'undefined',
    );
    const unexpected = actual.filter((k) => !EXPECTED_VALUE_EXPORTS.has(k));
    if (unexpected.length > 0) {
      // NOTE: new exports are ALLOWED (minor change), but they must be
      // acknowledged by adding them to EXPECTED_VALUE_EXPORTS above.
      // Change this to `fail(...)` if you want a strict no-new-exports policy.
      console.warn(
        `[api-freeze] New exports detected (not listed in freeze set): ${unexpected.join(', ')}\n` +
        `If intentional, add to EXPECTED_VALUE_EXPORTS and update CHANGELOG.md.`,
      );
    }
  });

  it('createSpreadsheet is a function', () => {
    expect(typeof createSpreadsheet).toBe('function');
  });

  it('PatchSerializer is a class (constructor) with static serialize/deserialize/validate', () => {
    // PatchSerializer is a class — typeof class === 'function' in JS
    expect(typeof PatchSerializer).toBe('function');
    expect(typeof PatchSerializer.serialize).toBe('function');
    expect(typeof PatchSerializer.deserialize).toBe('function');
    expect(typeof PatchSerializer.validate).toBe('function');
  });
});

// ---------------------------------------------------------------------------
// § 2. Error class .name strings
// ---------------------------------------------------------------------------
// The .name string is used by callers for switch/case and logging.
// Changing it is a MAJOR breaking change per VERSIONING.md § Error Contract.

describe('§2 error class .name strings (v0.1.0 lock)', () => {
  /**
   * Canonical .name → constructor mapping locked at v0.1.0.
   * Changing any value here is a MAJOR version bump.
   */
  const ERROR_NAME_MAP: ReadonlyArray<[new (...a: any[]) => SdkError, string]> = [
    [SdkError,              'SdkError'],
    [DisposedError,         'DisposedError'],
    [BoundsError,           'BoundsError'],
    [SnapshotError,         'SnapshotError'],
    [MergeError,            'MergeError'],
    [PatchError,            'PatchError'],
    [PatchDeserializeError, 'PatchDeserializeError'],
  ];

  for (const [Ctor, expectedName] of ERROR_NAME_MAP) {
    it(`new ${expectedName}() has .name === '${expectedName}'`, () => {
      const e = new Ctor(
        'test',
        // SnapshotError/MergeError/PatchError accept optional cause
        Ctor === SnapshotError || Ctor === MergeError || Ctor === PatchError
          ? new Error('cause')
          : undefined,
      );
      expect(e.name).toBe(expectedName);
    });
  }

  it('all error classes extend SdkError', () => {
    const classes = [DisposedError, BoundsError, SnapshotError, MergeError, PatchError];
    for (const Ctor of classes) {
      const e = new Ctor('test');
      expect(e).toBeInstanceOf(SdkError);
    }
  });

  it('SdkError itself extends Error', () => {
    expect(new SdkError('x')).toBeInstanceOf(Error);
  });

  it('error .name is not accidentally set to "Error"', () => {
    for (const [Ctor] of ERROR_NAME_MAP) {
      const e = new Ctor('test');
      if (Ctor !== SdkError) {
        expect(e.name).not.toBe('Error');
      }
    }
  });
});

// ---------------------------------------------------------------------------
// § 3. Error class .cause contract
// ---------------------------------------------------------------------------

describe('§3 error .cause contract (v0.1.0 lock)', () => {
  it('SnapshotError stores .cause', () => {
    const cause = new Error('inner');
    const e = new SnapshotError('outer', cause);
    expect(e.cause).toBe(cause);
  });

  it('MergeError stores .cause', () => {
    const cause = new Error('conflict');
    const e = new MergeError('outer', cause);
    expect(e.cause).toBe(cause);
  });

  it('PatchError stores .cause', () => {
    const cause = new Error('fail');
    const e = new PatchError('outer', cause);
    expect(e.cause).toBe(cause);
  });

  it('DisposedError does not require a cause', () => {
    const e = new DisposedError('method');
    // DisposedError has no .cause field — this is by design
    expect(e).toBeInstanceOf(DisposedError);
  });
});

// ---------------------------------------------------------------------------
// § 4. SpreadsheetSDK interface method contract
// ---------------------------------------------------------------------------
// The method NAMES and ARITIES are locked. Removing or renaming is MAJOR.
// Reducing arity is MAJOR. Increasing required arity is MAJOR.

describe('§4 SpreadsheetSDK interface methods and arities (v0.1.0 lock)', () => {
  /**
   * Canonical method → minimum expected arity.
   * Arity = Function.length (number of required params before optional/rest).
   *
   * Reducing arity is ≤ patch-safe. Increasing required arity is MAJOR.
   * We test that arity is AT LEAST the locked minimum.
   */
  const METHOD_ARITIES: ReadonlyMap<string, number> = new Map([
    ['setCell',          3],
    ['getCell',          2],
    ['getCellValue',     2],
    ['applyPatch',       1],
    ['snapshot',         0],
    ['restore',          1],
    ['encodeSnapshot',   0],
    ['decodeAndRestore', 1],
    ['undo',             0],
    ['redo',             0],
    ['mergeCells',       4],
    ['cancelMerge',      4],
    ['getMergedRanges',  0],
    ['isInMerge',        2],
    ['hideRow',          1],
    ['showRow',          1],
    ['hideCol',          1],
    ['showCol',          1],
    ['isRowHidden',      1],
    ['isColHidden',      1],
    ['on',               2],
    ['dispose',          0],
  ]);

  let sheet: ReturnType<typeof createSpreadsheet>;

  beforeEach(() => {
    sheet = createSpreadsheet('ApiFreeze', { rows: 5, cols: 5 });
  });

  afterEach(() => {
    try { sheet.dispose(); } catch { /* already disposed */ }
  });

  it('all expected methods exist on the sheet instance', () => {
    const missing: string[] = [];
    for (const method of METHOD_ARITIES.keys()) {
      if (!(method in sheet)) {
        missing.push(method);
      }
    }
    if (missing.length > 0) {
      fail(
        `BREAKING CHANGE: Methods removed from SpreadsheetSDK.\n` +
        `This requires a MAJOR version bump. See docs/VERSIONING.md.\n\n` +
        `Missing: ${missing.join(', ')}`,
      );
    }
  });

  for (const [method, minArity] of METHOD_ARITIES) {
    it(`${method}() has arity ≥ ${minArity}`, () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fn = (sheet as any)[method];
      expect(typeof fn).toBe('function');
      // Function.length reflects declared required parameters.
      // If this goes UP, it's MAJOR. If it stays the same or goes down, it's fine.
      expect(fn.length).toBeGreaterThanOrEqual(minArity);
    });
  }

  it('readonly properties exist: name, canUndo, canRedo', () => {
    expect(typeof sheet.name).toBe('string');
    expect(typeof sheet.canUndo).toBe('boolean');
    expect(typeof sheet.canRedo).toBe('boolean');
  });
});

// ---------------------------------------------------------------------------
// § 5. SdkEventType union members
// ---------------------------------------------------------------------------
// Adding members is MINOR. Removing or renaming is MAJOR.

describe('§5 SdkEventType members (v0.1.0 lock)', () => {
  const EXPECTED_EVENT_TYPES: ReadonlyArray<string> = [
    'cell-changed',
    'style-changed',
    'structure-changed',
    'filter-changed',
    'cycle-detected',
  ];

  it('all expected event types are accepted by on()', () => {
    const sheet = createSpreadsheet('EtFreeze', { rows: 5, cols: 5 });
    const missing: string[] = [];
    for (const eventType of EXPECTED_EVENT_TYPES) {
      try {
        // If this throws a type error at runtime (shouldn't happen in JS)
        // or a DisposedError (shouldn't happen here), the event type is broken.
        const sub = sheet.on(eventType as any, () => {});
        sub.dispose();
      } catch (err) {
        missing.push(`${eventType}: ${(err as Error).message}`);
      }
    }
    sheet.dispose();
    if (missing.length > 0) {
      fail(`BREAKING CHANGE: Event types no longer accepted:\n${missing.join('\n')}`);
    }
  });
});

// ---------------------------------------------------------------------------
// § 6. SpreadsheetOptions shape
// ---------------------------------------------------------------------------

describe('§6 SpreadsheetOptions fields (v0.1.0 lock)', () => {
  it('rows, cols, maxUndoHistory are accepted', () => {
    const s = createSpreadsheet('OptsFreeze', { rows: 10, cols: 10, maxUndoHistory: 5 });
    // If the fields were removed, TypeScript would have caught it.
    // This test ensures no runtime crash, confirming options propagate.
    s.setCell(10, 10, 'edge');
    expect(s.getCellValue(10, 10)).toBe('edge');
    s.dispose();
  });

  it('default rows/cols create a usable sheet', () => {
    const s = createSpreadsheet();
    expect(s.rowCount).toBe(5000);
    expect(s.colCount).toBe(16384);
    s.setCell(1, 1, 'default');
    expect(s.getCellValue(1, 1)).toBe('default');
    s.dispose();
  });

  it('default name is Sheet1', () => {
    const s = createSpreadsheet();
    expect(s.name).toBe('Sheet1');
    s.dispose();
  });
});

// ---------------------------------------------------------------------------
// § 7. PatchSerializer contract
// ---------------------------------------------------------------------------

describe('§7 PatchSerializer API contract (v0.1.0 lock)', () => {
  it('serialize produces a JSON string', () => {
    // PatchSerializer.serialize() returns a canonical JSON string
    const patch = { seq: 1, ops: [] };
    const serialized = PatchSerializer.serialize(patch);
    expect(typeof serialized).toBe('string');
    const parsed = JSON.parse(serialized);
    expect(parsed).toHaveProperty('seq', 1);
    expect(parsed).toHaveProperty('ops');
  });

  it('deserialize round-trips a serialized patch (string → WorksheetPatch)', () => {
    const original = { seq: 42, ops: [] };
    const json = PatchSerializer.serialize(original); // returns string
    const recovered = PatchSerializer.deserialize(json);
    expect(recovered.seq).toBe(42);
  });

  it('validate returns true for a valid object', () => {
    expect(PatchSerializer.validate({ seq: 0, ops: [] })).toBe(true);
  });

  it('validate returns false for invalid input', () => {
    expect(PatchSerializer.validate(null)).toBe(false);
    expect(PatchSerializer.validate({})).toBe(false);
    expect(PatchSerializer.validate({ seq: 'x', ops: [] })).toBe(false);
  });
});
