/**
 * SnapshotCodec.ts — Phase 7: Binary Snapshot Engine
 *
 * Encodes and decodes Worksheet state to/from a compact binary format designed
 * for fast load, minimal memory overhead, and streaming-friendly layout.
 *
 * =============================================================================
 * BINARY FORMAT: CSEX v1
 * =============================================================================
 *
 *  ┌─────────────────────────────────────────────────────────────────────────┐
 *  │  HEADER  (16 bytes)                                                     │
 *  │    magic:        4 bytes  [0x43 0x53 0x45 0x58] = "CSEX"               │
 *  │    version:      u16 LE   = 1                                           │
 *  │    sectionCount: u16 LE   = number of sections                         │
 *  │    reserved:     8 bytes  = 0x00 × 8                                   │
 *  ├─────────────────────────────────────────────────────────────────────────┤
 *  │  SECTION TABLE  (sectionCount × 12 bytes)                               │
 *  │    id:     u16 LE  section type identifier                              │
 *  │    flags:  u16 LE  reserved = 0                                         │
 *  │    offset: u32 LE  byte offset from start of buffer                    │
 *  │    length: u32 LE  byte count of section data                           │
 *  ├─────────────────────────────────────────────────────────────────────────┤
 *  │  SECTION DATA (variable, in table order)                                │
 *  └─────────────────────────────────────────────────────────────────────────┘
 *
 * Section IDs:
 *   0x0001  CELLS       cell values, formulas, styles, comments, icons, spill
 *   0x0002  MERGES      merged region descriptors
 *   0x0003  VISIBILITY  hidden rows and hidden columns
 *   0x0004  DAG         formula dependency edges (predecessors map)
 *   0x0005  VOLATILES   volatile cell registrations
 *
 * Cell value tags (vtag byte):
 *   0x00  null     (no additional data)
 *   0x01  number   f64 LE (8 bytes)
 *   0x02  string   u32 len + UTF-8 bytes
 *   0x03  boolean  u8  (0 = false, 1 = true)
 *   0x04  richtext u32 len + JSON UTF-8 bytes
 *   0x05  entity   u32 len + JSON UTF-8 bytes
 *
 * Cell flags byte:
 *   bit 0 (0x01) — has formula
 *   bit 1 (0x02) — has style
 *   bit 2 (0x04) — has comments
 *   bit 3 (0x08) — has icon
 *   bit 4 (0x10) — has spillSource
 *   bit 5 (0x20) — has spilledFrom
 *
 * =============================================================================
 * MEMORY PROFILE
 * =============================================================================
 *
 * Structural overhead per cell (minimum — null value, no extras):
 *   row(4) + col(4) + vtag(1) + flags(1) = 10 bytes
 *
 * Typical number cell (no formula, no style):
 *   10 + f64(8) = 18 bytes
 *
 * Typical formula cell (string formula, no style):
 *   10 + f64(8) + formula_u16_len(2) + formula_utf8 ≈ ~30 bytes
 *
 * Vs JSON.stringify(worksheet): variable but typically ~60–120 bytes/cell.
 *
 * =============================================================================
 * DESIGN DECISIONS
 * =============================================================================
 *
 * 1. Section table with offsets enables streaming load: skip sections not needed.
 * 2. Style / comments / icon use JSON blobs because these types are complex and
 *    rare. Binary encoding of arbitrary nested objects provides diminishing
 *    returns — JSON blobs are still smaller than V8 heap objects.
 * 3. Formula strings use u16 prefix (cap 65535 bytes) — no realistic formula
 *    approaches this limit.
 * 4. All multi-byte numerics are little-endian (LE) — consistent with Typed
 *    Array spec and x86/ARM native endianness.
 * 5. TextEncoder / TextDecoder: globally available in Node.js ≥ 11 and all
 *    modern browsers — no polyfill required.
 */

import type { Cell, MergedRegion, Address, CellStyle, CellComment, CellIcon } from '../types';
import type { SerializedDrawingLayer } from '../DrawingLayer';

// ---------------------------------------------------------------------------
// Snapshot types (plain objects — no class instances)
// ---------------------------------------------------------------------------

/** One cell entry in the snapshot. */
export type CellEntry = {
  row: number;
  col: number;
  cell: Cell;
};

/** One formula cell's dependency list in the snapshot. */
export type DagEdge = {
  row: number;
  col: number;
  deps: Address[];
};

/**
 * Plain-object representation of all serialisable Worksheet state.
 *
 * Does NOT include: column widths, row heights, column filters,
 * conditional formatting rules, or formula engine reference.
 * Those are either UI-layer concerns or reconstructed separately.
 *
 * Phase 8+ will extend this type with additional sections.
 */
export type WorksheetSnapshot = {
  /** Format version (currently 1). Validated on decode. */
  version: number;
  /** All populated cells. */
  cells: CellEntry[];
  /** Worksheet row count. Optional for v1/v2 binary compatibility. */
  rowCount?: number;
  /** Worksheet column count. Optional for v1/v2 binary compatibility. */
  colCount?: number;
  /** All merged regions. */
  merges: MergedRegion[];
  /** Indices of hidden rows (1-based). */
  hiddenRows: number[];
  /** Indices of hidden columns (1-based). */
  hiddenCols: number[];
  /** Explicit row heights. Optional for v1/v2 binary compatibility. */
  rowHeights?: Array<{ row: number; height: number }>;
  /** Explicit column widths. Optional for v1/v2 binary compatibility. */
  columnWidths?: Array<{ col: number; width: number }>;
  /** Whole-row default styles. Optional for v1/v2 binary compatibility. */
  rowStyles?: Array<{ row: number; style: CellStyle }>;
  /** Whole-column default styles. Optional for v1/v2 binary compatibility. */
  columnStyles?: Array<{ col: number; style: CellStyle }>;
  /** Floating and cell-bound drawing objects. Optional for v1/v2 binary compatibility. */
  drawings?: SerializedDrawingLayer;
  /** Formula dependency edges (predecessors). Successors are derived. */
  dagEdges: DagEdge[];
  /** Volatile cell registrations (e.g., NOW, RAND). */
  volatiles: Address[];
};

// ---------------------------------------------------------------------------
// Snapshot versioning & upgrade pipeline (Phase 8)
// ---------------------------------------------------------------------------

/**
 * A single-step snapshot upgrader from one binary format version to the next.
 * Upgraders operate on *decoded* WorksheetSnapshot objects — not raw binary.
 * Snapshots are immutable: upgrade() must return a new object.
 */
export interface SnapshotUpgrader {
  /** Source format version this upgrader handles. */
  readonly fromVersion: number;
  /** Target format version this upgrader produces. */
  readonly toVersion:   number;
  /**
   * Transform a snapshot from `fromVersion` to `toVersion`.
   * Must return a **new** object (do not mutate the input).
   */
  upgrade(snapshot: WorksheetSnapshot): WorksheetSnapshot;
}

/**
 * Registry of known snapshot upgraders.
 * Chains multiple upgraders automatically to bridge arbitrary version gaps.
 *
 * Usage:
 * ```ts
 * snapshotUpgraderRegistry.register({ fromVersion: 2, toVersion: 3, upgrade(s) { ... } });
 * const upgraded = snapshotUpgraderRegistry.apply(snapshot, FORMAT_VERSION);
 * ```
 */
export class SnapshotUpgraderRegistry {
  private readonly map = new Map<number, SnapshotUpgrader>();

  /**
   * Register an upgrader.
   * @throws if another upgrader is already registered for the same `fromVersion`.
   */
  register(upgrader: SnapshotUpgrader): void {
    if (this.map.has(upgrader.fromVersion)) {
      throw new Error(
        `SnapshotUpgraderRegistry: upgrader for fromVersion=${upgrader.fromVersion} already registered.`,
      );
    }
    this.map.set(upgrader.fromVersion, upgrader);
  }

  /**
   * Walk the upgrade chain from `snapshot.version` up to `targetVersion`.
   * Returns the snapshot unchanged when already at target.
   * @throws if no upgrader is registered for an intermediate version.
   */
  apply(snapshot: WorksheetSnapshot, targetVersion: number): WorksheetSnapshot {
    let current = snapshot;
    while (current.version < targetVersion) {
      const upgrader = this.map.get(current.version);
      if (!upgrader) {
        throw new Error(
          `SnapshotUpgraderRegistry: no upgrader found for v${current.version} → v${current.version + 1}.`,
        );
      }
      current = upgrader.upgrade(current);
    }
    return current;
  }
}

/**
 * Module-level upgrader registry, pre-populated with all known version steps.
 * Import and use this to register custom application-level upgraders.
 */
export const snapshotUpgraderRegistry = new SnapshotUpgraderRegistry();

// v1 → v2: header now carries flags (u16) + CRC32 checksum (u32).
// The WorksheetSnapshot payload shape is identical; only the binary
// on-disk representation changed.  Upgrading just bumps the version field.
snapshotUpgraderRegistry.register({
  fromVersion: 1,
  toVersion:   2,
  upgrade(snapshot: WorksheetSnapshot): WorksheetSnapshot {
    return { ...snapshot, version: 2 };
  },
});

// ---------------------------------------------------------------------------
// Binary format constants
// ---------------------------------------------------------------------------

/** File signature: "CSEX" (CyberSheet EXcel). */
const MAGIC = new Uint8Array([0x43, 0x53, 0x45, 0x58]);

/** Current supported format version. */
export const FORMAT_VERSION = 2;

/** Total header size in bytes (v1 and v2 share the same 16-byte budget). */
const HEADER_BYTES       = 16;
const SECTION_DESC_BYTES = 12;

/**
 * Byte offset of the CRC32 checksum field within the v2+ header.
 * The field is set to 0 before computing the checksum, then patched.
 */
const CHECKSUM_OFFSET = 8;

/** Header flags field bit: CRC32 checksum is present and valid. */
const HDR_FLAG_CRC32 = 0x0001;

const SEC_CELLS      = 0x0001;
const SEC_MERGES     = 0x0002;
const SEC_VISIBILITY = 0x0003;
const SEC_DAG        = 0x0004;
const SEC_VOLATILES  = 0x0005;

const VTAG_NULL     = 0x00;
const VTAG_NUMBER   = 0x01;
const VTAG_STRING   = 0x02;
const VTAG_BOOLEAN  = 0x03;
const VTAG_RICHTEXT = 0x04;
const VTAG_ENTITY   = 0x05;

const FLAG_FORMULA   = 0x01;
const FLAG_STYLE     = 0x02;
const FLAG_COMMENTS  = 0x04;
const FLAG_ICON      = 0x08;
const FLAG_SPILLSRC  = 0x10;
const FLAG_SPILLFROM = 0x20;

// ---------------------------------------------------------------------------
// CRC32 — integrity checksum (Phase 8)
// ---------------------------------------------------------------------------

/**
 * Pre-computed CRC32 lookup table.
 * Polynomial 0xEDB88320 (reflected form) — compatible with IEEE 802.3,
 * zlib, gzip, and the `crc32` utility.
 */
const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c;
  }
  return t;
})();

/**
 * Compute CRC32 checksum of a Uint8Array.
 * Returns an unsigned 32-bit integer.
 *
 * Exported for downstream consumers who want to verify snapshot integrity
 * without going through the full codec decode path.
 */
export function crc32(data: Uint8Array): number {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc = CRC32_TABLE[(crc ^ data[i]) & 0xFF]! ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0; // unsigned 32-bit
}

// ---------------------------------------------------------------------------
// BinaryWriter — growable output buffer
// ---------------------------------------------------------------------------

/**
 * Resizable write-only buffer using a Uint8Array + DataView pair.
 * Grows by doubling when capacity is exceeded (amortised O(1) append).
 * All multi-byte values are written in little-endian order.
 */
class BinaryWriter {
  private buf: Uint8Array;
  private view: DataView;
  private pos = 0;
  private readonly enc = new TextEncoder();

  constructor(initialCapacity = 4096) {
    this.buf  = new Uint8Array(Math.max(16, initialCapacity));
    this.view = new DataView(this.buf.buffer);
  }

  get position(): number { return this.pos; }

  private _ensure(n: number): void {
    const need = this.pos + n;
    if (need <= this.buf.byteLength) return;
    let cap = this.buf.byteLength;
    do { cap *= 2; } while (cap < need);
    const next = new Uint8Array(cap);
    next.set(this.buf.subarray(0, this.pos));
    this.buf  = next;
    this.view = new DataView(this.buf.buffer);
  }

  writeU8(v: number): void  { this._ensure(1); this.view.setUint8(this.pos,    v);        this.pos += 1; }
  writeU16(v: number): void { this._ensure(2); this.view.setUint16(this.pos,  v, true);   this.pos += 2; }
  writeU32(v: number): void { this._ensure(4); this.view.setUint32(this.pos,  v, true);   this.pos += 4; }
  writeF64(v: number): void { this._ensure(8); this.view.setFloat64(this.pos, v, true);   this.pos += 8; }

  writeBytes(src: Uint8Array): void {
    this._ensure(src.byteLength);
    this.buf.set(src, this.pos);
    this.pos += src.byteLength;
  }

  /** Write u32-prefixed UTF-8 string. */
  writeString(s: string): void {
    const b = this.enc.encode(s);
    this.writeU32(b.byteLength);
    this.writeBytes(b);
  }

  /** Write u16-prefixed UTF-8 string (for short strings like formulas, max 65535 bytes). */
  writeShortString(s: string): void {
    const b = this.enc.encode(s);
    this.writeU16(b.byteLength);
    this.writeBytes(b);
  }

  /** Patch a u32 at an absolute byte position without moving the write cursor. */
  patchU32At(at: number, v: number): void {
    this.view.setUint32(at, v, true);
  }

  /** Return a copy of the valid region of the buffer. */
  toBuffer(): Uint8Array {
    return this.buf.slice(0, this.pos);
  }
}

// ---------------------------------------------------------------------------
// BinaryReader — cursor over a Uint8Array
// ---------------------------------------------------------------------------

/**
 * Read-only cursor over a Uint8Array.
 * All multi-byte values are read in little-endian order.
 * `seek()` supports random access for the section-based format.
 */
class BinaryReader {
  private readonly view: DataView;
  private pos = 0;
  private readonly dec = new TextDecoder();

  constructor(buf: Uint8Array) {
    // Use the buffer's full underlying ArrayBuffer, offset by buf.byteOffset.
    this.view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  }

  get position(): number { return this.pos; }

  readU8(): number  { const v = this.view.getUint8(this.pos);               this.pos += 1; return v; }
  readU16(): number { const v = this.view.getUint16(this.pos, true);        this.pos += 2; return v; }
  readU32(): number { const v = this.view.getUint32(this.pos, true);        this.pos += 4; return v; }
  readF64(): number { const v = this.view.getFloat64(this.pos, true);       this.pos += 8; return v; }

  readBytes(n: number): Uint8Array {
    const out = new Uint8Array(this.view.buffer, this.view.byteOffset + this.pos, n);
    this.pos += n;
    return out;
  }

  /** Read u32-prefixed UTF-8 string. */
  readString(): string {
    const len   = this.readU32();
    const bytes = this.readBytes(len);
    return this.dec.decode(bytes);
  }

  /** Read u16-prefixed UTF-8 string. */
  readShortString(): string {
    const len   = this.readU16();
    const bytes = this.readBytes(len);
    return this.dec.decode(bytes);
  }

  /** Seek to an absolute byte position (0-based within the original Uint8Array). */
  seek(pos: number): void { this.pos = pos; }
}

// ---------------------------------------------------------------------------
// SnapshotCodec — public API
// ---------------------------------------------------------------------------

/**
 * Encodes and decodes WorksheetSnapshot ↔ Uint8Array.
 *
 * Stateless — all methods are pure functions of their arguments.
 * Safe to use as a module-level singleton or as a constructor-created instance.
 *
 * Usage:
 * ```ts
 * const codec = new SnapshotCodec();
 * const buf   = codec.encode(worksheet.extractSnapshot());
 * // ... persist buf to disk / IndexedDB / network ...
 * const snap  = codec.decode(buf);
 * worksheet.applySnapshot(snap);
 * ```
 */
export class SnapshotCodec {
  // ── encode ────────────────────────────────────────────────────────────────

  /**
   * Encode a WorksheetSnapshot to a compact binary Uint8Array.
   *
   * @param snapshot  Plain WorksheetSnapshot object.
   * @returns         Binary buffer in CSEX v1 format.
   * @complexity      O(V + E) where V = cell+merge+visibility count, E = DAG edges.
   */
  encode(snapshot: WorksheetSnapshot): Uint8Array {
    // Build sections independently, then assemble header + table + data.
    const sections: Array<{ id: number; data: Uint8Array }> = [
      { id: SEC_CELLS,      data: this._encodeCells(snapshot.cells)                              },
      { id: SEC_MERGES,     data: this._encodeMerges(snapshot.merges)                            },
      { id: SEC_VISIBILITY, data: this._encodeVisibility(snapshot.hiddenRows, snapshot.hiddenCols) },
      { id: SEC_DAG,        data: this._encodeDAG(snapshot.dagEdges)                             },
      { id: SEC_VOLATILES,  data: this._encodeVolatiles(snapshot.volatiles)                      },
    ];

    const headerSize = HEADER_BYTES + sections.length * SECTION_DESC_BYTES;
    const dataSize   = sections.reduce((acc, s) => acc + s.data.byteLength, 0);
    const w = new BinaryWriter(headerSize + dataSize + 64);

    // ── Header (v2 layout, 16 bytes) ──────────────────────────────────────────────
    w.writeBytes(MAGIC);          // [0..3]  magic
    w.writeU16(FORMAT_VERSION);   // [4..5]  version = 2
    w.writeU16(HDR_FLAG_CRC32);   // [6..7]  flags   = CRC32 enabled
    w.writeU32(0);                // [8..11] checksum placeholder (patched below)
    w.writeU16(sections.length);  // [12..13] section count
    w.writeU16(0);                // [14..15] reserved

    // ── Section table (offsets patched after data is written) ───────────────
    const tableBase = w.position;   // = HEADER_BYTES = 16
    for (const { id } of sections) {
      w.writeU16(id);  // id
      w.writeU16(0);   // flags (reserved)
      w.writeU32(0);   // offset — placeholder
      w.writeU32(0);   // length — placeholder
    }

    // ── Section data ────────────────────────────────────────────────────────
    for (let i = 0; i < sections.length; i++) {
      const offset = w.position;
      w.writeBytes(sections[i].data);
      const length = sections[i].data.byteLength;

      // Patch section table entry: +4 = offset field, +8 = length field.
      const entryBase = tableBase + i * SECTION_DESC_BYTES;
      w.patchU32At(entryBase + 4, offset);
      w.patchU32At(entryBase + 8, length);
    }

    // CRC32: compute over the complete buffer while checksum field is still 0,
    // then patch the checksum into the returned copy.
    const result = w.toBuffer();
    const checksum = crc32(result);
    new DataView(result.buffer, result.byteOffset).setUint32(CHECKSUM_OFFSET, checksum, true);
    return result;
  }

  // ── decode ────────────────────────────────────────────────────────────────

  /**
   * Decode a CSEX binary buffer back to a WorksheetSnapshot.
   *
   * @param buf  Binary buffer produced by encode().
   * @throws     Error if magic bytes mismatch or version is unsupported.
   * @complexity O(V + E).
   */
  decode(buf: Uint8Array): WorksheetSnapshot {
    const r = new BinaryReader(buf);

    // ── Validate header ─────────────────────────────────────────────────────
    const magic = r.readBytes(4);
    for (let i = 0; i < 4; i++) {
      if (magic[i] !== MAGIC[i]) {
        throw new Error('SnapshotCodec: invalid magic bytes (expected CSEX format).');
      }
    }

    const version = r.readU16();

    let sectionCount: number;

    if (version === 1) {
      // v1 header layout: [4..5]=version [6..7]=sectionCount [8..15]=reserved
      // No flags, no checksum.  Section table starts at byte 16.
      sectionCount = r.readU16();
      r.seek(16); // skip 8 reserved bytes
    } else if (version === FORMAT_VERSION) {
      // v2 header layout: [4..5]=version [6..7]=flags [8..11]=checksum
      //                   [12..13]=sectionCount [14..15]=reserved
      const flags         = r.readU16();
      const storedCRC     = r.readU32();
      sectionCount        = r.readU16();
      r.readU16(); // reserved

      if (flags & HDR_FLAG_CRC32) {
        // Verify integrity: re-compute CRC with the checksum field zeroed.
        const copy = buf.slice();
        new DataView(copy.buffer).setUint32(CHECKSUM_OFFSET, 0, true);
        const computed = crc32(copy);
        if (computed !== storedCRC) {
          throw new Error(
            `SnapshotCodec: checksum mismatch (stored 0x${storedCRC.toString(16).padStart(8, '0')}, ` +
            `computed 0x${computed.toString(16).padStart(8, '0')}) — buffer may be corrupted.`,
          );
        }
      }
    } else {
      throw new Error(
        `SnapshotCodec: unsupported format version ${version} ` +
        `(current: ${FORMAT_VERSION}). Register a SnapshotUpgrader to decode older formats.`,
      );
    }

    // ── Read section table ──────────────────────────────────────────────────
    type SectionMeta = { id: number; offset: number };
    const table: SectionMeta[] = [];
    for (let i = 0; i < sectionCount; i++) {
      const id     = r.readU16();
      /*flags*/     r.readU16();
      const offset = r.readU32();
      /*length*/    r.readU32();
      table.push({ id, offset });
    }

    // ── Dispatch sections ───────────────────────────────────────────────────
    const snapshot: WorksheetSnapshot = {
      version,
      cells:      [],
      merges:     [],
      hiddenRows: [],
      hiddenCols: [],
      dagEdges:   [],
      volatiles:  [],
    };

    for (const { id, offset } of table) {
      r.seek(offset);
      switch (id) {
        case SEC_CELLS:      snapshot.cells    = this._decodeCells(r);           break;
        case SEC_MERGES:     snapshot.merges   = this._decodeMerges(r);          break;
        case SEC_VISIBILITY: this._decodeVisibility(r, snapshot);                break;
        case SEC_DAG:        snapshot.dagEdges = this._decodeDAG(r);             break;
        case SEC_VOLATILES:  snapshot.volatiles = this._decodeVolatiles(r);      break;
      }
    }

    // Apply registered upgraders if the decoded snapshot pre-dates FORMAT_VERSION.
    if (snapshot.version < FORMAT_VERSION) {
      return snapshotUpgraderRegistry.apply(snapshot, FORMAT_VERSION);
    }
    return snapshot;
  }

  // ── Section encoders ──────────────────────────────────────────────────────

  private _encodeCells(cells: CellEntry[]): Uint8Array {
    // Estimate capacity: ~32 bytes per cell average.
    const w = new BinaryWriter(Math.max(64, cells.length * 32 + 4));
    w.writeU32(cells.length);
    for (const { row, col, cell } of cells) {
      this._encodeCellEntry(w, row, col, cell);
    }
    return w.toBuffer();
  }

  private _encodeCellEntry(w: BinaryWriter, row: number, col: number, cell: Cell): void {
    w.writeU32(row);
    w.writeU32(col);
    this._encodeCellValue(w, cell.value);

    let flags = 0;
    if (cell.formula    !== undefined) flags |= FLAG_FORMULA;
    if (cell.style      !== undefined) flags |= FLAG_STYLE;
    if (cell.comments   !== undefined) flags |= FLAG_COMMENTS;
    if (cell.icon       !== undefined) flags |= FLAG_ICON;
    if (cell.spillSource  !== undefined) flags |= FLAG_SPILLSRC;
    if (cell.spilledFrom  !== undefined) flags |= FLAG_SPILLFROM;
    w.writeU8(flags);

    if (flags & FLAG_FORMULA)   w.writeShortString(cell.formula!);
    if (flags & FLAG_STYLE)     w.writeString(JSON.stringify(cell.style));
    if (flags & FLAG_COMMENTS)  w.writeString(JSON.stringify(cell.comments));
    if (flags & FLAG_ICON)      w.writeString(JSON.stringify(cell.icon));
    if (flags & FLAG_SPILLSRC) {
      const s = cell.spillSource!;
      w.writeU32(s.endAddress.row);
      w.writeU32(s.endAddress.col);
      w.writeU32(s.dimensions[0]);
      w.writeU32(s.dimensions[1]);
    }
    if (flags & FLAG_SPILLFROM) {
      w.writeU32(cell.spilledFrom!.row);
      w.writeU32(cell.spilledFrom!.col);
    }
  }

  private _encodeCellValue(w: BinaryWriter, value: Cell['value']): void {
    if (value === null || value === undefined) {
      w.writeU8(VTAG_NULL);
      return;
    }
    if (typeof value === 'number') {
      w.writeU8(VTAG_NUMBER);
      w.writeF64(value);
      return;
    }
    if (typeof value === 'boolean') {
      w.writeU8(VTAG_BOOLEAN);
      w.writeU8(value ? 1 : 0);
      return;
    }
    if (typeof value === 'string') {
      w.writeU8(VTAG_STRING);
      w.writeString(value);
      return;
    }
    // RichTextValue: { runs: RichTextRun[] }
    if (typeof value === 'object' && 'runs' in value) {
      w.writeU8(VTAG_RICHTEXT);
      w.writeString(JSON.stringify(value));
      return;
    }
    // EntityValue (any other structured object)
    w.writeU8(VTAG_ENTITY);
    w.writeString(JSON.stringify(value));
  }

  private _encodeMerges(merges: MergedRegion[]): Uint8Array {
    const w = new BinaryWriter(Math.max(8, merges.length * 16 + 4));
    w.writeU32(merges.length);
    for (const m of merges) {
      w.writeU32(m.startRow);
      w.writeU32(m.startCol);
      w.writeU32(m.endRow);
      w.writeU32(m.endCol);
    }
    return w.toBuffer();
  }

  private _encodeVisibility(hiddenRows: number[], hiddenCols: number[]): Uint8Array {
    const w = new BinaryWriter(Math.max(8, (hiddenRows.length + hiddenCols.length) * 4 + 8));
    w.writeU32(hiddenRows.length);
    for (const r of hiddenRows) w.writeU32(r);
    w.writeU32(hiddenCols.length);
    for (const c of hiddenCols) w.writeU32(c);
    return w.toBuffer();
  }

  private _encodeDAG(edges: DagEdge[]): Uint8Array {
    const totalDeps = edges.reduce((acc, e) => acc + e.deps.length, 0);
    const w = new BinaryWriter(Math.max(8, edges.length * 12 + totalDeps * 8 + 4));
    w.writeU32(edges.length);
    for (const { row, col, deps } of edges) {
      w.writeU32(row);
      w.writeU32(col);
      w.writeU32(deps.length);
      for (const { row: dr, col: dc } of deps) {
        w.writeU32(dr);
        w.writeU32(dc);
      }
    }
    return w.toBuffer();
  }

  private _encodeVolatiles(volatiles: Address[]): Uint8Array {
    const w = new BinaryWriter(Math.max(8, volatiles.length * 8 + 4));
    w.writeU32(volatiles.length);
    for (const { row, col } of volatiles) {
      w.writeU32(row);
      w.writeU32(col);
    }
    return w.toBuffer();
  }

  // ── Section decoders ──────────────────────────────────────────────────────

  private _decodeCells(r: BinaryReader): CellEntry[] {
    const count = r.readU32();
    const cells: CellEntry[] = [];
    for (let i = 0; i < count; i++) {
      const row   = r.readU32();
      const col   = r.readU32();
      const value = this._decodeCellValue(r);
      const flags = r.readU8();

      // Allocate with all fields present (mono-shape invariant).
      const cell: Cell = {
        value,
        formula:     undefined,
        style:       undefined,
        comments:    undefined,
        icon:        undefined,
        spillSource: undefined,
        spilledFrom: undefined,
      };

      if (flags & FLAG_FORMULA)  cell.formula   = r.readShortString();
      if (flags & FLAG_STYLE)    cell.style    = JSON.parse(r.readString())  as CellStyle;
      if (flags & FLAG_COMMENTS) cell.comments = JSON.parse(r.readString())  as CellComment[];
      if (flags & FLAG_ICON)     cell.icon     = JSON.parse(r.readString())  as CellIcon;
      if (flags & FLAG_SPILLSRC) {
        const endRow  = r.readU32();
        const endCol  = r.readU32();
        const dimRows = r.readU32();
        const dimCols = r.readU32();
        cell.spillSource = {
          endAddress: { row: endRow, col: endCol },
          dimensions: [dimRows, dimCols],
        };
      }
      if (flags & FLAG_SPILLFROM) {
        cell.spilledFrom = { row: r.readU32(), col: r.readU32() };
      }

      cells.push({ row, col, cell });
    }
    return cells;
  }

  private _decodeCellValue(r: BinaryReader): Cell['value'] {
    const tag = r.readU8();
    switch (tag) {
      case VTAG_NULL:     return null;
      case VTAG_NUMBER:   return r.readF64();
      case VTAG_BOOLEAN:  return r.readU8() !== 0;
      case VTAG_STRING:   return r.readString();
      case VTAG_RICHTEXT: return JSON.parse(r.readString());
      case VTAG_ENTITY:   return JSON.parse(r.readString());
      default:
        throw new Error(`SnapshotCodec: unknown value tag 0x${tag.toString(16).padStart(2, '0')}.`);
    }
  }

  private _decodeMerges(r: BinaryReader): MergedRegion[] {
    const count  = r.readU32();
    const merges: MergedRegion[] = [];
    for (let i = 0; i < count; i++) {
      merges.push({
        startRow: r.readU32(),
        startCol: r.readU32(),
        endRow:   r.readU32(),
        endCol:   r.readU32(),
      });
    }
    return merges;
  }

  private _decodeVisibility(r: BinaryReader, snap: WorksheetSnapshot): void {
    const rowCount = r.readU32();
    for (let i = 0; i < rowCount; i++) snap.hiddenRows.push(r.readU32());
    const colCount = r.readU32();
    for (let i = 0; i < colCount; i++) snap.hiddenCols.push(r.readU32());
  }

  private _decodeDAG(r: BinaryReader): DagEdge[] {
    const count = r.readU32();
    const edges: DagEdge[] = [];
    for (let i = 0; i < count; i++) {
      const row      = r.readU32();
      const col      = r.readU32();
      const depCount = r.readU32();
      const deps: Address[] = [];
      for (let d = 0; d < depCount; d++) {
        deps.push({ row: r.readU32(), col: r.readU32() });
      }
      edges.push({ row, col, deps });
    }
    return edges;
  }

  private _decodeVolatiles(r: BinaryReader): Address[] {
    const count     = r.readU32();
    const volatiles: Address[] = [];
    for (let i = 0; i < count; i++) {
      volatiles.push({ row: r.readU32(), col: r.readU32() });
    }
    return volatiles;
  }
}

// ---------------------------------------------------------------------------
// Convenience singleton
// ---------------------------------------------------------------------------

/** Shared stateless codec instance — use for common encode/decode operations. */
export const snapshotCodec = new SnapshotCodec();
