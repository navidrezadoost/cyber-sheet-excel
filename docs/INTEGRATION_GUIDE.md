# CyberSheet Integration Guide

This guide documents the embedder-facing API shipped in **Phase 6** (events, comments, custom components, centralised configuration, drag-and-drop loading). Use it as the single source of truth when integrating CyberSheet into your application.

## Packages

```bash
npm install @cyber-sheet/core @cyber-sheet/react @cyber-sheet/renderer-canvas @cyber-sheet/io-xlsx
```

| Package | Role |
|---------|------|
| `@cyber-sheet/core` | `Workbook`, `Worksheet`, `CommandManager`, events, formulas |
| `@cyber-sheet/react` | `ExcelApp`, `CyberSheet`, config, UI shell |
| `@cyber-sheet/renderer-canvas` | Canvas renderer (used internally; advanced use only) |
| `@cyber-sheet/io-xlsx` | `loadXlsxFromArrayBuffer`, `exportXLSX`, `loadXlsxFromUrl` |

---

## Quick Start

Minimal integration: controlled workbook, configuration, file loading, and events.

```tsx
import React, { useCallback, useMemo, useState } from 'react';
import { Workbook } from '@cyber-sheet/core';
import { ExcelApp } from '@cyber-sheet/react';
import '@cyber-sheet/react/dist/excel-app.css';

export function SpreadsheetApp() {
  const [workbook, setWorkbook] = useState(() => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    return wb;
  });

  const handleEvent = useCallback((event: string, payload: unknown) => {
    if (event === 'cell-value-change') console.log('Cell changed:', payload);
  }, []);

  return (
    <ExcelApp
      workbook={workbook}
      fileName="Report.xlsx"
      onWorkbookLoaded={setWorkbook}
      onEvent={handleEvent}
      config={{
        enabledTabs: ['home', 'data', 'review'],
        allowEdit: true,
        enableComments: true,
        authorName: 'Jane Doe',
      }}
    />
  );
}
```

**Required for file open / drag-and-drop:** pass `onWorkbookLoaded` so the parent can replace the `workbook` prop when a file is loaded.

**Coordinates:** cell addresses use **1-based** `{ row, col }` indices (Excel-style). Row 1, column A is `{ row: 1, col: 1 }`.

---

## Configuration Reference

Pass `config` to `ExcelApp`. All fields are optional; unset fields use defaults from `DEFAULT_APP_CONFIG`.

### App flags

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `allowOpen` | `boolean` | `true` | File tab, backstage open, drag-and-drop, Ctrl+O / Ctrl+N |
| `allowSave` | `boolean` | `true` | Save shortcuts and backstage save flows |
| `allowExport` | `boolean` | `true` | Backstage export panel |
| `defaultFormat` | `'xlsx' \| 'csv'` | `'xlsx'` | Default export format hint |
| `showRibbon` | `boolean` | `true` | Ribbon tabs and ribbon content |
| `showFormulaBar` | `boolean` | `true` | Formula bar below ribbon |
| `showSheetTabs` | `boolean` | `true` | Sheet tab strip |
| `showStatusBar` | `boolean` | `true` | Status bar with zoom |
| `showContextMenu` | `boolean` | `true` | Right-click context menu and mini toolbar |
| `allowEdit` | `boolean` | `true` | Cell edits, paste, undo/redo, structural commands |
| `allowFormatting` | `boolean` | `true` | Style, border, merge, conditional formatting commands |
| `allowFormulaEdit` | `boolean` | `true` | Formula entry (values starting with `=`) |
| `enableComments` | `boolean` | `true` | Comments UI, panel, and comment commands |
| `showCommentPanel` | `boolean` | `false` | Comment sidebar open on initial render |
| `enableCustomComponents` | `boolean` | `true` | Custom cell icons and React portal overlays |
| `authorName` | `string` | `'You'` | Default author for new comments |
| `authorAvatar` | `string` | — | Optional avatar URL or data URI for new comments |
| `eventFilter` | `(event: string) => boolean` | — | When set, only matching events are forwarded to `onEvent` |

### Ribbon visibility

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `enabledTabs` | `CyberSheetTabId[]` | all tabs | Whitelist of ribbon tabs to show |
| `enabledGroups` | `CyberSheetHomeGroupsConfig` | all `true` | Home tab group visibility |

**Tab IDs:** `'home'`, `'insert'`, `'pageLayout'`, `'formulas'`, `'data'`, `'review'`, `'view'`, `'automate'`, `'help'`.

**Home group IDs:** `clipboard`, `font`, `alignment`, `number`, `styles`, `cells`, `editing`.

Example — data-only viewer with formatting disabled:

```tsx
config={{
  enabledTabs: ['home', 'data'],
  enabledGroups: { clipboard: true, font: false, alignment: false, number: true, styles: false, cells: false, editing: true },
  allowEdit: false,
  allowFormatting: false,
  showContextMenu: false,
}}
```

### Fonts

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `fonts.families` | `string[]` | Calibri, Arial, … | Font family dropdown options |
| `fonts.customFonts` | `CyberSheetFontDefinition[]` | `[]` | `@font-face` definitions loaded at runtime |
| `fonts.defaultFamily` | `string` | `'Calibri'` | Default cell font |
| `fonts.defaultSize` | `number` | `11` | Default cell font size (pt) |
| `fonts.sizes` | `number[]` | 8–72 | Font size dropdown options |

See also [GLOBAL_SETTINGS.md](./guides/GLOBAL_SETTINGS.md) for module-level `configureCyberSheet()`.

### React hooks

When building custom UI inside `CyberSheetConfigProvider`:

```tsx
import { useCyberSheetAppConfig, useCyberSheetResolvedConfig } from '@cyber-sheet/react';

const appConfig = useCyberSheetAppConfig();       // app flags only
const resolved = useCyberSheetResolvedConfig();   // app flags + fonts
```

`ExcelApp` wraps its tree in `CyberSheetConfigProvider` automatically.

### Command permissions

`ExcelApp` uses a guarded `CommandManager`. Commands blocked by config are silently dropped (no throw). Direct sheet mutations (in-cell edit, formula bar) respect the same rules via `canEditCellValue` / `canStartCellEdit`.

| Config | Blocks |
|--------|--------|
| `allowEdit: false` | Value edits, paste, clear, insert/delete, sort, hyperlinks, filters, custom components, undo/redo |
| `allowFormatting: false` | Style, merge, batch format commands |
| `allowFormulaEdit: false` | Values starting with `=` |
| `enableComments: false` | Comment-related commands |
| `enableCustomComponents: false` | `SetCellComponentCommand` |

---

## Event API

### Lazy bridging

CyberSheet uses a **lazy event bridge**: worksheet and renderer listeners are wired only when the first subscriber calls `.on(event)`. When the last subscriber unsubscribes, the bridge deactivates. This keeps idle embeds cheap.

Bridges are registered on `workbook.eventBus` at sheet creation and on the canvas renderer when `CyberSheet` mounts.

### Subscribing from React

**Option A — `ExcelApp` prop (simplest):**

```tsx
<ExcelApp
  onEvent={(event, payload) => { /* ... */ }}
  config={{ eventFilter: (e) => e.startsWith('cell-') }}
/>
```

**Option B — `CyberSheet` ref (headless / custom shell):**

```tsx
const ref = useRef<CyberSheetHandle>(null);

useEffect(() => {
  const unsub = ref.current?.on('selection-change', (payload) => {
    console.log(payload);
  });
  return () => unsub?.();
}, []);

<CyberSheet ref={ref} workbook={workbook} commandManager={commandManager} />
```

**Option C — direct bus access:**

```tsx
const unsub = workbook.eventBus.on('cell-value-change', handler);
// later: unsub();
```

Each `.on()` returns an unsubscribe function.

### Event catalog

| Event | Source | Payload (summary) |
|-------|--------|-------------------|
| `cell-value-change` | Worksheet | `{ row, col, oldValue, newValue, formula? }` |
| `cell-formula-change` | Worksheet | `{ row, col, formula }` |
| `cell-style-change` | Worksheet | `{ row, col, style }` |
| `cell-hover` | Worksheet | `{ row, col }` |
| `cell-hover-end` | Worksheet | `{ row, col }` |
| `cell-mousedown` | Worksheet | `{ row, col }` |
| `cell-dblclick` | Worksheet | `{ row, col }` |
| `comment-add` | Worksheet | `{ row, col, commentId, author, text }` |
| `comment-edit` | Worksheet | `{ row, col, commentId, author, text }` |
| `comment-delete` | Worksheet | `{ row, col, commentId }` |
| `selection-change` | Renderer | `{ ranges, activeCell }` |
| `scroll` | Renderer | `{ scrollTop, scrollLeft }` |
| `scroll-end` | Renderer | `{ scrollTop, scrollLeft }` (debounced ~150ms) |
| `command-execute` | CommandManager | `{ commandType, description?, command }` |
| `command-undo` | CommandManager | `{ commandType, description?, command }` |
| `command-redo` | CommandManager | `{ commandType, description?, command }` |

Events forwarded by `onEvent` match this list. Custom events may be added in future releases; subscribe only to documented names for stability.

---

## Custom Cell Components

Embedders can attach **icons** (canvas-drawn) or **React components** (portaled over the grid) to individual cells.

### Types

```typescript
type CustomCellComponent = {
  type: 'icon' | 'react-component';
  id: string;                              // registry key for react-component
  props?: Record<string, unknown>;
  position?: 'left' | 'right' | 'overlay'; // icon placement
  size?: number;                           // icon size in px
};
```

### Registration and assignment

```tsx
import { useRef, useEffect } from 'react';
import { CyberSheet, type CyberSheetHandle, type CustomCellComponentRenderProps } from '@cyber-sheet/react';

function StatusBadge({ row, col, props }: CustomCellComponentRenderProps) {
  return (
    <span style={{ fontSize: 10, padding: '2px 6px', background: '#e0e0e0', borderRadius: 4 }}>
      {String(props?.label ?? 'OK')}
    </span>
  );
}

function App() {
  const sheetRef = useRef<CyberSheetHandle>(null);

  useEffect(() => {
    sheetRef.current?.registerComponent('status-badge', StatusBadge);
  }, []);

  const pinIcon = () => {
    sheetRef.current?.setCellComponent(
      { row: 2, col: 3 },
      { type: 'icon', id: 'pin', position: 'right', size: 14 },
    );
  };

  const attachBadge = () => {
    sheetRef.current?.setCellComponent(
      { row: 2, col: 3 },
      { type: 'react-component', id: 'status-badge', props: { label: 'Review' }, position: 'overlay' },
    );
  };

  return <CyberSheet ref={sheetRef} workbook={workbook} commandManager={cm} enableCustomComponents />;
}
```

### Handle methods

| Method | Description |
|--------|-------------|
| `registerComponent(id, component)` | Register a React component for `type: 'react-component'` |
| `unregisterComponent(id)` | Remove registration |
| `listRegisteredComponents()` | List registered IDs |
| `setCellComponent(address, component \| null)` | Assign or update (undoable when `commandManager` is provided) |
| `getCellComponent(address)` | Read descriptor |
| `clearCellComponent(address)` | Remove component |

Set `config.enableCustomComponents: false` or `enableCustomComponents={false}` on `CyberSheet` to disable.

**Undo:** assignments through `CyberSheetHandle.setCellComponent` with a `commandManager` use `SetCellComponentCommand` and participate in undo/redo.

---

## Comments

### Built-in UI

- **Review ribbon** → Open/Close Comments Panel
- **Comment panel** — sidebar with sheet filter, threads, replies, resolve/reopen
- **Cell comment dialog** — edit or delete the latest comment on a cell
- **Shift+F2** — new comment on selected cell (when `enableComments` is true)

Configure author identity:

```tsx
config={{
  enableComments: true,
  showCommentPanel: true,
  authorName: 'Jane Doe',
  authorAvatar: 'https://example.com/avatar.png',
}}
```

### Programmatic API (Worksheet)

```typescript
// Add root comment
sheet.addComment({ row: 5, col: 3 }, {
  text: 'Please verify',
  author: 'Jane Doe',
  authorAvatar: 'https://…',
});

// Reply (threaded)
sheet.addComment({ row: 5, col: 3 }, {
  text: 'Done',
  author: 'Bob',
  parentId: rootComment.id,
});

// Update / resolve
sheet.updateComment({ row: 5, col: 3 }, commentId, { text: 'Updated', resolved: true });

// Delete
sheet.deleteComment({ row: 5, col: 3 }, commentId, /* emitEvent */ true);

// Query
sheet.getComments({ row: 5, col: 3 });
sheet.getNextCommentCell({ row: 1, col: 1 }, 'next' | 'prev');
```

Subscribe to `comment-add`, `comment-edit`, `comment-delete` for live sync with external systems. The comment panel uses lazy bridging and only subscribes while open.

See [COMMENTS_API.md](./api/COMMENTS_API.md) and [QUICK_START_COMMENTS.md](./guides/QUICK_START_COMMENTS.md) for import/export details.

---

## File I/O

### Open

| Method | Requires | Notes |
|--------|----------|-------|
| File → Open (backstage) | `allowOpen`, `onWorkbookLoaded` | Browser file picker |
| Drag-and-drop on grid | `allowOpen`, `onWorkbookLoaded` | `.xlsx` / `.xls` (parsed as XLSX) |
| Programmatic | — | `loadXlsxFromArrayBuffer(new Uint8Array(buffer))` |

```typescript
import { loadXlsxFromArrayBuffer, loadXlsxFromUrl } from '@cyber-sheet/io-xlsx';

const wb = loadXlsxFromArrayBuffer(new Uint8Array(await file.arrayBuffer()));
// or
const wb = await loadXlsxFromUrl('/reports/q1.xlsx');
setWorkbook(wb);
```

### Export (backstage Export panel)

| Format | Status | Notes |
|--------|--------|-------|
| **XLSX** | ✅ Supported | `exportXLSX(workbook)` from `@cyber-sheet/io-xlsx` |
| **CSV** | ✅ Supported | Active sheet, basic escaping |
| **PDF** | ⚠️ Print-based | Shows message to use browser Print → Save as PDF (Ctrl+P) |
| **ODS** | ❌ Not yet | Throws *"ODS export coming soon"* |
| **TXT / HTML** | ❌ Stub | Not implemented |

Programmatic XLSX export:

```typescript
import { exportXLSX } from '@cyber-sheet/io-xlsx';

const buffer = await exportXLSX(workbook);
const blob = new Blob([buffer], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
});
// trigger download…
```

---

## Embedding Patterns

### Read-only viewer

```tsx
<ExcelApp
  workbook={workbook}
  config={{
    allowOpen: false,
    allowEdit: false,
    allowFormatting: false,
    allowFormulaEdit: false,
    showContextMenu: false,
    enabledTabs: ['home', 'view'],
    enabledGroups: {
      clipboard: true, font: false, alignment: false, number: false,
      styles: false, cells: false, editing: false,
    },
  }}
/>
```

### Feature-restricted editor

Allow data entry but block formatting and formulas:

```tsx
config={{
  allowEdit: true,
  allowFormatting: false,
  allowFormulaEdit: false,
  enabledTabs: ['home', 'data'],
}}
```

### Controlled workbook

Keep `Workbook` in React state (or external store). Always pass the current instance as `workbook` and use `onWorkbookLoaded` for replacements:

```tsx
const [workbook, setWorkbook] = useState(initialWorkbook);
<ExcelApp workbook={workbook} onWorkbookLoaded={setWorkbook} />
```

Mutations made inside CyberSheet mutate the same object in memory; for immutable patterns, clone or rebuild the workbook in `onWorkbookLoaded` and on save.

### Collaboration-ready setup

Use events + config to sync with a backend without forking the UI:

```tsx
<ExcelApp
  workbook={workbook}
  onWorkbookLoaded={setWorkbook}
  config={{ authorName: currentUser.name, authorAvatar: currentUser.avatarUrl }}
  onEvent={(event, payload) => {
    if (event === 'cell-value-change') {
      collaborationBus.publish('cell', payload);
    }
    if (event.startsWith('comment-')) {
      collaborationBus.publish('comment', { event, payload });
    }
  }}
/>
```

Apply remote updates by mutating the worksheet and calling `renderer.scheduleRedraw()` if you use a headless `CyberSheet` ref.

### Headless grid (custom chrome)

Use `CyberSheet` without `ExcelApp` when you provide your own ribbon, tabs, and formula bar:

```tsx
<CyberSheet
  ref={ref}
  workbook={workbook}
  commandManager={commandManager}
  onSelectionChange={setSelection}
  onRendererReady={setRenderer}
  enableCustomComponents
/>
```

Wrap with `CyberSheetConfigProvider` if you need shared config hooks in child components.

---

## API Reference

### `ExcelAppProps`

| Prop | Type | Description |
|------|------|-------------|
| `workbook` | `Workbook` | **Required.** Active workbook instance |
| `fileName` | `string` | Window title (default `'Book1'`) |
| `config` | `CyberSheetConfigInput` | Feature flags, tabs, fonts |
| `onWorkbookLoaded` | `(wb: Workbook) => void` | New workbook from open / drag-drop |
| `onEvent` | `(event, payload) => void` | Forwarded EventBus events |
| `onSave` | `() => void` | Quick-access Save |
| `onUndo` / `onRedo` | `() => void` | Optional title-bar overrides |
| `style` | `CSSProperties` | Root container style |

### `CyberSheetHandle`

Extends `EventAPI` (`on`, `once`, `off`):

| Member | Description |
|--------|-------------|
| `getWorkbook()` | Current `Workbook` |
| `getActiveSheet()` | Active `Worksheet` |
| `executeCommand(command)` | Run a `Command` through the app command manager |
| `registerComponent` / `unregisterComponent` / `listRegisteredComponents` | Custom component registry |
| `setCellComponent` / `getCellComponent` / `clearCellComponent` | Per-cell components |

### `Workbook` (public surface)

| Member | Description |
|--------|-------------|
| `eventBus` | `EventBus` for app-wide events |
| `addSheet(name, rows?, cols?)` | Create worksheet |
| `deleteSheet(name)` | Remove sheet (≥1 must remain) |
| `renameSheet(oldName, newName)` | Rename with Excel naming rules |
| `getSheet(name)` | Lookup by name |
| `getSheetNames()` | All sheet names |
| `activeSheet` | Current `Worksheet` |
| `activeSheetName` | Get/set active sheet by name |
| `setFormulaEngine(engine)` | Attach shared formula engine |
| Pivot APIs | `createPivot`, `getPivotSnapshot`, `deletePivot`, slicer helpers — see core types |

### `Worksheet` (common methods)

| Member | Description |
|--------|-------------|
| `getCell(addr)` | Read cell |
| `setCellValue(addr, value)` | Write value |
| `setCellFormula(addr, formula)` | Write formula |
| `getCellValue(addr)` / `getCellStyle(addr)` | Convenience accessors; `getCellStyle` returns the effective cell style including row/column defaults |
| `getDirectCellStyle(addr)` | Read only the cell-level style override |
| `insertRows(rowIndex, count?)` / `insertColumns(colIndex, count?)` | Insert one or more tracks and remap cells, formulas, merges, visibility, sizing, styles, and image anchors |
| `reorderRows(fromIndex, toIndex, count?)` / `reorderColumns(fromIndex, toIndex, count?)` | Move a contiguous track block for drag-and-drop style interactions |
| `setRowStyle(row, style)` / `setColumnStyle(col, style)` | Apply dynamic whole-row or whole-column defaults to existing and future cells |
| `clearRowStyle(row)` / `clearColumnStyle(col)` | Remove whole-row or whole-column defaults |
| `getDrawingLayer()` / `insertImage(options)` / `getImages()` / `deleteImage(id)` | Manage floating and cell-bound picture metadata |
| `addComment` / `getComments` / `updateComment` / `deleteComment` | Comment threads |
| `setCellComponent` / `getCellComponent` / `clearCellComponent` / `getAllCellComponents` | Custom components |

Structural operations update formulas with the same tokenized reference pipeline used by copy/paste, then rebuild dependency metadata from the rewritten formulas. Cell-bound images carry `cellAnchor` metadata and move when row/column tracks are inserted or reordered.

### `CommandManager`

| Member | Description |
|--------|-------------|
| `execute(command)` | Run and push to undo stack |
| `undo()` / `redo()` | History navigation |
| `canUndo()` / `canRedo()` | Availability |
| `subscribe(listener)` | Notify on history changes |
| `clear()` | Clear undo/redo stacks |
| `getHistorySize()` | `{ undoCount, redoCount }` |
| `getUndoHistory()` | Debug descriptions |
| `setEventBus(eventBus)` | Attach bus for command events |

Implement custom undoable operations by creating classes that satisfy the `Command` interface (`execute`, `undo`, optional `description`).

### Worksheet Structure Commands

The core package exports snapshot-backed commands for structural operations:

| Command | Description |
|---------|-------------|
| `InsertRowsCommand` / `InsertColumnsCommand` | Undoable track insertion |
| `ReorderRowsCommand` / `ReorderColumnsCommand` | Undoable drag/drop-style track movement |
| `SetRowStyleCommand` / `SetColumnStyleCommand` | Undoable dynamic row/column styling |
| `InsertImageCommand` | Undoable picture insertion |

```typescript
import {
  CommandManager,
  InsertRowsCommand,
  InsertImageCommand,
  Worksheet,
} from '@cyber-sheet/core';

const sheet = new Worksheet('Sheet1');
const history = new CommandManager(100, sheet);

history.execute(new InsertRowsCommand(sheet, 2, 3));
history.execute(new InsertImageCommand(sheet, {
  source: 'data:image/png;base64,...',
  sourceType: 'dataUri',
  placement: 'cell',
  cellAnchor: { row: 2, col: 1 },
  size: { width: 120, height: 80 },
}));
```

---

## Stability Notes

- Treat this document and the exported TypeScript types as the **integration contract**.
- Prefer `config`, `onEvent`, and `CyberSheetHandle` over private component props.
- PDF and ODS export status is intentionally documented as incomplete; do not rely on those paths in production until Phase 6.6 / 6.7 land.
- For deeper topics, see [GLOBAL_SETTINGS.md](./guides/GLOBAL_SETTINGS.md), [COMMENTS_API.md](./api/COMMENTS_API.md), and [KEYBOARD_SHORTCUTS.md](./KEYBOARD_SHORTCUTS.md).
