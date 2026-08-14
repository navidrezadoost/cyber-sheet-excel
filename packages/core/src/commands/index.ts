/**
 * commands/index.ts
 *
 * Central export point for all command classes
 */

// Data Commands
export {
  SortCommand,
  ToggleAutoFilterCommand,
  ClearFilterCommand,
  SetDataValidationCommand,
  ClearDataValidationCommand,
  RemoveDuplicatesCommand,
  TextToColumnsCommand,
  GroupOutlineCommand,
  UngroupOutlineCommand,
  AutoSumCommand,
} from './DataCommands';

export type {
  Range,
  SortLevel,
  DataValidationRule,
  FilterState,
  AutoSumPlan,
} from './DataCommands';

// View Commands
export {
  FreezePanesCommand,
  SplitWindowCommand,
  SetZoomCommand,
  ZoomToSelectionCommand,
  SetViewModeCommand,
  ToggleShowOptionCommand,
  HideWindowCommand,
  NewWindowCommand,
  ArrangeWindowsCommand,
} from './ViewCommands';

export type {
  FreezePanesState,
  SplitState,
  ViewMode,
  ShowOptions,
} from './ViewCommands';

// Review Commands
export {
  AddCommentCommand,
  DeleteCommentCommand,
  ToggleCommentsVisibilityCommand,
  ProtectSheetCommand,
  UnprotectSheetCommand,
  ProtectWorkbookCommand,
  UnprotectWorkbookCommand,
  SetAllowEditRangeCommand,
  RemoveAllowEditRangeCommand,
  SpellCheckCommand,
  ToggleTrackChangesCommand,
  AcceptChangeCommand,
  RejectChangeCommand,
} from './ReviewCommands';

export type {
  Comment,
  SheetProtection,
  WorkbookProtection,
  AllowEditRange,
  TrackedChange,
} from './ReviewCommands';

// Drawing Commands
export {
  DeleteDrawingObjectsCommand,
  CopyDrawingObjectsCommand,
  MoveDrawingObjectsCommand,
  ResizeDrawingObjectCommand,
  RotateDrawingObjectCommand,
  AddDrawingObjectCommand,
  GroupDrawingObjectsCommand,
} from './DrawingCommands';
