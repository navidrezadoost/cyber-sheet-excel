import type { Command } from '../CommandManager';
import type { AddPictureOptions, PictureObject } from '../DrawingLayer';
import type { WorksheetSnapshot } from '../persistence/SnapshotCodec';
import type { CellStyle } from '../types';
import type { Worksheet } from '../worksheet';

abstract class WorksheetSnapshotCommand implements Command {
  private before?: WorksheetSnapshot;
  private after?: WorksheetSnapshot;

  constructor(
    protected readonly worksheet: Worksheet,
    public readonly description: string
  ) {}

  execute(): void {
    if (this.after) {
      this.worksheet.applySnapshot(this.after);
      return;
    }

    this.before = this.worksheet.extractSnapshot();
    this.applyOperation();
    this.after = this.worksheet.extractSnapshot();
  }

  undo(): void {
    if (!this.before) {
      throw new Error(`Cannot undo "${this.description}" before execute`);
    }
    this.worksheet.applySnapshot(this.before);
  }

  protected abstract applyOperation(): void;
}

export class InsertRowsCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly rowIndex: number,
    private readonly count = 1
  ) {
    super(worksheet, `Insert ${count} row${count === 1 ? '' : 's'}`);
  }

  protected applyOperation(): void {
    this.worksheet.insertRows(this.rowIndex, this.count);
  }
}

export class InsertColumnsCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly colIndex: number,
    private readonly count = 1
  ) {
    super(worksheet, `Insert ${count} column${count === 1 ? '' : 's'}`);
  }

  protected applyOperation(): void {
    this.worksheet.insertColumns(this.colIndex, this.count);
  }
}

export class ReorderRowsCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly fromIndex: number,
    private readonly toIndex: number,
    private readonly count = 1
  ) {
    super(worksheet, `Move ${count} row${count === 1 ? '' : 's'}`);
  }

  protected applyOperation(): void {
    this.worksheet.reorderRows(this.fromIndex, this.toIndex, this.count);
  }
}

export class ReorderColumnsCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly fromIndex: number,
    private readonly toIndex: number,
    private readonly count = 1
  ) {
    super(worksheet, `Move ${count} column${count === 1 ? '' : 's'}`);
  }

  protected applyOperation(): void {
    this.worksheet.reorderColumns(this.fromIndex, this.toIndex, this.count);
  }
}

export class SetRowStyleCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly row: number,
    private readonly style: CellStyle | undefined
  ) {
    super(worksheet, 'Set row style');
  }

  protected applyOperation(): void {
    this.worksheet.setRowStyle(this.row, this.style);
  }
}

export class SetColumnStyleCommand extends WorksheetSnapshotCommand {
  constructor(
    worksheet: Worksheet,
    private readonly col: number,
    private readonly style: CellStyle | undefined
  ) {
    super(worksheet, 'Set column style');
  }

  protected applyOperation(): void {
    this.worksheet.setColumnStyle(this.col, this.style);
  }
}

export class InsertImageCommand extends WorksheetSnapshotCommand {
  private insertedImageId?: string;

  constructor(
    worksheet: Worksheet,
    private readonly options: AddPictureOptions
  ) {
    super(worksheet, 'Insert image');
  }

  get imageId(): string | undefined {
    return this.insertedImageId;
  }

  getImage(): PictureObject | undefined {
    return this.insertedImageId
      ? this.worksheet.getDrawingLayer().getObject(this.insertedImageId) as PictureObject | undefined
      : undefined;
  }

  protected applyOperation(): void {
    const image = this.worksheet.insertImage(this.cloneOptions(this.options));
    this.insertedImageId = image.id;
  }

  private cloneOptions(options: AddPictureOptions): AddPictureOptions {
    return {
      ...options,
      position: options.position ? { ...options.position } : undefined,
      size: options.size ? { ...options.size } : undefined,
      cellAnchor: options.cellAnchor ? { ...options.cellAnchor } : undefined,
    };
  }
}
