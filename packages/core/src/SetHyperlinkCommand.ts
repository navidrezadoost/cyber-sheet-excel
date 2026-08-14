import type { Command } from './CommandManager';
import type { Address, CellHyperlink, CellStyle, ExtendedCellValue } from './types';
import type { Worksheet } from './worksheet';
import { getDefaultHyperlinkDisplayText, HYPERLINK_COLOR } from './hyperlinkUtils';

export interface SetHyperlinkOptions {
  displayText?: string;
}

function cloneStyle(style?: CellStyle): CellStyle | undefined {
  return style ? { ...style } : undefined;
}

function stripHyperlinkFormatting(style?: CellStyle): CellStyle | undefined {
  if (!style) return undefined;
  const next = { ...style };
  if (next.color === HYPERLINK_COLOR || next.color === '#0563C1') {
    delete next.color;
  }
  if (next.underline === 'single' || next.underline === true) {
    delete next.underline;
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

export class SetHyperlinkCommand implements Command {
  readonly description: string;

  private addr: Address;
  private previousHyperlink?: CellHyperlink;
  private previousValue: ExtendedCellValue;
  private previousStyle?: CellStyle;
  private newHyperlink: CellHyperlink | null;
  private displayText?: string;

  constructor(
    private worksheet: Worksheet,
    addr: Address,
    hyperlink: CellHyperlink | null,
    options?: SetHyperlinkOptions
  ) {
    this.addr = { row: addr.row, col: addr.col };
    this.newHyperlink = hyperlink ? { ...hyperlink } : null;
    this.displayText = options?.displayText;

    const cell = worksheet.getCell(addr);
    this.previousHyperlink = cell?.hyperlink ? { ...cell.hyperlink } : undefined;
    this.previousValue = worksheet.getCellValue(addr);
    this.previousStyle = cloneStyle(worksheet.getDirectCellStyle(addr));

    this.description = hyperlink
      ? `Set hyperlink on (${addr.row},${addr.col})`
      : `Remove hyperlink from (${addr.row},${addr.col})`;
  }

  execute(): void {
    if (this.newHyperlink) {
      const styleBeforeLink = cloneStyle(this.previousStyle);
      this.worksheet.setHyperlink(this.addr, {
        ...this.newHyperlink,
        previousStyle: styleBeforeLink,
      });
      const displayValue = this.displayText?.trim()
        || getDefaultHyperlinkDisplayText(this.newHyperlink);
      if (displayValue) {
        this.worksheet.setCellValue(this.addr, displayValue);
      }
      const baseStyle = this.worksheet.getDirectCellStyle(this.addr) ?? styleBeforeLink ?? {};
      this.worksheet.setCellStyle(this.addr, {
        ...baseStyle,
        color: HYPERLINK_COLOR,
        underline: 'single',
      });
      return;
    }

    const existing = this.worksheet.getCell(this.addr)?.hyperlink;
    const restoreStyle = cloneStyle(existing?.previousStyle)
      ?? stripHyperlinkFormatting(this.worksheet.getDirectCellStyle(this.addr));

    this.worksheet.setHyperlink(this.addr, undefined);
    this.worksheet.setCellStyle(this.addr, restoreStyle);
  }

  undo(): void {
    this.worksheet.setHyperlink(this.addr, this.previousHyperlink);
    this.worksheet.setCellValue(this.addr, this.previousValue);
    this.worksheet.setCellStyle(this.addr, this.previousStyle);
  }
}
