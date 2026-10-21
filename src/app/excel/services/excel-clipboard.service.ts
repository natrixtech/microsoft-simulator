import { Injectable } from '@angular/core';
import { CellData, CellFormat } from '../models/excel.types';
import {
  FormulaEngineService,
  formatAddress,
  parseAddress,
  parseRange,
} from './formula-engine.service';
import { ExcelStateService } from './excel-state.service';
import { ExcelActionLoggerService } from './excel-action-logger.service';

interface ClipboardItem {
  rowOffset: number;
  colOffset: number;
  raw: string | number;
  formula?: string;
  format?: CellFormat;
}

@Injectable({
  providedIn: 'root',
})
export class ExcelClipboardService {
  private clipboardData: ClipboardItem[] = [];
  private isCutMode = false;
  private cutRange: string | null = null;
  private sourceRow = 0;
  private sourceCol = 0;

  constructor(
    private state: ExcelStateService,
    private formulaEngine: FormulaEngineService,
    private actionLogger: ExcelActionLoggerService
  ) {}

  copy(rangeStr: string): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;

    this.clipboardData = [];
    this.isCutMode = false;
    this.cutRange = null;

    this.sourceRow = parsed.startRow;
    this.sourceCol = parsed.startCol;
    const baseRow = parsed.startRow;
    const baseCol = parsed.startCol;

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const cell = this.state.getCell(addr);
        const fmt = this.state.cellFormats().get(addr);

        if (cell || fmt) {
          this.clipboardData.push({
            rowOffset: r - baseRow,
            colOffset: c - baseCol,
            raw: cell?.raw ?? '',
            formula: cell?.formula,
            format: fmt
              ? {
                  ...fmt,
                  border: fmt.border ? { ...fmt.border } : undefined,
                }
              : undefined,
          });
        }
      }
    }

    this.actionLogger.log('copy', rangeStr);
  }

  cut(rangeStr: string): void {
    this.copy(rangeStr);
    this.isCutMode = true;
    this.cutRange = rangeStr;
    this.actionLogger.log('cut', rangeStr);
  }

  paste(targetAddr: string): void {
    if (this.clipboardData.length === 0) return;

    const parsedTarget = parseAddress(targetAddr);
    if (!parsedTarget) return;

    const targetRow = parsedTarget.row;
    const targetCol = parsedTarget.col;
    const rowDelta = targetRow - this.sourceRow;
    const colDelta = targetCol - this.sourceCol;

    // If it was cut, clear original cells and formats before writing destination so overlapping cut/paste never wipes pasted data
    if (this.isCutMode && this.cutRange) {
      this.state.clearCellsAndFormats(this.cutRange);
    }

    // Apply clipboard items shifted by targetRow/Col
    for (const item of this.clipboardData) {
      const r = targetRow + item.rowOffset;
      const c = targetCol + item.colOffset;
      if (r >= this.state.maxRows || c >= this.state.maxCols) continue;

      const destAddr = formatAddress({ row: r, col: c });

      let rawToSet: string | number = item.raw;
      if (item.formula) {
        // Cut moves formulas without shifting references; Copy shifts references by rowDelta and colDelta
        rawToSet = this.isCutMode
          ? item.formula
          : this.formulaEngine.shiftFormula(item.formula, rowDelta, colDelta);
      }

      this.state.setCellValue(destAddr, rawToSet);

      if (item.format) {
        this.state.setRangeFormat(destAddr, item.format);
      }
    }

    if (this.isCutMode) {
      this.isCutMode = false;
      this.cutRange = null;
    }

    this.actionLogger.log('paste', targetAddr);
  }
}
