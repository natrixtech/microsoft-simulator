import { Injectable, computed, signal } from '@angular/core';
import {
  BorderAction,
  CellBorderEdges,
  CellCoord,
  CellData,
  CellFormat,
  ChartItem,
  ConditionalFormatRule,
  DataValidationRule,
  NumberFormatType,
  PivotTableItem,
  SheetFilter,
  SheetSnapshot,
  SheetTabInfo,
  SortOrder,
  TextAlign,
} from '../models/excel.types';
import {
  FormulaEngineService,
  colLetterToIndex,
  formatAddress,
  formatIsoDate,
  indexToColLetter,
  parseAddress,
  parseDateParts,
  parseRange,
} from './formula-engine.service';
import { COLORS } from 'shared';

@Injectable({
  providedIn: 'root',
})
export class ExcelStateService {
  readonly maxRows = 50;
  readonly maxCols = 26;

  // Cells data store: map of "A1" -> CellData
  readonly cells = signal<Map<string, CellData>>(new Map());

  // Formats per cell: map of "A1" -> CellFormat
  readonly cellFormats = signal<Map<string, CellFormat>>(new Map());

  // Column widths: map of "A" -> width in px
  readonly columnWidths = signal<Map<string, number>>(new Map());

  // Row heights: map of row index (0-indexed) -> height in px
  readonly rowHeights = signal<Map<number, number>>(new Map());

  // Selection state
  readonly activeCell = signal<string>('A1');
  readonly selectedRange = signal<string>('A1');
  readonly selectionCoords = computed(() => parseRange(this.selectedRange()) || { startRow: 0, startCol: 0, endRow: 0, endCol: 0 });

  // Editing state
  readonly isEditing = signal<boolean>(false);
  readonly editingAddress = signal<string>('A1');
  readonly editInput = signal<string>('');

  // Features
  readonly charts = signal<ChartItem[]>([]);
  readonly comments = signal<Map<string, string>>(new Map());
  readonly pivotTables = signal<PivotTableItem[]>([]);
  readonly filter = signal<SheetFilter | null>(null);
  readonly conditionalFormats = signal<ConditionalFormatRule[]>([]);
  readonly dataValidations = signal<DataValidationRule[]>([]);

  // Selected chart (for delete / move)
  readonly selectedChartId = signal<string | null>(null);

  // Selected pivot table (for delete / move)
  readonly selectedPivotTableId = signal<string | null>(null);

  // Multiple Sheets
  readonly sheets = signal<SheetTabInfo[]>([{ id: 'sheet1', name: 'Sheet1' }]);
  readonly activeSheetId = signal<string>('sheet1');
  private sheetStore = new Map<string, SheetSnapshot>();

  constructor(private formulaEngine: FormulaEngineService) {}

  /**
   * Initializes state from startingSheet object.
   */
  initFromStartingSheet(sheet: any): void {
    this.sheets.set([{ id: 'sheet1', name: 'Sheet1' }]);
    this.activeSheetId.set('sheet1');
    this.sheetStore.clear();
    this.loadSheetData(sheet);
  }

  loadSheetData(sheet: any): void {
    const newCells = new Map<string, CellData>();
    const newFormats = new Map<string, CellFormat>();
    const newWidths = new Map<string, number>();

    if (sheet?.columnWidths && typeof sheet.columnWidths === 'object') {
      for (const [col, w] of Object.entries(sheet.columnWidths)) {
        newWidths.set(col.toUpperCase(), Number(w));
      }
    }
    this.columnWidths.set(newWidths);

    const newHeights = new Map<number, number>();
    if (sheet?.rowHeights && typeof sheet.rowHeights === 'object') {
      for (const [r, h] of Object.entries(sheet.rowHeights)) {
        const rowIdx = Number(r);
        if (!isNaN(rowIdx)) {
          newHeights.set(rowIdx, Number(h));
        }
      }
    }
    this.rowHeights.set(newHeights);

    // Load initial formats
    if (sheet?.formats && typeof sheet.formats === 'object') {
      for (const [rangeKey, fmt] of Object.entries(sheet.formats)) {
        const parsed = parseRange(rangeKey);
        if (parsed) {
          for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
              const addr = formatAddress({ row: r, col: c });
              const existing = newFormats.get(addr) || {};
              newFormats.set(addr, {
                ...existing,
                ...(fmt as CellFormat),
                border: (fmt as CellFormat).border ? { ...(fmt as CellFormat).border } : undefined,
              });
            }
          }
        }
      }
    }
    this.cellFormats.set(newFormats);

    // Load initial cells
    if (sheet?.cells && typeof sheet.cells === 'object') {
      for (const [addr, raw] of Object.entries(sheet.cells)) {
        const norm = addr.toUpperCase().replace(/\$/g, '');
        const rawVal = raw as string | number;
        const fmt = newFormats.get(norm);

        if (typeof rawVal === 'number') {
          // JSON numbers stay numbers
          newCells.set(norm, {
            raw: rawVal,
            value: rawVal,
            displayValue: this.formulaEngine.formatDisplayValue(rawVal, fmt?.numberFormat, fmt?.decimals),
            format: fmt,
          });
        } else if (typeof rawVal === 'string') {
          const trimmed = rawVal.trim();
          if (trimmed.startsWith('=')) {
            newCells.set(norm, {
              raw: rawVal,
              formula: trimmed,
              value: '',
              displayValue: '',
              format: fmt,
            });
          } else {
            // Check if it's a date string (YYYY-MM-DD or M/D/YYYY)
            const dateParts = parseDateParts(trimmed);
            if (dateParts) {
              const iso = formatIsoDate(dateParts.year, dateParts.month, dateParts.day);
              newCells.set(norm, {
                raw: rawVal,
                value: iso,
                displayValue: this.formulaEngine.formatDisplayValue(iso, fmt?.numberFormat, fmt?.decimals),
                format: fmt,
              });
            } else {
              // JSON strings stay strings (e.g. "101" remains text and left-aligned)
              newCells.set(norm, {
                raw: rawVal,
                value: rawVal,
                displayValue: this.formulaEngine.formatDisplayValue(rawVal, fmt?.numberFormat, fmt?.decimals),
                format: fmt,
              });
            }
          }
        }
      }
    }
    this.cells.set(newCells);

    // Reset overlay elements unless provided
    this.charts.set([]);
    this.comments.set(new Map());
    this.pivotTables.set([]);
    this.filter.set(null);
    this.conditionalFormats.set([]);
    this.dataValidations.set([]);
    this.activeCell.set('A1');
    this.selectedRange.set('A1');
    this.isEditing.set(false);
    this.selectedChartId.set(null);

    // Recalculate all formulas
    this.recalculateAll();
  }

  /**
   * Recalculates all cells in the sheet.
   */
  recalculateAll(): void {
    const curCells = new Map(this.cells());
    const curFormats = this.cellFormats();
    const evaluating = new Set<string>();

    const evaluateCell = (normAddr: string): any => {
      const cell = curCells.get(normAddr);
      if (!cell) return 0;
      if (!cell.formula) {
        return cell.value;
      }

      if (evaluating.has(normAddr)) {
        cell.value = '#REF!';
        cell.error = '#REF!';
        cell.displayValue = '#REF!';
        return '#REF!';
      }

      if (cell.value !== undefined && cell.value !== '') {
        return cell.value;
      }

      evaluating.add(normAddr);
      try {
        const evalRes = this.formulaEngine.evaluate(
          cell.formula,
          normAddr,
          getCellValue,
          getRangeValues
        );
        cell.value = evalRes.value;
        cell.error = evalRes.error;
        const fmt = curFormats.get(normAddr);
        cell.displayValue = this.formulaEngine.formatDisplayValue(evalRes.value, fmt?.numberFormat, fmt?.decimals);
      } finally {
        evaluating.delete(normAddr);
      }
      return cell.value;
    };

    const getCellValue = (addr: string): any => {
      const normAddr = addr.toUpperCase().replace(/\$/g, '');
      const cell = curCells.get(normAddr);
      if (!cell) return 0;
      if (cell.formula) {
        if (evaluating.has(normAddr)) {
          cell.value = '#REF!';
          cell.error = '#REF!';
          cell.displayValue = '#REF!';
          return '#REF!';
        }
        if (cell.value === undefined || cell.value === '') {
          return evaluateCell(normAddr);
        }
      }
      return cell.value;
    };

    const getRangeValues = (rangeStr: string): any[][] => {
      const parsed = parseRange(rangeStr);
      if (!parsed) return [['#REF!']];
      const result: any[][] = [];
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        const rowVals: any[] = [];
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          const addr = formatAddress({ row: r, col: c });
          const normAddr = addr.toUpperCase().replace(/\$/g, '');
          const cell = curCells.get(normAddr);
          if (!cell) {
            rowVals.push(null);
          } else if (cell.formula) {
            const v = getCellValue(addr);
            rowVals.push(v === '' || v === undefined || v === null ? null : v);
          } else if (cell.value === '' || cell.value === undefined || cell.value === null) {
            rowVals.push(null);
          } else {
            rowVals.push(cell.value);
          }
        }
        result.push(rowVals);
      }
      return result;
    };

    // First reset formula cell values
    for (const [addr, cell] of curCells.entries()) {
      if (cell.formula) {
        cell.value = '';
        cell.error = undefined;
      }
    }

    // Now evaluate each formula
    for (const [addr, cell] of curCells.entries()) {
      if (cell.formula) {
        if (cell.value === undefined || cell.value === '') {
          evaluateCell(addr);
        }
      } else {
        const fmt = curFormats.get(addr);
        cell.displayValue = this.formulaEngine.formatDisplayValue(cell.value, fmt?.numberFormat, fmt?.decimals);
      }
    }

    this.cells.set(curCells);
  }

  /**
   * Gets single cell data.
   */
  getCell(addr: string): CellData | undefined {
    return this.cells().get(addr.toUpperCase().replace(/\$/g, ''));
  }

  /**
   * Sets single cell raw value and triggers recalculation.
   */
  setCellValue(addr: string, rawVal: string | number): void {
    const norm = addr.toUpperCase().replace(/\$/g, '');
    const curCells = new Map(this.cells());
    const curFmt = this.cellFormats().get(norm);

    if (rawVal === '' || rawVal === null || rawVal === undefined) {
      curCells.delete(norm);
    } else {
      const parsed = this.formulaEngine.parseInput(rawVal);
      curCells.set(norm, {
        raw: parsed.raw,
        formula: parsed.formula,
        value: parsed.isFormula ? '' : parsed.value,
        displayValue: parsed.isFormula ? '' : this.formulaEngine.formatDisplayValue(parsed.value, curFmt?.numberFormat, curFmt?.decimals),
        format: curFmt,
      });
    }

    this.cells.set(curCells);
    this.recalculateAll();
  }

  /**
   * Sets format for an entire range.
   */
  setRangeFormat(rangeStr: string, updates: Partial<CellFormat>): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;

    const curFormats = new Map(this.cellFormats());
    const curCells = new Map(this.cells());

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const existing = curFormats.get(addr) || {};
        const merged: CellFormat = {
          ...existing,
          ...updates,
          border: updates.border !== undefined ? (updates.border ? { ...updates.border } : undefined) : existing.border,
        };

        if ('numberFormat' in updates && (updates.numberFormat === 'automatic' || !updates.numberFormat)) {
          delete merged.numberFormat;
          delete merged.decimals;
        }

        if ('fillColor' in updates && !updates.fillColor) {
          delete merged.fillColor;
        }

        if ('textColor' in updates && !updates.textColor) {
          delete merged.textColor;
        }

        curFormats.set(addr, merged);

        // Update cell displayValue if numberFormat or decimals changed
        const cell = curCells.get(addr);
        if (cell) {
          cell.format = merged;
          cell.displayValue = this.formulaEngine.formatDisplayValue(cell.value, merged.numberFormat, merged.decimals);
        }
      }
    }

    this.cellFormats.set(curFormats);
    this.cells.set(curCells);
  }

  /**
   * Applies borders to range.
   */
  applyBorderToRange(rangeStr: string, action: BorderAction): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;

    const curFormats = new Map(this.cellFormats());

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const existing = curFormats.get(addr) || {};
        const border: CellBorderEdges = { ...(existing.border || {}) };

        switch (action) {
          case 'noBorder':
            delete existing.border;
            break;
          case 'all':
            existing.border = { top: 'thin', bottom: 'thin', left: 'thin', right: 'thin' };
            break;
          case 'bottom':
            if (r === parsed.endRow) {
              border.bottom = 'thin';
              existing.border = border;
            }
            break;
          case 'top':
            if (r === parsed.startRow) {
              border.top = 'thin';
              existing.border = border;
            }
            break;
          case 'left':
            if (c === parsed.startCol) {
              border.left = 'thin';
              existing.border = border;
            }
            break;
          case 'right':
            if (c === parsed.endCol) {
              border.right = 'thin';
              existing.border = border;
            }
            break;
          case 'outside':
            if (r === parsed.startRow) border.top = 'thin';
            if (r === parsed.endRow) border.bottom = 'thin';
            if (c === parsed.startCol) border.left = 'thin';
            if (c === parsed.endCol) border.right = 'thin';
            existing.border = border;
            break;
          case 'thickOutside':
            if (r === parsed.startRow) border.top = 'thick';
            if (r === parsed.endRow) border.bottom = 'thick';
            if (c === parsed.startCol) border.left = 'thick';
            if (c === parsed.endCol) border.right = 'thick';
            existing.border = border;
            break;
        }

        curFormats.set(addr, { ...existing });
      }
    }

    this.cellFormats.set(curFormats);
  }

  applyBorderToSelection(action: BorderAction): void {
    this.applyBorderToRange(this.selectedRange(), action);
  }

  /**
   * Format currently selected range.
   */
  formatSelection(updates: Partial<CellFormat>): void {
    this.setRangeFormat(this.selectedRange(), updates);
  }

  /**
   * Select a single cell or range.
   */
  selectCell(addr: string): void {
    const norm = addr.toUpperCase().replace(/\$/g, '');
    this.activeCell.set(norm);
    this.selectedRange.set(norm);
    this.selectedChartId.set(null);
  }

  selectRange(rangeStr: string): void {
    this.selectedRange.set(rangeStr.toUpperCase());
    const parsed = parseRange(rangeStr);
    if (parsed) {
      const first = formatAddress({ row: parsed.startRow, col: parsed.startCol });
      this.activeCell.set(first);
    }
    this.selectedChartId.set(null);
  }

  selectColumn(colIndex: number): void {
    const letter = indexToColLetter(colIndex);
    const range = `${letter}1:${letter}${this.maxRows}`;
    this.selectRange(range);
  }

  selectRow(rowIndex: number): void {
    const rowNum = rowIndex + 1;
    const range = `A${rowNum}:${indexToColLetter(this.maxCols - 1)}${rowNum}`;
    this.selectRange(range);
  }

  selectAll(): void {
    const range = `A1:${indexToColLetter(this.maxCols - 1)}${this.maxRows}`;
    this.selectRange(range);
  }

  /**
   * Clears content of selected range.
   */
  clearSelection(): void {
    const parsed = this.selectionCoords();
    const curCells = new Map(this.cells());
    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        curCells.delete(addr);
      }
    }
    this.cells.set(curCells);
    this.recalculateAll();
  }

  /**
   * Clears both cell values and cell formats for a range.
   */
  clearCellsAndFormats(rangeStr: string): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;

    const curCells = new Map(this.cells());
    const curFormats = new Map(this.cellFormats());

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        curCells.delete(addr);
        curFormats.delete(addr);
      }
    }

    this.cells.set(curCells);
    this.cellFormats.set(curFormats);
    this.recalculateAll();
  }

  clearCellAndFormat(addr: string): void {
    this.clearCellsAndFormats(addr);
  }

  /**
   * Returns the current region (contiguous data block) around a cell,
   * bounded by blank rows and blank columns (as in Excel's CurrentRegion).
   */
  getCurrentRegion(startAddr: string): string {
    const p = parseAddress(startAddr);
    if (!p) return startAddr;

    const hasData = (r: number, c: number): boolean => {
      if (r < 0 || r >= this.maxRows || c < 0 || c >= this.maxCols) return false;
      const addr = formatAddress({ row: r, col: c });
      const cell = this.getCell(addr);
      return (
        !!cell &&
        cell.value !== '' &&
        cell.value !== null &&
        cell.value !== undefined &&
        String(cell.raw || '').trim() !== ''
      );
    };

    let minR = p.row;
    let maxR = p.row;
    let minC = p.col;
    let maxC = p.col;

    if (!hasData(p.row, p.col)) {
      let foundNeighbor = false;
      for (let dr = -1; dr <= 1; dr++) {
        for (let dc = -1; dc <= 1; dc++) {
          if (dr === 0 && dc === 0) continue;
          if (hasData(p.row + dr, p.col + dc)) {
            minR = Math.min(minR, p.row + dr);
            maxR = Math.max(maxR, p.row + dr);
            minC = Math.min(minC, p.col + dc);
            maxC = Math.max(maxC, p.col + dc);
            foundNeighbor = true;
          }
        }
      }
      if (!foundNeighbor) {
        return startAddr;
      }
    }

    let changed = true;
    while (changed) {
      changed = false;

      // Expand UP
      if (minR > 0) {
        const r = minR - 1;
        const cStart = Math.max(0, minC - 1);
        const cEnd = Math.min(this.maxCols - 1, maxC + 1);
        let rowHasData = false;
        for (let c = cStart; c <= cEnd; c++) {
          if (hasData(r, c)) {
            rowHasData = true;
            break;
          }
        }
        if (rowHasData) {
          minR = r;
          changed = true;
        }
      }

      // Expand DOWN
      if (maxR < this.maxRows - 1) {
        const r = maxR + 1;
        const cStart = Math.max(0, minC - 1);
        const cEnd = Math.min(this.maxCols - 1, maxC + 1);
        let rowHasData = false;
        for (let c = cStart; c <= cEnd; c++) {
          if (hasData(r, c)) {
            rowHasData = true;
            break;
          }
        }
        if (rowHasData) {
          maxR = r;
          changed = true;
        }
      }

      // Expand LEFT
      if (minC > 0) {
        const c = minC - 1;
        const rStart = Math.max(0, minR - 1);
        const rEnd = Math.min(this.maxRows - 1, maxR + 1);
        let colHasData = false;
        for (let r = rStart; r <= rEnd; r++) {
          if (hasData(r, c)) {
            colHasData = true;
            break;
          }
        }
        if (colHasData) {
          minC = c;
          changed = true;
        }
      }

      // Expand RIGHT
      if (maxC < this.maxCols - 1) {
        const c = maxC + 1;
        const rStart = Math.max(0, minR - 1);
        const rEnd = Math.min(this.maxRows - 1, maxR + 1);
        let colHasData = false;
        for (let r = rStart; r <= rEnd; r++) {
          if (hasData(r, c)) {
            colHasData = true;
            break;
          }
        }
        if (colHasData) {
          maxC = c;
          changed = true;
        }
      }
    }

    // Shrink blank edges if any
    while (minR <= maxR) {
      let hasInRow = false;
      for (let c = minC; c <= maxC; c++) {
        if (hasData(minR, c)) { hasInRow = true; break; }
      }
      if (!hasInRow) minR++;
      else break;
    }

    while (maxR >= minR) {
      let hasInRow = false;
      for (let c = minC; c <= maxC; c++) {
        if (hasData(maxR, c)) { hasInRow = true; break; }
      }
      if (!hasInRow) maxR--;
      else break;
    }

    while (minC <= maxC) {
      let hasInCol = false;
      for (let r = minR; r <= maxR; r++) {
        if (hasData(r, minC)) { hasInCol = true; break; }
      }
      if (!hasInCol) minC++;
      else break;
    }

    while (maxC >= minC) {
      let hasInCol = false;
      for (let r = minR; r <= maxR; r++) {
        if (hasData(r, maxC)) { hasInCol = true; break; }
      }
      if (!hasInCol) maxC--;
      else break;
    }

    if (minR > maxR || minC > maxC) return startAddr;
    if (minR === maxR && minC === maxC) return formatAddress({ row: minR, col: minC });

    return `${formatAddress({ row: minR, col: minC })}:${formatAddress({ row: maxR, col: maxC })}`;
  }

  /**
   * Sorts rows within range by column.
   */
  sortRange(rangeStr: string, byColLetter: string, order: SortOrder, hasHeader: boolean): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;

    const targetCol = colLetterToIndex(byColLetter);
    const startRow = hasHeader ? parsed.startRow + 1 : parsed.startRow;
    if (startRow >= parsed.endRow) return;

    // Collect rows
    interface RowData {
      rowIdx: number;
      sortKey: any;
      cells: Map<number, CellData | undefined>;
      formats: Map<number, CellFormat | undefined>;
    }

    const rows: RowData[] = [];
    const curCells = this.cells();
    const curFormats = this.cellFormats();

    for (let r = startRow; r <= parsed.endRow; r++) {
      const cellMap = new Map<number, CellData | undefined>();
      const fmtMap = new Map<number, CellFormat | undefined>();
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        cellMap.set(c, curCells.get(addr));
        fmtMap.set(c, curFormats.get(addr));
      }

      const sortAddr = formatAddress({ row: r, col: targetCol });
      const sortCell = curCells.get(sortAddr);
      const sortKey = sortCell?.value ?? '';

      rows.push({
        rowIdx: r,
        sortKey,
        cells: cellMap,
        formats: fmtMap,
      });
    }

    // Sort rows
    rows.sort((a, b) => {
      const k1 = a.sortKey;
      const k2 = b.sortKey;

      const d1 = parseDateParts(k1);
      const d2 = parseDateParts(k2);
      if (d1 && d2) {
        if (d1.utcTime === d2.utcTime) return 0;
        return order === 'asc' ? (d1.utcTime < d2.utcTime ? -1 : 1) : (d1.utcTime > d2.utcTime ? -1 : 1);
      }

      const n1 = Number(k1);
      const n2 = Number(k2);
      if (!isNaN(n1) && !isNaN(n2)) {
        return order === 'asc' ? n1 - n2 : n2 - n1;
      }

      const s1 = String(k1).toLowerCase();
      const s2 = String(k2).toLowerCase();
      if (s1 < s2) return order === 'asc' ? -1 : 1;
      if (s1 > s2) return order === 'asc' ? 1 : -1;
      return 0;
    });

    // Write back sorted rows
    const newCells = new Map(curCells);
    const newFormats = new Map(curFormats);

    for (let i = 0; i < rows.length; i++) {
      const targetRow = startRow + i;
      const rowData = rows[i];
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const targetAddr = formatAddress({ row: targetRow, col: c });
        const cell = rowData.cells.get(c);
        const fmt = rowData.formats.get(c);

        if (cell) {
          newCells.set(targetAddr, { ...cell });
        } else {
          newCells.delete(targetAddr);
        }

        if (fmt) {
          newFormats.set(targetAddr, { ...fmt });
        } else {
          newFormats.delete(targetAddr);
        }
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.recalculateAll();
  }

  /**
   * Removes duplicate rows within range.
   */
  removeDuplicates(rangeStr: string, hasHeader: boolean): { removedCount: number; remainingCount: number } {
    const parsed = parseRange(rangeStr);
    if (!parsed) return { removedCount: 0, remainingCount: 0 };

    const startRow = hasHeader ? parsed.startRow + 1 : parsed.startRow;
    const curCells = this.cells();
    const curFormats = this.cellFormats();

    const seenRows = new Set<string>();
    const uniqueRows: Array<{ cells: Map<number, CellData | undefined>; formats: Map<number, CellFormat | undefined> }> = [];
    let removedCount = 0;

    for (let r = startRow; r <= parsed.endRow; r++) {
      const rowKeyParts: string[] = [];
      const rowCells = new Map<number, CellData | undefined>();
      const rowFmts = new Map<number, CellFormat | undefined>();

      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const cell = curCells.get(addr);
        rowCells.set(c, cell);
        rowFmts.set(c, curFormats.get(addr));
        rowKeyParts.push(String(cell?.value ?? '').trim().toLowerCase());
      }

      const rowKey = rowKeyParts.join('||');
      if (seenRows.has(rowKey)) {
        removedCount++;
      } else {
        seenRows.add(rowKey);
        uniqueRows.push({ cells: rowCells, formats: rowFmts });
      }
    }

    if (removedCount === 0) {
      return { removedCount: 0, remainingCount: uniqueRows.length };
    }

    const newCells = new Map(curCells);
    const newFormats = new Map(curFormats);

    // Write back unique rows
    for (let i = 0; i < uniqueRows.length; i++) {
      const targetRow = startRow + i;
      const u = uniqueRows[i];
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const targetAddr = formatAddress({ row: targetRow, col: c });
        const cell = u.cells.get(c);
        const fmt = u.formats.get(c);

        if (cell) {
          newCells.set(targetAddr, { ...cell });
        } else {
          newCells.delete(targetAddr);
        }

        if (fmt) {
          newFormats.set(targetAddr, { ...fmt });
        } else {
          newFormats.delete(targetAddr);
        }
      }
    }

    // Clear removed rows at bottom of range
    for (let r = startRow + uniqueRows.length; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const targetAddr = formatAddress({ row: r, col: c });
        newCells.delete(targetAddr);
        newFormats.delete(targetAddr);
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.recalculateAll();

    return { removedCount, remainingCount: uniqueRows.length };
  }

  /**
   * Filter handling.
   */
  setFilter(filter: SheetFilter | null): void {
    this.filter.set(filter);
  }

  isRowFilteredOut(rowIdx: number): boolean {
    const f = this.filter();
    if (!f) return false;
    if (rowIdx <= f.headerRow || rowIdx > f.endRow) return false;

    const curCells = this.cells();

    for (const [colIdxStr, colFilter] of Object.entries(f.columnFilters)) {
      const colIdx = Number(colIdxStr);
      const addr = formatAddress({ row: rowIdx, col: colIdx });
      const cellVal = curCells.get(addr)?.value ?? '';

      // Check selectedValues filter (checkboxes)
      if (colFilter.selectedValues && colFilter.selectedValues.length > 0) {
        const strVal = String(cellVal).trim();
        if (!colFilter.selectedValues.includes(strVal)) {
          return true; // hidden
        }
      }

      // Check condition filter
      if (colFilter.condition) {
        const cond = colFilter.condition;
        const matches = this.evaluateFilterCondition(cellVal, cond);
        if (!matches) return true; // hidden
      }
    }

    return false;
  }

  evaluateFilterCondition(val: any, cond: { type: string; value?: any; value2?: any }): boolean {
    const d1 = parseDateParts(val);
    const d2 = parseDateParts(cond.value);
    if (d1 && d2) {
      if (cond.type === 'equals') return d1.utcTime === d2.utcTime;
      if (cond.type === 'greaterThan') return d1.utcTime > d2.utcTime;
      if (cond.type === 'lessThan') return d1.utcTime < d2.utcTime;
      if (cond.type === 'between') {
        const d3 = parseDateParts(cond.value2);
        const tMin = Math.min(d2.utcTime, d3 ? d3.utcTime : d2.utcTime);
        const tMax = Math.max(d2.utcTime, d3 ? d3.utcTime : d2.utcTime);
        return d1.utcTime >= tMin && d1.utcTime <= tMax;
      }
    }

    const str = String(val ?? '').trim().toLowerCase();
    const num = Number(val);

    switch (cond.type) {
      case 'equals':
        if (!isNaN(num) && cond.value !== undefined && !isNaN(Number(cond.value))) {
          return Math.abs(num - Number(cond.value)) < 0.0001;
        }
        return str === String(cond.value ?? '').trim().toLowerCase();

      case 'greaterThan':
        return !isNaN(num) && cond.value !== undefined && num > Number(cond.value);

      case 'lessThan':
        return !isNaN(num) && cond.value !== undefined && num < Number(cond.value);

      case 'between':
        return (
          !isNaN(num) &&
          cond.value !== undefined &&
          cond.value2 !== undefined &&
          num >= Math.min(Number(cond.value), Number(cond.value2)) &&
          num <= Math.max(Number(cond.value), Number(cond.value2))
        );

      case 'contains':
        return str.includes(String(cond.value ?? '').trim().toLowerCase());

      case 'isEmpty':
        return val === '' || val === null || val === undefined;

      default:
        return true;
    }
  }

  /**
   * Check if a cell is affected by conditional formatting rules.
   */
  getConditionalFillColor(addr: string): string | null {
    const p = parseAddress(addr);
    if (!p) return null;
    const curVal = this.getCell(addr)?.value;

    for (const rule of this.conditionalFormats()) {
      const rangeP = parseRange(rule.range);
      if (
        rangeP &&
        p.row >= rangeP.startRow &&
        p.row <= rangeP.endRow &&
        p.col >= rangeP.startCol &&
        p.col <= rangeP.endCol
      ) {
        if (this.evaluateFilterCondition(curVal, { type: rule.condition, value: rule.value, value2: rule.value2 })) {
          // Map named colors if necessary
          return COLORS[rule.fillColor] || rule.fillColor;
        }
      }
    }

    return null;
  }

  /**
   * Data validation check for cell.
   */
  getDataValidation(addr: string): DataValidationRule | null {
    const p = parseAddress(addr);
    if (!p) return null;

    for (const rule of this.dataValidations()) {
      const rangeP = parseRange(rule.range);
      if (
        rangeP &&
        p.row >= rangeP.startRow &&
        p.row <= rangeP.endRow &&
        p.col >= rangeP.startCol &&
        p.col <= rangeP.endCol
      ) {
        return rule;
      }
    }

    return null;
  }

  /**
   * Charts and PivotTables.
   */
  addChart(chart: ChartItem): void {
    this.charts.update((list) => [...list, chart]);
  }

  deleteChart(id: string): void {
    this.charts.update((list) => list.filter((c) => c.id !== id));
    if (this.selectedChartId() === id) {
      this.selectedChartId.set(null);
    }
  }

  addPivotTable(pt: PivotTableItem): void {
    this.pivotTables.update((list) => [...list, pt]);
  }

  deletePivotTable(id: string): void {
    this.pivotTables.update((list) => list.filter((p) => p.id !== id));
    if (this.selectedPivotTableId() === id) {
      this.selectedPivotTableId.set(null);
    }
  }

  movePivotTable(id: string, targetCell: string): boolean {
    const list = this.pivotTables();
    const idx = list.findIndex((p) => p.id === id);
    if (idx === -1) return false;
    const updated = { ...list[idx], targetCell: targetCell.toUpperCase() };
    const nextList = [...list];
    nextList[idx] = updated;
    this.pivotTables.set(nextList);
    return true;
  }

  setComment(cellAddr: string, text: string): void {
    const norm = cellAddr.toUpperCase().replace(/\$/g, '');
    const cur = new Map(this.comments());
    if (!text || text.trim() === '') {
      cur.delete(norm);
    } else {
      cur.set(norm, text.trim());
    }
    this.comments.set(cur);
  }

  deleteComment(cellAddr: string): void {
    const norm = cellAddr.toUpperCase().replace(/\$/g, '');
    const cur = new Map(this.comments());
    cur.delete(norm);
    this.comments.set(cur);
  }

  addConditionalFormat(rule: ConditionalFormatRule): void {
    this.conditionalFormats.update((list) => [...list, rule]);
  }

  addDataValidation(rule: DataValidationRule): void {
    this.dataValidations.update((list) => [...list, rule]);
  }

  /**
   * Creates a snapshot of the current sheet state.
  /**
   * Captures snapshot of only the active sheet's cells and features.
   */
  captureCurrentSheetOnly(): SheetSnapshot {
    const cellsObj: Record<string, string | number> = {};
    for (const [addr, c] of this.cells().entries()) {
      cellsObj[addr] = c.formula || (typeof c.value === 'number' ? c.value : c.raw);
    }

    const formatsObj: Record<string, CellFormat> = {};
    for (const [addr, f] of this.cellFormats().entries()) {
      formatsObj[addr] = { ...f, border: f.border ? { ...f.border } : undefined };
    }

    const widthsObj: Record<string, number> = {};
    for (const [col, w] of this.columnWidths().entries()) {
      widthsObj[col] = w;
    }

    const heightsObj: Record<number, number> = {};
    for (const [row, h] of this.rowHeights().entries()) {
      heightsObj[row] = h;
    }

    const commentsObj: Record<string, string> = {};
    for (const [addr, comm] of this.comments().entries()) {
      commentsObj[addr] = comm;
    }

    return {
      cells: cellsObj,
      formats: formatsObj,
      columnWidths: widthsObj,
      rowHeights: heightsObj,
      charts: this.charts().map((c) => ({ ...c })),
      comments: commentsObj,
      pivotTables: this.pivotTables().map((pt) => ({ ...pt, rows: [...pt.rows] })),
      filter: this.filter() ? { ...this.filter()!, columnFilters: { ...this.filter()!.columnFilters } } : null,
      conditionalFormats: this.conditionalFormats().map((r) => ({ ...r })),
      dataValidations: this.dataValidations().map((dv) => ({ ...dv, allowed: [...dv.allowed] })),
    };
  }

  /**
   * Restores only the active sheet's cells and features without changing workbook sheet tabs.
   */
  restoreSheetOnly(snap: SheetSnapshot): void {
    this.loadSheetData({
      cells: snap.cells,
      formats: snap.formats,
      columnWidths: snap.columnWidths,
      rowHeights: snap.rowHeights,
    });

    const commMap = new Map<string, string>();
    for (const [addr, txt] of Object.entries(snap.comments || {})) {
      commMap.set(addr, txt);
    }
    this.comments.set(commMap);

    this.charts.set((snap.charts || []).map((c) => ({ ...c })));
    this.pivotTables.set((snap.pivotTables || []).map((pt) => ({ ...pt, rows: [...pt.rows] })));
    this.filter.set(snap.filter ? { ...snap.filter, columnFilters: { ...snap.filter.columnFilters } } : null);
    this.conditionalFormats.set((snap.conditionalFormats || []).map((r) => ({ ...r })));
    this.dataValidations.set((snap.dataValidations || []).map((dv) => ({ ...dv, allowed: [...dv.allowed] })));
  }

  /**
   * Switches to another worksheet tab, preserving each sheet's state.
   */
  switchSheet(id: string): void {
    if (id === this.activeSheetId()) return;
    this.sheetStore.set(this.activeSheetId(), this.captureCurrentSheetOnly());
    this.activeSheetId.set(id);
    const existing = this.sheetStore.get(id);
    if (existing) {
      this.restoreSheetOnly(existing);
    } else {
      this.restoreSheetOnly({
        cells: {},
        formats: {},
        columnWidths: {},
        rowHeights: {},
        charts: [],
        comments: {},
        pivotTables: [],
        filter: null,
        conditionalFormats: [],
        dataValidations: [],
      });
    }
  }

  /**
   * Adds a new worksheet tab.
   */
  addSheet(): string {
    let nextNum = this.sheets().length + 1;
    let nextName = `Sheet${nextNum}`;
    while (this.sheets().some((s) => s.name.toLowerCase() === nextName.toLowerCase())) {
      nextNum++;
      nextName = `Sheet${nextNum}`;
    }
    const newId = `sheet_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
    const newSheet: SheetTabInfo = { id: newId, name: nextName };
    this.sheets.set([...this.sheets(), newSheet]);
    this.switchSheet(newId);
    return newId;
  }

  /**
   * Renames a worksheet tab.
   */
  renameSheet(id: string, newName: string): boolean {
    const clean = newName.trim();
    if (!clean) return false;
    const exists = this.sheets().some((s) => s.id !== id && s.name.toLowerCase() === clean.toLowerCase());
    if (exists) return false;
    this.sheets.set(
      this.sheets().map((s) => (s.id === id ? { ...s, name: clean } : s))
    );
    return true;
  }

  /**
   * Moves/reorders a worksheet tab from one index to another.
   */
  moveSheet(fromIndex: number, toIndex: number): void {
    if (fromIndex < 0 || fromIndex >= this.sheets().length) return;
    if (toIndex < 0 || toIndex >= this.sheets().length) return;
    if (fromIndex === toIndex) return;
    const arr = [...this.sheets()];
    const [moved] = arr.splice(fromIndex, 1);
    arr.splice(toIndex, 0, moved);
    this.sheets.set(arr);
  }

  /**
   * Deletes a worksheet tab if more than 1 sheet exists.
   */
  deleteSheet(id: string): boolean {
    if (this.sheets().length <= 1) return false;
    const curSheets = this.sheets();
    const idx = curSheets.findIndex((s) => s.id === id);
    if (idx === -1) return false;

    const remaining = curSheets.filter((s) => s.id !== id);
    this.sheetStore.delete(id);
    this.sheets.set(remaining);

    if (this.activeSheetId() === id) {
      const nextActiveIdx = Math.min(idx, remaining.length - 1);
      this.switchSheet(remaining[nextActiveIdx].id);
    }
    return true;
  }

  /**
   * Structurally inserts a row at rowIdx (0-indexed). Defaults to active cell row.
   */
  insertRow(rowIdx?: number): void {
    const targetRow = rowIdx !== undefined ? rowIdx : (parseAddress(this.activeCell())?.row ?? 0);
    if (targetRow < 0 || targetRow >= this.maxRows) return;

    const oldCells = this.cells();
    const oldFormats = this.cellFormats();
    const oldComments = this.comments();
    const oldHeights = this.rowHeights();

    const newCells = new Map<string, CellData>();
    const newFormats = new Map<string, CellFormat>();
    const newComments = new Map<string, string>();
    const newHeights = new Map<number, number>();

    for (let r = 0; r < this.maxRows; r++) {
      const h = oldHeights.get(r);
      if (h !== undefined) {
        if (r < targetRow) {
          newHeights.set(r, h);
        } else if (r + 1 < this.maxRows) {
          newHeights.set(r + 1, h);
        }
      }
    }
    this.rowHeights.set(newHeights);

    for (const [addr, cell] of oldCells.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newCells.set(addr, { ...cell });
      } else if (p.row + 1 < this.maxRows) {
        const newAddr = formatAddress({ row: p.row + 1, col: p.col });
        newCells.set(newAddr, { ...cell });
      }
    }

    for (const [addr, fmt] of oldFormats.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newFormats.set(addr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      } else if (p.row + 1 < this.maxRows) {
        const newAddr = formatAddress({ row: p.row + 1, col: p.col });
        newFormats.set(newAddr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      }
    }

    for (const [addr, comm] of oldComments.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newComments.set(addr, comm);
      } else if (p.row + 1 < this.maxRows) {
        const newAddr = formatAddress({ row: p.row + 1, col: p.col });
        newComments.set(newAddr, comm);
      }
    }

    for (const [addr, cell] of newCells.entries()) {
      if (cell.formula) {
        cell.formula = this.formulaEngine.updateFormulaOnRowColChange(cell.formula, 'insertRow', targetRow);
        cell.raw = cell.formula;
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.comments.set(newComments);

    const newCond = this.conditionalFormats().map((cf) => ({
      ...cf,
      range: this.formulaEngine.updateRangeOnRowColChange(cf.range, 'insertRow', targetRow),
    }));
    this.conditionalFormats.set(newCond);

    const newDv = this.dataValidations().map((dv) => ({
      ...dv,
      range: this.formulaEngine.updateRangeOnRowColChange(dv.range, 'insertRow', targetRow),
    }));
    this.dataValidations.set(newDv);

    const newCharts = this.charts().map((c) => ({
      ...c,
      range: this.formulaEngine.updateRangeOnRowColChange(c.range, 'insertRow', targetRow),
    }));
    this.charts.set(newCharts);

    const newPivots = this.pivotTables().map((pt) => ({
      ...pt,
      range: this.formulaEngine.updateRangeOnRowColChange(pt.range, 'insertRow', targetRow),
    }));
    this.pivotTables.set(newPivots);

    const active = parseAddress(this.activeCell());
    if (active && active.row >= targetRow) {
      this.activeCell.set(formatAddress({ row: Math.min(this.maxRows - 1, active.row + 1), col: active.col }));
    }
    const newSelRange = this.formulaEngine.updateRangeOnRowColChange(this.selectedRange(), 'insertRow', targetRow);
    if (newSelRange && newSelRange !== '#REF!') {
      this.selectedRange.set(newSelRange);
    }

    this.recalculateAll();
  }

  /**
   * Structurally deletes a row at rowIdx (0-indexed). Defaults to active cell row.
   */
  deleteRow(rowIdx?: number): void {
    const targetRow = rowIdx !== undefined ? rowIdx : (parseAddress(this.activeCell())?.row ?? 0);
    if (targetRow < 0 || targetRow >= this.maxRows) return;

    const oldCells = this.cells();
    const oldFormats = this.cellFormats();
    const oldComments = this.comments();
    const oldHeights = this.rowHeights();

    const newCells = new Map<string, CellData>();
    const newFormats = new Map<string, CellFormat>();
    const newComments = new Map<string, string>();
    const newHeights = new Map<number, number>();

    for (let r = 0; r < this.maxRows; r++) {
      const h = oldHeights.get(r);
      if (h !== undefined) {
        if (r < targetRow) {
          newHeights.set(r, h);
        } else if (r > targetRow) {
          newHeights.set(r - 1, h);
        }
      }
    }
    this.rowHeights.set(newHeights);

    for (const [addr, cell] of oldCells.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newCells.set(addr, { ...cell });
      } else if (p.row > targetRow) {
        const newAddr = formatAddress({ row: p.row - 1, col: p.col });
        newCells.set(newAddr, { ...cell });
      }
    }

    for (const [addr, fmt] of oldFormats.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newFormats.set(addr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      } else if (p.row > targetRow) {
        const newAddr = formatAddress({ row: p.row - 1, col: p.col });
        newFormats.set(newAddr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      }
    }

    for (const [addr, comm] of oldComments.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.row < targetRow) {
        newComments.set(addr, comm);
      } else if (p.row > targetRow) {
        const newAddr = formatAddress({ row: p.row - 1, col: p.col });
        newComments.set(newAddr, comm);
      }
    }

    for (const [addr, cell] of newCells.entries()) {
      if (cell.formula) {
        cell.formula = this.formulaEngine.updateFormulaOnRowColChange(cell.formula, 'deleteRow', targetRow);
        cell.raw = cell.formula;
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.comments.set(newComments);

    const newCond = this.conditionalFormats()
      .map((cf) => ({
        ...cf,
        range: this.formulaEngine.updateRangeOnRowColChange(cf.range, 'deleteRow', targetRow),
      }))
      .filter((cf) => cf.range && cf.range !== '#REF!');
    this.conditionalFormats.set(newCond);

    const newDv = this.dataValidations()
      .map((dv) => ({
        ...dv,
        range: this.formulaEngine.updateRangeOnRowColChange(dv.range, 'deleteRow', targetRow),
      }))
      .filter((dv) => dv.range && dv.range !== '#REF!');
    this.dataValidations.set(newDv);

    const newCharts = this.charts().map((c) => ({
      ...c,
      range: this.formulaEngine.updateRangeOnRowColChange(c.range, 'deleteRow', targetRow),
    }));
    this.charts.set(newCharts);

    const newPivots = this.pivotTables().map((pt) => ({
      ...pt,
      range: this.formulaEngine.updateRangeOnRowColChange(pt.range, 'deleteRow', targetRow),
    }));
    this.pivotTables.set(newPivots);

    const active = parseAddress(this.activeCell());
    if (active) {
      const nextRow = active.row > targetRow ? active.row - 1 : Math.min(active.row, this.maxRows - 1);
      this.activeCell.set(formatAddress({ row: nextRow, col: active.col }));
    }
    const newSelRange = this.formulaEngine.updateRangeOnRowColChange(this.selectedRange(), 'deleteRow', targetRow);
    if (newSelRange && newSelRange !== '#REF!') {
      this.selectedRange.set(newSelRange);
    } else {
      this.selectedRange.set(this.activeCell());
    }

    this.recalculateAll();
  }

  /**
   * Structurally inserts a column at colIdx (0-indexed). Defaults to active cell column.
   */
  insertColumn(colIdx?: number): void {
    const targetCol = colIdx !== undefined ? colIdx : (parseAddress(this.activeCell())?.col ?? 0);
    if (targetCol < 0 || targetCol >= this.maxCols) return;

    const oldCells = this.cells();
    const oldFormats = this.cellFormats();
    const oldComments = this.comments();
    const oldWidths = this.columnWidths();

    const newCells = new Map<string, CellData>();
    const newFormats = new Map<string, CellFormat>();
    const newComments = new Map<string, string>();
    const newWidths = new Map<string, number>();

    for (let c = 0; c < this.maxCols; c++) {
      const colLetter = indexToColLetter(c);
      const w = oldWidths.get(colLetter);
      if (w !== undefined) {
        if (c < targetCol) {
          newWidths.set(colLetter, w);
        } else if (c + 1 < this.maxCols) {
          newWidths.set(indexToColLetter(c + 1), w);
        }
      }
    }
    this.columnWidths.set(newWidths);

    for (const [addr, cell] of oldCells.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newCells.set(addr, { ...cell });
      } else if (p.col + 1 < this.maxCols) {
        const newAddr = formatAddress({ row: p.row, col: p.col + 1 });
        newCells.set(newAddr, { ...cell });
      }
    }

    for (const [addr, fmt] of oldFormats.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newFormats.set(addr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      } else if (p.col + 1 < this.maxCols) {
        const newAddr = formatAddress({ row: p.row, col: p.col + 1 });
        newFormats.set(newAddr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      }
    }

    for (const [addr, comm] of oldComments.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newComments.set(addr, comm);
      } else if (p.col + 1 < this.maxCols) {
        const newAddr = formatAddress({ row: p.row, col: p.col + 1 });
        newComments.set(newAddr, comm);
      }
    }

    for (const [addr, cell] of newCells.entries()) {
      if (cell.formula) {
        cell.formula = this.formulaEngine.updateFormulaOnRowColChange(cell.formula, 'insertCol', targetCol);
        cell.raw = cell.formula;
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.comments.set(newComments);

    const newCond = this.conditionalFormats().map((cf) => ({
      ...cf,
      range: this.formulaEngine.updateRangeOnRowColChange(cf.range, 'insertCol', targetCol),
    }));
    this.conditionalFormats.set(newCond);

    const newDv = this.dataValidations().map((dv) => ({
      ...dv,
      range: this.formulaEngine.updateRangeOnRowColChange(dv.range, 'insertCol', targetCol),
    }));
    this.dataValidations.set(newDv);

    const newCharts = this.charts().map((c) => ({
      ...c,
      range: this.formulaEngine.updateRangeOnRowColChange(c.range, 'insertCol', targetCol),
    }));
    this.charts.set(newCharts);

    const newPivots = this.pivotTables().map((pt) => ({
      ...pt,
      range: this.formulaEngine.updateRangeOnRowColChange(pt.range, 'insertCol', targetCol),
    }));
    this.pivotTables.set(newPivots);

    const active = parseAddress(this.activeCell());
    if (active && active.col >= targetCol) {
      this.activeCell.set(formatAddress({ row: active.row, col: Math.min(this.maxCols - 1, active.col + 1) }));
    }
    const newSelRange = this.formulaEngine.updateRangeOnRowColChange(this.selectedRange(), 'insertCol', targetCol);
    if (newSelRange && newSelRange !== '#REF!') {
      this.selectedRange.set(newSelRange);
    }

    this.recalculateAll();
  }

  /**
   * Structurally deletes a column at colIdx (0-indexed). Defaults to active cell column.
   */
  deleteColumn(colIdx?: number): void {
    const targetCol = colIdx !== undefined ? colIdx : (parseAddress(this.activeCell())?.col ?? 0);
    if (targetCol < 0 || targetCol >= this.maxCols) return;

    const oldCells = this.cells();
    const oldFormats = this.cellFormats();
    const oldComments = this.comments();
    const oldWidths = this.columnWidths();

    const newCells = new Map<string, CellData>();
    const newFormats = new Map<string, CellFormat>();
    const newComments = new Map<string, string>();
    const newWidths = new Map<string, number>();

    for (let c = 0; c < this.maxCols; c++) {
      const colLetter = indexToColLetter(c);
      const w = oldWidths.get(colLetter);
      if (w !== undefined) {
        if (c < targetCol) {
          newWidths.set(colLetter, w);
        } else if (c > targetCol) {
          newWidths.set(indexToColLetter(c - 1), w);
        }
      }
    }
    this.columnWidths.set(newWidths);

    for (const [addr, cell] of oldCells.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newCells.set(addr, { ...cell });
      } else if (p.col > targetCol) {
        const newAddr = formatAddress({ row: p.row, col: p.col - 1 });
        newCells.set(newAddr, { ...cell });
      }
    }

    for (const [addr, fmt] of oldFormats.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newFormats.set(addr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      } else if (p.col > targetCol) {
        const newAddr = formatAddress({ row: p.row, col: p.col - 1 });
        newFormats.set(newAddr, { ...fmt, border: fmt.border ? { ...fmt.border } : undefined });
      }
    }

    for (const [addr, comm] of oldComments.entries()) {
      const p = parseAddress(addr);
      if (!p) continue;
      if (p.col < targetCol) {
        newComments.set(addr, comm);
      } else if (p.col > targetCol) {
        const newAddr = formatAddress({ row: p.row, col: p.col - 1 });
        newComments.set(newAddr, comm);
      }
    }

    for (const [addr, cell] of newCells.entries()) {
      if (cell.formula) {
        cell.formula = this.formulaEngine.updateFormulaOnRowColChange(cell.formula, 'deleteCol', targetCol);
        cell.raw = cell.formula;
      }
    }

    this.cells.set(newCells);
    this.cellFormats.set(newFormats);
    this.comments.set(newComments);

    const newCond = this.conditionalFormats()
      .map((cf) => ({
        ...cf,
        range: this.formulaEngine.updateRangeOnRowColChange(cf.range, 'deleteCol', targetCol),
      }))
      .filter((cf) => cf.range && cf.range !== '#REF!');
    this.conditionalFormats.set(newCond);

    const newDv = this.dataValidations()
      .map((dv) => ({
        ...dv,
        range: this.formulaEngine.updateRangeOnRowColChange(dv.range, 'deleteCol', targetCol),
      }))
      .filter((dv) => dv.range && dv.range !== '#REF!');
    this.dataValidations.set(newDv);

    const newCharts = this.charts().map((c) => ({
      ...c,
      range: this.formulaEngine.updateRangeOnRowColChange(c.range, 'deleteCol', targetCol),
    }));
    this.charts.set(newCharts);

    const newPivots = this.pivotTables().map((pt) => ({
      ...pt,
      range: this.formulaEngine.updateRangeOnRowColChange(pt.range, 'deleteCol', targetCol),
    }));
    this.pivotTables.set(newPivots);

    const active = parseAddress(this.activeCell());
    if (active) {
      const nextCol = active.col > targetCol ? active.col - 1 : Math.min(active.col, this.maxCols - 1);
      this.activeCell.set(formatAddress({ row: active.row, col: nextCol }));
    }
    const newSelRange = this.formulaEngine.updateRangeOnRowColChange(this.selectedRange(), 'deleteCol', targetCol);
    if (newSelRange && newSelRange !== '#REF!') {
      this.selectedRange.set(newSelRange);
    } else {
      this.selectedRange.set(this.activeCell());
    }

    this.recalculateAll();
  }

  /**
   * Captures complete snapshot of current sheet state and workbook sheets for undo/redo.
   */
  getSnapshot(): SheetSnapshot {
    const currentSheetSnap = this.captureCurrentSheetOnly();

    // Store current in sheetStore
    this.sheetStore.set(this.activeSheetId(), currentSheetSnap);

    const storeDataObj: Record<string, SheetSnapshot> = {};
    for (const [sId, sSnap] of this.sheetStore.entries()) {
      storeDataObj[sId] = sSnap;
    }

    return {
      ...currentSheetSnap,
      sheets: this.sheets().map((s) => ({ ...s })),
      activeSheetId: this.activeSheetId(),
      sheetStoreData: storeDataObj,
    };
  }

  /**
   * Restores sheet state from snapshot.
   */
  restoreSnapshot(snap: SheetSnapshot): void {
    if (snap.sheets && snap.sheets.length > 0) {
      this.sheets.set(snap.sheets.map((s) => ({ ...s })));
      this.activeSheetId.set(snap.activeSheetId || snap.sheets[0].id);
      if (snap.sheetStoreData) {
        this.sheetStore.clear();
        for (const [sId, sData] of Object.entries(snap.sheetStoreData)) {
          this.sheetStore.set(sId, sData);
        }
      }
    }

    this.restoreSheetOnly(snap);
  }

  /**
   * Moves or copies a range of cells to a new location.
   * If copy is false (MOVE): source cells, formats, and comments are cleared; formulas keep original cell references.
   * If copy is true (COPY): source remains intact; formulas shift relative references via formulaEngine.shiftFormula.
   */
  moveOrCopyRange(sourceRange: string, targetTopLeft: string, copy: boolean): void {
    const parsedSource = parseRange(sourceRange);
    if (!parsedSource) return;

    const parsedTarget = parseAddress(targetTopLeft);
    if (!parsedTarget) return;

    const numRows = parsedSource.endRow - parsedSource.startRow + 1;
    const numCols = parsedSource.endCol - parsedSource.startCol + 1;
    const targetStartRow = parsedTarget.row;
    const targetStartCol = parsedTarget.col;
    const rowDelta = targetStartRow - parsedSource.startRow;
    const colDelta = targetStartCol - parsedSource.startCol;

    if (rowDelta === 0 && colDelta === 0) return;

    // 1. Gather all items from the source range
    interface SourceItem {
      rOffset: number;
      cOffset: number;
      cell?: CellData;
      format?: CellFormat;
      comment?: string;
    }

    const items: SourceItem[] = [];
    const curCells = new Map(this.cells());
    const curFormats = new Map(this.cellFormats());
    const curComments = new Map(this.comments());

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const srcRow = parsedSource.startRow + r;
        const srcCol = parsedSource.startCol + c;
        const srcAddr = formatAddress({ row: srcRow, col: srcCol });

        const cell = curCells.get(srcAddr);
        const fmt = curFormats.get(srcAddr);
        const comm = curComments.get(srcAddr);

        items.push({
          rOffset: r,
          cOffset: c,
          cell: cell ? { ...cell } : undefined,
          format: fmt ? { ...fmt, border: fmt.border ? { ...fmt.border } : undefined } : undefined,
          comment: comm,
        });
      }
    }

    // 2. If MOVE (!copy), clear the source cells, formats, and comments
    if (!copy) {
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          const srcRow = parsedSource.startRow + r;
          const srcCol = parsedSource.startCol + c;
          const srcAddr = formatAddress({ row: srcRow, col: srcCol });
          curCells.delete(srcAddr);
          curFormats.delete(srcAddr);
          curComments.delete(srcAddr);
        }
      }
    }

    // 3. Clear destination rectangle to replace existing destination cells
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const destRow = targetStartRow + r;
        const destCol = targetStartCol + c;
        if (destRow < this.maxRows && destCol < this.maxCols) {
          const destAddr = formatAddress({ row: destRow, col: destCol });
          curCells.delete(destAddr);
          curFormats.delete(destAddr);
          curComments.delete(destAddr);
        }
      }
    }

    // 4. Write source items to destination
    for (const item of items) {
      const destRow = targetStartRow + item.rOffset;
      const destCol = targetStartCol + item.cOffset;
      if (destRow >= this.maxRows || destCol >= this.maxCols) continue;

      const destAddr = formatAddress({ row: destRow, col: destCol });

      if (item.cell) {
        let rawToSet: string | number = item.cell.raw;
        let formulaToSet: string | undefined = item.cell.formula;

        if (item.cell.formula) {
          if (copy) {
            formulaToSet = this.formulaEngine.shiftFormula(item.cell.formula, rowDelta, colDelta);
            rawToSet = formulaToSet;
          } else {
            formulaToSet = item.cell.formula;
            rawToSet = item.cell.raw;
          }
        }

        if (formulaToSet) {
          curCells.set(destAddr, {
            raw: rawToSet,
            formula: formulaToSet,
            value: '',
            displayValue: '',
            format: item.format,
          });
        } else {
          curCells.set(destAddr, {
            raw: rawToSet,
            value: item.cell.value,
            displayValue: item.cell.displayValue,
            format: item.format,
          });
        }
      }

      if (item.format) {
        curFormats.set(destAddr, {
          ...item.format,
          border: item.format.border ? { ...item.format.border } : undefined,
        });
      }

      if (item.comment) {
        curComments.set(destAddr, item.comment);
      }
    }

    this.cells.set(curCells);
    this.cellFormats.set(curFormats);
    this.comments.set(curComments);
    this.recalculateAll();
  }
}
