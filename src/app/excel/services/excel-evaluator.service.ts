import { Injectable } from '@angular/core';
import { TaskGoal } from 'shared';
import { ExcelStateService } from './excel-state.service';
import { ExcelActionLoggerService } from './excel-action-logger.service';
import {
  colLetterToIndex,
  formatAddress,
  parseAddress,
  parseDateParts,
  parseRange,
} from './formula-engine.service';
import { COLORS } from 'shared';

@Injectable({
  providedIn: 'root',
})
export class ExcelEvaluatorService {
  constructor(
    private state: ExcelStateService,
    private actionLogger: ExcelActionLoggerService
  ) {}

  /**
   * Evaluates whether all goals for current task are passed.
   * Returns { passed: boolean, failedGoals: TaskGoal[] }
   */
  evaluateGoals(goals: TaskGoal[]): { passed: boolean; failedGoals: TaskGoal[] } {
    if (!goals || goals.length === 0) {
      return { passed: true, failedGoals: [] };
    }

    const failed: TaskGoal[] = [];

    for (const goal of goals) {
      const ok = this.evaluateSingleGoal(goal);
      if (!ok) {
        failed.push(goal);
      }
    }

    return {
      passed: failed.length === 0,
      failedGoals: failed,
    };
  }

  private evaluateSingleGoal(goal: TaskGoal): boolean {
    const type = goal.type;

    switch (type) {
      case 'cellValue':
        return this.evalCellValue(goal);
      case 'cellFormula':
        return this.evalCellFormula(goal);
      case 'cellFormat':
        return this.evalCellFormat(goal);
      case 'sorted':
        return this.evalSorted(goal);
      case 'filter':
        return this.evalFilter(goal);
      case 'chart':
        return this.evalChart(goal);
      case 'pivotTable':
        return this.evalPivotTable(goal);
      case 'comment':
        return this.evalComment(goal);
      case 'conditionalFormat':
        return this.evalConditionalFormat(goal);
      case 'dataValidation':
        return this.evalDataValidation(goal);
      case 'noDuplicates':
        return this.evalNoDuplicates(goal);
      case 'action':
        return this.evalAction(goal);
      default:
        return false;
    }
  }

  // 1. cellValue
  private evalCellValue(goal: any): boolean {
    const cells = goal.cells as Record<string, any>;
    if (!cells || typeof cells !== 'object') return false;

    for (const [addr, expected] of Object.entries(cells)) {
      const cell = this.state.getCell(addr);
      const actualVal = cell?.value;
      if (!this.valuesEqual(actualVal, expected)) {
        return false;
      }
    }
    return true;
  }

  // 2. cellFormula
  private evalCellFormula(goal: any): boolean {
    const cellAddr = goal.cell;
    if (!cellAddr) return false;

    const cell = this.state.getCell(cellAddr);
    if (!cell || !cell.formula) return false;

    const formulaRaw = cell.formula.trim();
    // Normalize spaces and lowercase for comparison
    const normFormula = formulaRaw.replace(/\s+/g, '').toUpperCase();

    // Check equals if given
    if (goal.equals !== undefined) {
      const expectedNorm = String(goal.equals).replace(/\s+/g, '').toUpperCase();
      if (normFormula !== expectedNorm) return false;
    }

    // Check function if given
    if (goal.function !== undefined) {
      const fnName = String(goal.function).trim().toUpperCase();
      const fnRegex = new RegExp(`[=(+\\-*/,]${fnName}\\(`, 'i');
      if (!fnRegex.test(normFormula) && !normFormula.startsWith(`=${fnName}(`)) {
        return false;
      }
    }

    // Check references if given
    if (goal.references !== undefined) {
      const actualRefs = this.extractFormulaReferences(formulaRaw).sort();
      const expectedRefs = this.normalizeExpectedReferences(goal.references).sort();

      if (actualRefs.length !== expectedRefs.length) {
        return false;
      }
      for (let i = 0; i < actualRefs.length; i++) {
        if (actualRefs[i] !== expectedRefs[i]) {
          return false;
        }
      }
    }

    // Check absolute if given
    if (goal.absolute !== undefined) {
      const hasDollar = formulaRaw.includes('$');
      if (goal.absolute && !hasDollar) return false;
      if (!goal.absolute && hasDollar) return false;
    }

    return true;
  }

  private extractFormulaReferences(formula: string): string[] {
    const cleanFormula = formula.startsWith('=') ? formula.slice(1) : formula;
    const refRegex = /(?<![A-Za-z0-9_])(\$?[A-Za-z]+\$?[1-9]\d*(?:\s*:\s*\$?[A-Za-z]+\$?[1-9]\d*)?)(?!\s*\()/g;
    const matches: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = refRegex.exec(cleanFormula)) !== null) {
      matches.push(m[1].replace(/[\s$]/g, '').toUpperCase());
    }
    return matches;
  }

  private normalizeExpectedReferences(refs: any): string[] {
    if (Array.isArray(refs)) {
      return refs.flatMap((r) => this.normalizeExpectedReferences(r));
    }
    const str = String(refs);
    const refRegex = /(?<![A-Za-z0-9_])(\$?[A-Za-z]+\$?[1-9]\d*(?:\s*:\s*\$?[A-Za-z]+\$?[1-9]\d*)?)(?!\s*\()/g;
    const matches: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = refRegex.exec(str)) !== null) {
      matches.push(m[1].replace(/[\s$]/g, '').toUpperCase());
    }
    if (matches.length === 0 && str.trim()) {
      matches.push(str.replace(/[\s$]/g, '').toUpperCase());
    }
    return matches;
  }

  // 3. cellFormat
  private evalCellFormat(goal: any): boolean {
    const rangeStr = goal.range || goal.cell;
    if (!rangeStr) return false;

    const parsed = parseRange(rangeStr);
    if (!parsed) return false;

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const fmt = this.state.cellFormats().get(addr) || {};

        if (goal.bold !== undefined && !!fmt.bold !== !!goal.bold) return false;
        if (goal.italic !== undefined && !!fmt.italic !== !!goal.italic) return false;
        if (goal.underline !== undefined && !!fmt.underline !== !!goal.underline) return false;
        if (goal.fontSize !== undefined) {
          const actualSize = fmt.fontSize !== undefined ? fmt.fontSize : 10;
          if (actualSize !== Number(goal.fontSize)) return false;
        }
        if (goal.fontFamily !== undefined) {
          const actualFont = fmt.fontFamily || 'Calibri';
          if (actualFont.toLowerCase() !== String(goal.fontFamily).toLowerCase()) return false;
        }

        if (goal.numberFormat !== undefined) {
          const expectedFmt = String(goal.numberFormat).toLowerCase().trim();
          const rawActual = fmt.numberFormat ? String(fmt.numberFormat).toLowerCase().trim() : '';
          const normExpected = (!expectedFmt || expectedFmt === 'general') ? 'automatic' : expectedFmt;
          const normActual = (!rawActual || rawActual === 'general' || rawActual === 'automatic') ? 'automatic' : rawActual;
          if (normActual !== normExpected) return false;
        }

        if (goal.align !== undefined && fmt.align !== goal.align) return false;

        if (goal.textColor !== undefined) {
          if (!this.colorsEqual(fmt.textColor, goal.textColor)) return false;
        }

        if (goal.fillColor !== undefined) {
          if (!this.colorsEqual(fmt.fillColor, goal.fillColor)) return false;
        }
      }
    }

    return true;
  }

  // 4. sorted
  private evalSorted(goal: any): boolean {
    const rangeStr = goal.range;
    const byColumn = goal.byColumn;
    const order = goal.order || 'asc';
    const hasHeader = !!goal.hasHeader;

    if (!rangeStr || !byColumn) return false;

    const parsed = parseRange(rangeStr);
    if (!parsed) return false;

    const colIdx = colLetterToIndex(byColumn);
    const startRow = hasHeader ? parsed.startRow + 1 : parsed.startRow;

    let prevVal: any = null;

    for (let r = startRow; r <= parsed.endRow; r++) {
      const addr = formatAddress({ row: r, col: colIdx });
      const curVal = this.state.getCell(addr)?.value ?? '';

      if (prevVal !== null) {
        const comp = this.compareValues(prevVal, curVal);
        if (order === 'asc' && comp > 0) return false;
        if (order === 'desc' && comp < 0) return false;
      }
      prevVal = curVal;
    }

    return true;
  }

  // 5. filter
  private evalFilter(goal: any): boolean {
    const f = this.state.filter();
    if (!f) return false;

    if (goal.range) {
      const p1 = parseRange(goal.range);
      const p2 = parseRange(f.range);
      if (!p1 || !p2) return false;
      if (
        p1.startRow !== p2.startRow ||
        p1.endRow !== p2.endRow ||
        p1.startCol !== p2.startCol ||
        p1.endCol !== p2.endCol
      ) {
        return false;
      }
    }

    if (goal.column) {
      const colIdx = colLetterToIndex(goal.column);
      const colF = f.columnFilters[colIdx];
      if (!colF) return false;

      if (goal.condition) {
        if (!colF.condition || colF.condition.type !== goal.condition) return false;
        if (goal.value !== undefined && !this.valuesEqual(colF.condition.value, goal.value)) {
          return false;
        }
      }
    }

    return true;
  }

  // 6. chart
  private evalChart(goal: any): boolean {
    const charts = this.state.charts();

    if (goal.count !== undefined) {
      if (charts.length !== Number(goal.count)) return false;
    }

    // Must find at least one chart that matches all other given settings
    const matching = charts.find((ch) => {
      if (goal.chartType && ch.chartType.toLowerCase() !== String(goal.chartType).toLowerCase()) {
        return false;
      }
      if (goal.range) {
        const p1 = parseRange(goal.range);
        const p2 = parseRange(ch.range);
        if (!p1 || !p2) return false;
        if (
          p1.startRow !== p2.startRow ||
          p1.endRow !== p2.endRow ||
          p1.startCol !== p2.startCol ||
          p1.endCol !== p2.endCol
        ) {
          return false;
        }
      }
      if (goal.title !== undefined) {
        const t1 = ch.title.trim().toLowerCase();
        const t2 = String(goal.title).trim().toLowerCase();
        if (t1 !== t2) return false;
      }
      return true;
    });

    return !!matching;
  }

  // 7. pivotTable
  private evalPivotTable(goal: any): boolean {
    const pts = this.state.pivotTables();

    const matching = pts.find((pt) => {
      if (goal.range) {
        const p1 = parseRange(goal.range);
        const p2 = parseRange(pt.range);
        if (!p1 || !p2) return false;
        if (
          p1.startRow !== p2.startRow ||
          p1.endRow !== p2.endRow ||
          p1.startCol !== p2.startCol ||
          p1.endCol !== p2.endCol
        ) {
          return false;
        }
      }
      if (goal.rowsField !== undefined) {
        if (pt.rowsField.trim().toLowerCase() !== String(goal.rowsField).trim().toLowerCase()) {
          return false;
        }
      }
      if (goal.valuesField !== undefined) {
        if (pt.valuesField.trim().toLowerCase() !== String(goal.valuesField).trim().toLowerCase()) {
          return false;
        }
      }
      if (goal.summarize !== undefined) {
        if (pt.summarize.toUpperCase() !== String(goal.summarize).toUpperCase()) {
          return false;
        }
      }
      return true;
    });

    return !!matching;
  }

  // 8. comment
  private evalComment(goal: any): boolean {
    const cellAddr = goal.cell;
    if (!cellAddr) return false;

    const norm = cellAddr.toUpperCase().replace(/\$/g, '');
    const comment = this.state.comments().get(norm);
    if (!comment) return false;

    if (goal.contains !== undefined) {
      const expected = String(goal.contains).trim().toLowerCase();
      if (!comment.toLowerCase().includes(expected)) {
        return false;
      }
    }

    return true;
  }

  // 9. conditionalFormat
  private evalConditionalFormat(goal: any): boolean {
    const rules = this.state.conditionalFormats();

    const matching = rules.find((r) => {
      if (goal.range) {
        const p1 = parseRange(goal.range);
        const p2 = parseRange(r.range);
        if (!p1 || !p2) return false;
        if (
          p1.startRow !== p2.startRow ||
          p1.endRow !== p2.endRow ||
          p1.startCol !== p2.startCol ||
          p1.endCol !== p2.endCol
        ) {
          return false;
        }
      }
      if (goal.condition !== undefined && r.condition !== goal.condition) return false;
      if (goal.value !== undefined && !this.valuesEqual(r.value, goal.value)) return false;
      if (goal.fillColor !== undefined && !this.colorsEqual(r.fillColor, goal.fillColor)) {
        return false;
      }
      return true;
    });

    return !!matching;
  }

  // 10. dataValidation
  private evalDataValidation(goal: any): boolean {
    const rules = this.state.dataValidations();

    const matching = rules.find((dv) => {
      if (goal.range) {
        const p1 = parseRange(goal.range);
        const p2 = parseRange(dv.range);
        if (!p1 || !p2) return false;
        if (
          p1.startRow !== p2.startRow ||
          p1.endRow !== p2.endRow ||
          p1.startCol !== p2.startCol ||
          p1.endCol !== p2.endCol
        ) {
          return false;
        }
      }
      if (goal.allowed !== undefined) {
        const expectedList = (goal.allowed as string[]).map((s) => String(s).trim().toLowerCase());
        const actualList = dv.allowed.map((s) => String(s).trim().toLowerCase());
        if (expectedList.length !== actualList.length) return false;
        for (let i = 0; i < expectedList.length; i++) {
          if (expectedList[i] !== actualList[i]) return false;
        }
      }
      return true;
    });

    return !!matching;
  }

  // 11. noDuplicates
  private evalNoDuplicates(goal: any): boolean {
    const rangeStr = goal.range;
    if (!rangeStr) return false;

    const parsed = parseRange(rangeStr);
    if (!parsed) return false;

    const seenRows = new Set<string>();

    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      const rowParts: string[] = [];
      let isEmpty = true;

      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const val = this.state.getCell(addr)?.value ?? '';
        if (val !== '') isEmpty = false;
        rowParts.push(String(val).trim().toLowerCase());
      }

      if (isEmpty) continue;

      const key = rowParts.join('||');
      if (seenRows.has(key)) {
        return false;
      }
      seenRows.add(key);
    }

    return true;
  }

  // 12. action
  private evalAction(goal: any): boolean {
    const action = goal.action;
    const range = goal.range || goal.cell;
    return this.actionLogger.hasAction(action, range);
  }

  // Helper comparisons
  private valuesEqual(a: any, b: any): boolean {
    if (a === undefined || a === null) a = '';
    if (b === undefined || b === null) b = '';

    const d1 = parseDateParts(a);
    const d2 = parseDateParts(b);
    if (d1 && d2) {
      return d1.utcTime === d2.utcTime;
    }

    const n1 = Number(a);
    const n2 = Number(b);
    if (!isNaN(n1) && !isNaN(n2) && a !== '' && b !== '') {
      return Math.abs(n1 - n2) <= 0.001; // ±0.001 tolerance
    }

    return String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
  }

  private compareValues(a: any, b: any): number {
    const d1 = parseDateParts(a);
    const d2 = parseDateParts(b);
    if (d1 && d2) {
      if (d1.utcTime === d2.utcTime) return 0;
      return d1.utcTime < d2.utcTime ? -1 : 1;
    }

    const n1 = Number(a);
    const n2 = Number(b);
    if (!isNaN(n1) && !isNaN(n2) && a !== '' && b !== '') {
      if (Math.abs(n1 - n2) < 0.0001) return 0;
      return n1 < n2 ? -1 : 1;
    }
    const s1 = String(a ?? '').trim().toLowerCase();
    const s2 = String(b ?? '').trim().toLowerCase();
    if (s1 < s2) return -1;
    if (s1 > s2) return 1;
    return 0;
  }

  private colorsEqual(c1?: string, c2?: string): boolean {
    if (!c1 || !c2) return false;
    const hex1 = (COLORS[c1] || c1).toLowerCase();
    const hex2 = (COLORS[c2] || c2).toLowerCase();
    if (hex1 === hex2) return true;
    return c1.toLowerCase() === c2.toLowerCase();
  }
}
