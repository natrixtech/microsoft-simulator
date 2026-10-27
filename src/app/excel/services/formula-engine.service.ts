import { Injectable } from '@angular/core';
import { CellCoord, CellRangeCoord } from '../models/excel.types';

export interface Token {
  type:
    | 'NUMBER'
    | 'STRING'
    | 'BOOLEAN'
    | 'IDENT'
    | 'CELL'
    | 'RANGE'
    | 'OP'
    | 'COMMA'
    | 'LPAREN'
    | 'RPAREN'
    | 'PERCENT'
    | 'ERROR';
  value: string;
}

export type CellValueGetter = (cellAddr: string) => any;
export type RangeValuesGetter = (rangeAddr: string) => any[][];

export function colLetterToIndex(colStr: string): number {
  const clean = colStr.replace(/\$/g, '').toUpperCase();
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function indexToColLetter(index: number): string {
  let temp = index + 1;
  let letter = '';
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

export function parseAddress(addr: string): { row: number; col: number; rowAbs: boolean; colAbs: boolean } | null {
  const m = String(addr).trim().match(/^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/);
  if (!m) return null;
  const colAbs = m[1] === '$';
  const col = colLetterToIndex(m[2]);
  const rowAbs = m[3] === '$';
  const row = parseInt(m[4], 10) - 1;
  return { row, col, rowAbs, colAbs };
}

export function formatAddress(coord: CellCoord, absCol = false, absRow = false): string {
  const c = (absCol ? '$' : '') + indexToColLetter(coord.col);
  const r = (absRow ? '$' : '') + (coord.row + 1);
  return c + r;
}

export function parseRange(rangeStr: string): CellRangeCoord | null {
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    const p = parseAddress(parts[0]);
    if (!p) return null;
    return { startRow: p.row, startCol: p.col, endRow: p.row, endCol: p.col };
  }
  if (parts.length === 2) {
    const p1 = parseAddress(parts[0]);
    const p2 = parseAddress(parts[1]);
    if (!p1 || !p2) return null;
    return {
      startRow: Math.min(p1.row, p2.row),
      startCol: Math.min(p1.col, p2.col),
      endRow: Math.max(p1.row, p2.row),
      endCol: Math.max(p1.col, p2.col),
    };
  }
  return null;
}

export interface DateParts {
  year: number;
  month: number;
  day: number;
  utcTime: number;
}

export function parseDateParts(val: any): DateParts | null {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return null;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = val.getMonth() + 1;
    const d = val.getDate();
    return { year: y, month: m, day: d, utcTime: Date.UTC(y, m - 1, d) };
  }
  if (typeof val === 'string') {
    const s = val.trim();
    // YYYY-MM-DD
    const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (isoMatch) {
      const y = parseInt(isoMatch[1], 10);
      const m = parseInt(isoMatch[2], 10);
      const d = parseInt(isoMatch[3], 10);
      if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1000 && y <= 9999) {
        const dt = new Date(Date.UTC(y, m - 1, d));
        if (dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d) {
          return { year: y, month: m, day: d, utcTime: dt.getTime() };
        }
      }
      return null;
    }
    // M/D/YYYY
    const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
      const m = parseInt(slashMatch[1], 10);
      const d = parseInt(slashMatch[2], 10);
      const y = parseInt(slashMatch[3], 10);
      if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1000 && y <= 9999) {
        const dt = new Date(Date.UTC(y, m - 1, d));
        if (dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d) {
          return { year: y, month: m, day: d, utcTime: dt.getTime() };
        }
      }
      return null;
    }
  }
  return null;
}

export function formatIsoDate(y: number, m: number, d: number): string {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export interface ParsedInput {
  raw: string | number;
  value: any;
  formula?: string;
  isFormula: boolean;
}

@Injectable({
  providedIn: 'root',
})
export class FormulaEngineService {
  private evaluatingCells = new Set<string>();

  /**
   * Evaluates an expression or formula starting with '=' or plain value.
   */
  evaluate(
    expr: string,
    currentCell: string,
    getCellValue: CellValueGetter,
    getRangeValues: RangeValuesGetter
  ): { value: any; error?: string } {
    if (!expr || typeof expr !== 'string') {
      return { value: expr };
    }

    const trimmed = expr.trim();
    if (!trimmed.startsWith('=')) {
      // Check if it looks like a number
      if (!isNaN(Number(trimmed)) && trimmed !== '') {
        return { value: Number(trimmed) };
      }
      return { value: expr };
    }

    const normCell = currentCell.toUpperCase().replace(/\$/g, '');
    if (this.evaluatingCells.has(normCell)) {
      return { value: '#REF!', error: '#REF!' };
    }

    this.evaluatingCells.add(normCell);

    try {
      const formulaText = trimmed.substring(1).trim();
      const tokens = this.tokenize(formulaText);
      const parser = new FormulaParser(tokens, (cell) => {
        const parsed = parseAddress(cell);
        if (!parsed) return '#REF!';
        const cleanAddr = indexToColLetter(parsed.col) + (parsed.row + 1);
        if (this.evaluatingCells.has(cleanAddr)) {
          throw new Error('#REF!');
        }
        return getCellValue(cleanAddr);
      }, getRangeValues);

      let result = parser.parse();
      if (typeof result === 'number' && !isFinite(result)) {
        return { value: '#DIV/0!', error: '#DIV/0!' };
      }
      if (typeof result === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(result)) {
        return { value: result, error: result };
      }
      const nonParenTokens = tokens.filter((t) => t.type !== 'LPAREN' && t.type !== 'RPAREN');
      if (result === '' && nonParenTokens.length === 1 && nonParenTokens[0].type === 'CELL') {
        result = 0;
      }
      return { value: result };
    } catch (err: any) {
      const msg = err?.message;
      if (['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(msg)) {
        return { value: msg, error: msg };
      }
      return { value: '#VALUE!', error: '#VALUE!' };
    } finally {
      this.evaluatingCells.delete(normCell);
    }
  }

  /**
   * Tokenizes a formula string into structured tokens.
   */
  tokenize(formula: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    const len = formula.length;

    while (i < len) {
      const ch = formula[i];

      // Whitespace
      if (/\s/.test(ch)) {
        i++;
        continue;
      }

      // String literal
      if (ch === '"') {
        let str = '';
        i++;
        while (i < len && formula[i] !== '"') {
          if (formula[i] === '\\' && i + 1 < len) {
            i++;
          }
          str += formula[i];
          i++;
        }
        if (i < len && formula[i] === '"') {
          i++;
        }
        tokens.push({ type: 'STRING', value: str });
        continue;
      }

      // Numbers
      if (/[0-9]/.test(ch) || (ch === '.' && i + 1 < len && /[0-9]/.test(formula[i + 1]))) {
        let numStr = '';
        while (i < len && /[0-9.]/.test(formula[i])) {
          numStr += formula[i];
          i++;
        }
        tokens.push({ type: 'NUMBER', value: numStr });
        continue;
      }

      // Operators and punctuation
      if (ch === '(') {
        tokens.push({ type: 'LPAREN', value: '(' });
        i++;
        continue;
      }
      if (ch === ')') {
        tokens.push({ type: 'RPAREN', value: ')' });
        i++;
        continue;
      }
      if (ch === ',') {
        tokens.push({ type: 'COMMA', value: ',' });
        i++;
        continue;
      }
      if (ch === '%') {
        tokens.push({ type: 'PERCENT', value: '%' });
        i++;
        continue;
      }

      // Multi-char comparisons: <=, >=, <>
      if (i + 1 < len) {
        const twoChar = formula.substring(i, i + 2);
        if (['<=', '>=', '<>'].includes(twoChar)) {
          tokens.push({ type: 'OP', value: twoChar });
          i += 2;
          continue;
        }
      }

      // Single-char operators
      if (['+', '-', '*', '/', '^', '&', '=', '<', '>'].includes(ch)) {
        tokens.push({ type: 'OP', value: ch });
        i++;
        continue;
      }

      // Error literal starting with # (#REF!, #DIV/0!, #NAME?, #VALUE!, #N/A, etc.)
      if (ch === '#') {
        let ident = '';
        while (i < len && /[A-Za-z0-9_$#/?!]/.test(formula[i])) {
          ident += formula[i];
          i++;
        }
        tokens.push({ type: 'ERROR', value: ident.toUpperCase() });
        continue;
      }

      // Identifiers, Cell references, Ranges ($A$1:$B$5 or A1 or SUM)
      if (/[A-Za-z_$]/.test(ch)) {
        let ident = '';
        while (i < len && /[A-Za-z0-9_$:]/.test(formula[i])) {
          ident += formula[i];
          i++;
        }

        // Check if range or cell
        if (ident.includes(':')) {
          tokens.push({ type: 'RANGE', value: ident.toUpperCase() });
        } else if (/^\$?[A-Za-z]+\$?[0-9]+$/.test(ident)) {
          tokens.push({ type: 'CELL', value: ident.toUpperCase() });
        } else if (ident.toUpperCase() === 'TRUE' || ident.toUpperCase() === 'FALSE') {
          tokens.push({ type: 'BOOLEAN', value: ident.toUpperCase() });
        } else {
          tokens.push({ type: 'IDENT', value: ident.toUpperCase() });
        }
        continue;
      }

      // Skip unexpected char
      i++;
    }

    return tokens;
  }

  /**
   * Shifts cell/range references in formula by given row/col offsets for copy-paste.
   * Absolute references ($A$1) are preserved.
   */
  shiftFormula(formula: string, rowOffset: number, colOffset: number): string {
    if (!formula.startsWith('=')) return formula;

    // Regex to match cell references and ranges
    // e.g. $A$1, A1:B5, $C2:$D$4
    const cellRegex = /(\$?)([A-Za-z]+)(\$?)([1-9]\d*)/g;

    return formula.replace(cellRegex, (match, colAbs, colStr, rowAbs, rowStr) => {
      let col = colLetterToIndex(colStr);
      let row = parseInt(rowStr, 10) - 1;

      if (!colAbs) {
        col += colOffset;
      }
      if (!rowAbs) {
        row += rowOffset;
      }

      if (col < 0 || row < 0) {
        return '#REF!';
      }

      return (colAbs || '') + indexToColLetter(col) + (rowAbs || '') + (row + 1);
    });
  }

  /**
   * Updates cell/range references in formula when rows or columns are inserted or deleted structurally.
   * Unlike copy-relative shifting, absolute references ($) are also updated structurally to point to
   * the shifted row/col, and deleted rows/cols become #REF!.
   */
  updateFormulaOnRowColChange(
    formula: string,
    type: 'insertRow' | 'deleteRow' | 'insertCol' | 'deleteCol',
    index: number
  ): string {
    if (!formula.startsWith('=')) return formula;

    // 1. Process ranges first: e.g. A1:B5, $A$1:$B$5
    const rangeRegex = /(?<![A-Za-z0-9_])((\$?)([A-Za-z]+)(\$?)([1-9]\d*))\s*:\s*((\$?)([A-Za-z]+)(\$?)([1-9]\d*))(?![A-Za-z0-9_])/g;
    let res = formula.replace(
      rangeRegex,
      (match, p1, colAbs1, colStr1, rowAbs1, rowStr1, p2, colAbs2, colStr2, rowAbs2, rowStr2) => {
        const c1 = colLetterToIndex(colStr1);
        const r1 = parseInt(rowStr1, 10) - 1;
        const c2 = colLetterToIndex(colStr2);
        const r2 = parseInt(rowStr2, 10) - 1;

        const startRow = Math.min(r1, r2);
        const endRow = Math.max(r1, r2);
        const startCol = Math.min(c1, c2);
        const endCol = Math.max(c1, c2);

        let newR1 = r1;
        let newR2 = r2;
        let newC1 = c1;
        let newC2 = c2;

        if (type === 'insertRow') {
          if (index <= startRow) {
            newR1++;
            newR2++;
          } else if (index <= endRow + 1) {
            if (r1 >= r2) newR1++;
            else newR2++;
          }
        } else if (type === 'deleteRow') {
          if (index < startRow) {
            newR1--;
            newR2--;
          } else if (index >= startRow && index <= endRow) {
            if (startRow === endRow) {
              return '#REF!';
            }
            if (r1 >= r2) newR1--;
            else newR2--;
          }
        } else if (type === 'insertCol') {
          if (index <= startCol) {
            newC1++;
            newC2++;
          } else if (index <= endCol + 1) {
            if (c1 >= c2) newC1++;
            else newC2++;
          }
        } else if (type === 'deleteCol') {
          if (index < startCol) {
            newC1--;
            newC2--;
          } else if (index >= startCol && index <= endCol) {
            if (startCol === endCol) {
              return '#REF!';
            }
            if (c1 >= c2) newC1--;
            else newC2--;
          }
        }

        const part1 = (colAbs1 || '') + indexToColLetter(newC1) + (rowAbs1 || '') + (newR1 + 1);
        const part2 = (colAbs2 || '') + indexToColLetter(newC2) + (rowAbs2 || '') + (newR2 + 1);
        return `${part1}:${part2}`;
      }
    );

    // 2. Process single cell references (not part of range)
    const singleCellRegex = /(?<![A-Za-z0-9_:])(\$?)([A-Za-z]+)(\$?)([1-9]\d*)(?![A-Za-z0-9_:(]|(?:\s*:))/g;
    res = res.replace(singleCellRegex, (match, colAbs, colStr, rowAbs, rowStr) => {
      let col = colLetterToIndex(colStr);
      let row = parseInt(rowStr, 10) - 1;

      if (type === 'insertRow') {
        if (row >= index) row++;
      } else if (type === 'deleteRow') {
        if (row === index) return '#REF!';
        if (row > index) row--;
      } else if (type === 'insertCol') {
        if (col >= index) col++;
      } else if (type === 'deleteCol') {
        if (col === index) return '#REF!';
        if (col > index) col--;
      }

      return (colAbs || '') + indexToColLetter(col) + (rowAbs || '') + (row + 1);
    });

    return res;
  }

  updateRangeOnRowColChange(
    rangeStr: string,
    type: 'insertRow' | 'deleteRow' | 'insertCol' | 'deleteCol',
    index: number
  ): string {
    const dummyFormula = `=${rangeStr}`;
    const updated = this.updateFormulaOnRowColChange(dummyFormula, type, index);
    return updated.startsWith('=') ? updated.substring(1) : updated;
  }

  /**
   * Parses typed or pasted input into structured value and raw representation.
   */
  parseInput(rawVal: string | number): ParsedInput {
    if (rawVal === undefined || rawVal === null || rawVal === '') {
      return { raw: '', value: '', isFormula: false };
    }

    if (typeof rawVal === 'number') {
      return { raw: rawVal, value: rawVal, isFormula: false };
    }

    const str = String(rawVal);
    const trimmed = str.trim();

    // 1. Formula
    if (trimmed.startsWith('=')) {
      return {
        raw: str,
        value: '',
        formula: trimmed,
        isFormula: true,
      };
    }

    // 2. Error cells (#REF!, #DIV/0!, #NAME?, #VALUE!, #N/A)
    if (['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(trimmed)) {
      return { raw: str, value: trimmed, isFormula: false };
    }

    // 3. Numbers: 42, 3.5, -7, 1.50 and .5 become numbers. Anything else stays as it is today.
    const numRegex = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;
    if (numRegex.test(trimmed)) {
      const num = Number(trimmed);
      if (!isNaN(num)) {
        return { raw: str, value: num, isFormula: false };
      }
    }

    // 4. Dates: M/D/YYYY and YYYY-MM-DD only when real calendar date; stored as ISO YYYY-MM-DD
    const dateParts = parseDateParts(trimmed);
    if (dateParts) {
      const iso = formatIsoDate(dateParts.year, dateParts.month, dateParts.day);
      return { raw: str, value: iso, isFormula: false };
    }

    // 5. Default text
    return { raw: str, value: str, isFormula: false };
  }

  /**
   * Formats a raw value according to numberFormat and optional decimals.
   */
  formatDisplayValue(val: any, format?: string, decimals?: number): string {
    if (val === undefined || val === null || val === '') {
      return '';
    }

    if (typeof val === 'boolean') {
      return val ? 'TRUE' : 'FALSE';
    }

    if (typeof val === 'string' && (val.toLowerCase() === 'true' || val.toLowerCase() === 'false')) {
      return val.toUpperCase();
    }

    if (typeof val === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(val)) {
      return val;
    }

    const fmt = format ? format.toLowerCase() : 'automatic';

    if (fmt === 'currency') {
      const num = Number(val);
      if (isNaN(num)) return String(val);
      const d = decimals !== undefined ? Math.max(0, decimals) : 2;
      return (
        '$' +
        num.toLocaleString('en-US', {
          minimumFractionDigits: d,
          maximumFractionDigits: d,
          useGrouping: true,
        })
      );
    }

    if (fmt === 'number') {
      const num = Number(val);
      if (isNaN(num)) return String(val);
      const d = decimals !== undefined ? Math.max(0, decimals) : 2;
      return num.toLocaleString('en-US', {
        minimumFractionDigits: d,
        maximumFractionDigits: d,
        useGrouping: true,
      });
    }

    if (fmt === 'percent') {
      const num = Number(val);
      if (isNaN(num)) return String(val);
      const pct = num * 100;
      const d = decimals !== undefined ? Math.max(0, decimals) : (pct % 1 === 0 ? 0 : 1);
      return (
        pct.toLocaleString('en-US', {
          minimumFractionDigits: d,
          maximumFractionDigits: d,
        }) + '%'
      );
    }

    if (fmt === 'date') {
      if (typeof val === 'number' && !isNaN(val) && val > 0) {
        const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(val) * 86400000);
        return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
      }
      if (typeof val === 'string' && /^\d+(\.\d+)?$/.test(val.trim())) {
        const n = Number(val.trim());
        if (n > 0) {
          const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000);
          return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
        }
      }
      const parts = parseDateParts(val);
      if (parts) {
        return `${parts.month}/${parts.day}/${parts.year}`;
      }
      return String(val);
    }

    // Default / automatic / general: an ISO date value shows as m/d/yyyy
    const dateParts = parseDateParts(val);
    if (dateParts && (fmt === 'automatic' || fmt === 'general' || !format)) {
      return `${dateParts.month}/${dateParts.day}/${dateParts.year}`;
    }

    if (typeof val === 'number') {
      if (decimals !== undefined) {
        const d = Math.max(0, decimals);
        return val.toLocaleString('en-US', {
          minimumFractionDigits: d,
          maximumFractionDigits: d,
          useGrouping: false,
        });
      }
      // Display cleanly without floating point precision noise
      return Number(val.toFixed(8)).toString();
    }

    if (typeof val === 'string' && decimals !== undefined && !isNaN(Number(val)) && val.trim() !== '') {
      const num = Number(val);
      const d = Math.max(0, decimals);
      return num.toLocaleString('en-US', {
        minimumFractionDigits: d,
        maximumFractionDigits: d,
        useGrouping: false,
      });
    }

    return String(val);
  }

  getDecimals(val: any, fmt?: { decimals?: number; numberFormat?: string }): number {
    if (fmt?.decimals !== undefined) return fmt.decimals;
    if (fmt?.numberFormat === 'currency' || fmt?.numberFormat === 'number') return 2;
    if (fmt?.numberFormat === 'percent') return 0;
    if (typeof val === 'number') {
      const s = val.toString();
      const dot = s.indexOf('.');
      return dot === -1 ? 0 : s.length - dot - 1;
    }
    if (typeof val === 'string' && !isNaN(Number(val)) && val.trim() !== '') {
      const dot = val.indexOf('.');
      return dot === -1 ? 0 : val.length - dot - 1;
    }
    return 0;
  }
}

/**
 * Recursive-descent parser and evaluator for Excel expressions.
 */
class FormulaParser {
  private pos = 0;

  constructor(
    private tokens: Token[],
    private getCellValue: (cell: string) => any,
    private getRangeValues: (range: string) => any[][]
  ) {}

  parse(): any {
    if (this.tokens.length === 0) return '';
    const res = this.parseExpression();
    return res;
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private consume(expected?: string): Token {
    const cur = this.tokens[this.pos];
    if (!cur) {
      throw new Error('#VALUE!');
    }
    if (expected && cur.value !== expected && cur.type !== expected) {
      throw new Error('#VALUE!');
    }
    this.pos++;
    return cur;
  }

  private skipToArgEnd(): void {
    let depth = 0;
    while (this.pos < this.tokens.length) {
      const cur = this.tokens[this.pos];
      if (cur.type === 'LPAREN') {
        depth++;
      } else if (cur.type === 'RPAREN') {
        if (depth === 0) {
          break;
        }
        depth--;
      } else if (cur.type === 'COMMA' && depth === 0) {
        break;
      }
      this.pos++;
    }
  }

  // Precedence levels:
  // 1. Comparison: =, <>, <, <=, >, >=
  // 2. Concatenation: &
  // 3. Add / Subtract: +, -
  // 4. Multiply / Divide: *, /
  // 5. Exponent: ^
  // 6. Postfix Percent: %
  // 7. Unary: +, -
  // 8. Primary: Number, String, Boolean, Cell, Function Call, Parentheses

  private parseExpression(): any {
    return this.parseComparison();
  }

  private parseComparison(): any {
    let left = this.parseConcat();
    while (this.peek()?.type === 'OP' && ['=', '<>', '<', '<=', '>', '>='].includes(this.peek()!.value)) {
      const op = this.consume().value;
      const right = this.parseConcat();
      left = this.applyComparison(op, left, right);
    }
    return left;
  }

  private applyComparison(op: string, a: any, b: any): boolean {
    // If comparing dates, compare chronologically by UTC timestamp
    const d1 = parseDateParts(a);
    const d2 = parseDateParts(b);
    if (d1 && d2) {
      const t1 = d1.utcTime;
      const t2 = d2.utcTime;
      if (op === '=') return t1 === t2;
      if (op === '<>') return t1 !== t2;
      if (op === '<') return t1 < t2;
      if (op === '<=') return t1 <= t2;
      if (op === '>') return t1 > t2;
      if (op === '>=') return t1 >= t2;
    }

    // If comparing strings, case-insensitive
    if (typeof a === 'string' && typeof b === 'string') {
      const c1 = a.toLowerCase();
      const c2 = b.toLowerCase();
      if (op === '=') return c1 === c2;
      if (op === '<>') return c1 !== c2;
      if (op === '<') return c1 < c2;
      if (op === '<=') return c1 <= c2;
      if (op === '>') return c1 > c2;
      if (op === '>=') return c1 >= c2;
    }

    const n1 = Number(a);
    const n2 = Number(b);
    if (!isNaN(n1) && !isNaN(n2)) {
      if (op === '=') return Math.abs(n1 - n2) < 0.0000001;
      if (op === '<>') return Math.abs(n1 - n2) >= 0.0000001;
      if (op === '<') return n1 < n2;
      if (op === '<=') return n1 <= n2;
      if (op === '>') return n1 > n2;
      if (op === '>=') return n1 >= n2;
    }

    if (op === '=') return a == b;
    if (op === '<>') return a != b;
    return false;
  }

  private parseConcat(): any {
    let left = this.parseAdditive();
    while (this.peek()?.type === 'OP' && this.peek()?.value === '&') {
      this.consume('&');
      const right = this.parseAdditive();
      left = String(left ?? '') + String(right ?? '');
    }
    return left;
  }

  private parseAdditive(): any {
    let left = this.parseMultiplicative();
    while (this.peek()?.type === 'OP' && (this.peek()?.value === '+' || this.peek()?.value === '-')) {
      const op = this.consume().value;
      const right = this.parseMultiplicative();

      const d1 = parseDateParts(left);
      const d2 = parseDateParts(right);

      // Date arithmetic
      if (d1 && d2) {
        // Date - Date gives whole days in UTC
        if (op === '-') {
          left = Math.round((d1.utcTime - d2.utcTime) / 86400000);
          continue;
        }
        // Date + Date is invalid in Excel
        throw new Error('#VALUE!');
      }

      if (d1 && !d2) {
        const days = Number(right);
        if (isNaN(days)) throw new Error('#VALUE!');
        const newTime = op === '+' ? d1.utcTime + Math.round(days) * 86400000 : d1.utcTime - Math.round(days) * 86400000;
        const resDt = new Date(newTime);
        left = formatIsoDate(resDt.getUTCFullYear(), resDt.getUTCMonth() + 1, resDt.getUTCDate());
        continue;
      }

      if (!d1 && d2) {
        if (op === '+') {
          const days = Number(left);
          if (isNaN(days)) throw new Error('#VALUE!');
          const newTime = d2.utcTime + Math.round(days) * 86400000;
          const resDt = new Date(newTime);
          left = formatIsoDate(resDt.getUTCFullYear(), resDt.getUTCMonth() + 1, resDt.getUTCDate());
          continue;
        }
        // Number - Date is invalid in Excel
        throw new Error('#VALUE!');
      }

      const n1 = Number(left);
      const n2 = Number(right);
      if (isNaN(n1) || isNaN(n2)) throw new Error('#VALUE!');
      left = op === '+' ? n1 + n2 : n1 - n2;
    }
    return left;
  }

  private parseMultiplicative(): any {
    let left = this.parseExponent();
    while (this.peek()?.type === 'OP' && (this.peek()?.value === '*' || this.peek()?.value === '/')) {
      const op = this.consume().value;
      const right = this.parseExponent();
      const n1 = Number(left);
      const n2 = Number(right);
      if (isNaN(n1) || isNaN(n2)) throw new Error('#VALUE!');
      if (op === '/') {
        if (n2 === 0) throw new Error('#DIV/0!');
        left = n1 / n2;
      } else {
        left = n1 * n2;
      }
    }
    return left;
  }

  private parseExponent(): any {
    let left = this.parsePostfix();
    while (this.peek()?.type === 'OP' && this.peek()?.value === '^') {
      this.consume('^');
      const right = this.parsePostfix();
      const n1 = Number(left);
      const n2 = Number(right);
      if (isNaN(n1) || isNaN(n2)) throw new Error('#VALUE!');
      left = Math.pow(n1, n2);
    }
    return left;
  }

  private parsePostfix(): any {
    let val = this.parseUnary();
    while (this.peek()?.type === 'PERCENT') {
      this.consume('PERCENT');
      const n = Number(val);
      if (isNaN(n)) throw new Error('#VALUE!');
      val = n / 100;
    }
    return val;
  }

  private parseUnary(): any {
    if (this.peek()?.type === 'OP' && (this.peek()?.value === '+' || this.peek()?.value === '-')) {
      const op = this.consume().value;
      const operand = this.parseUnary();
      const n = Number(operand);
      if (isNaN(n)) throw new Error('#VALUE!');
      return op === '-' ? -n : n;
    }
    return this.parsePrimary();
  }

  private parsePrimary(): any {
    const cur = this.peek();
    if (!cur) throw new Error('#VALUE!');

    // Parentheses
    if (cur.type === 'LPAREN') {
      this.consume('(');
      const val = this.parseExpression();
      this.consume(')');
      return val;
    }

    // Number
    if (cur.type === 'NUMBER') {
      this.consume();
      return Number(cur.value);
    }

    // String
    if (cur.type === 'STRING') {
      this.consume();
      return cur.value;
    }

    // Boolean
    if (cur.type === 'BOOLEAN') {
      this.consume();
      return cur.value === 'TRUE';
    }

    // Error Literal (e.g. #REF!, #DIV/0!, #NAME?, #VALUE!, #N/A)
    if (cur.type === 'ERROR') {
      this.consume();
      throw new Error(cur.value);
    }

    // Cell Reference (e.g. A1, $B$5)
    if (cur.type === 'CELL') {
      this.consume();
      const val = this.getCellValue(cur.value);
      if (typeof val === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(val)) {
        throw new Error(val);
      }
      if (typeof val === 'number') return val;
      if (val === undefined || val === null || val === '') return '';
      if (!isNaN(Number(val))) return Number(val);
      return val;
    }

    // Range reference outside function call (e.g. A1:B5 alone evaluates to top-left or array)
    if (cur.type === 'RANGE') {
      this.consume();
      const vals = this.getRangeValues(cur.value);
      const topVal = vals[0]?.[0] ?? 0;
      if (typeof topVal === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(topVal)) {
        throw new Error(topVal);
      }
      return topVal;
    }

    // Function Call (e.g. SUM, AVERAGE, IF, etc.)
    if (cur.type === 'IDENT') {
      const fnName = this.consume().value;
      if (this.peek()?.type !== 'LPAREN') {
        throw new Error('#NAME?');
      }
      this.consume('(');

      if (fnName === 'IF') {
        const cond = this.parseExpression();
        let val1: any;
        let err1: any = null;
        if (this.peek()?.type === 'COMMA') {
          this.consume(',');
          try {
            val1 = this.parseExpression();
          } catch (e) {
            err1 = e;
            this.skipToArgEnd();
          }
        } else {
          throw new Error('#VALUE!');
        }

        let val2: any = false;
        let err2: any = null;
        let hasBranch2 = false;
        if (this.peek()?.type === 'COMMA') {
          this.consume(',');
          hasBranch2 = true;
          try {
            val2 = this.parseExpression();
          } catch (e) {
            err2 = e;
            this.skipToArgEnd();
          }
        }
        this.consume(')');

        const condition = typeof cond === 'string' && cond.toUpperCase() === 'FALSE' ? false : !!cond;
        if (condition) {
          if (err1) throw err1;
          return val1;
        } else {
          if (hasBranch2) {
            if (err2) throw err2;
            return val2;
          }
          return false;
        }
      }

      const args: any[] = [];
      if (this.peek()?.type !== 'RPAREN') {
        while (true) {
          // Check if argument is a RANGE token
          if (this.peek()?.type === 'RANGE') {
            const rangeToken = this.consume().value;
            args.push({ isRange: true, range: rangeToken, values: this.getRangeValues(rangeToken) });
          } else {
            args.push(this.parseExpression());
          }

          if (this.peek()?.type === 'COMMA') {
            this.consume(',');
          } else {
            break;
          }
        }
      }
      this.consume(')');
      return this.executeFunction(fnName, args);
    }

    throw new Error('#VALUE!');
  }

  private flattenNumbers(args: any[]): number[] {
    const nums: number[] = [];
    for (const arg of args) {
      if (arg && arg.isRange) {
        for (const row of arg.values as any[][]) {
          for (const cell of row) {
            if (typeof cell === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(cell)) {
              throw new Error(cell);
            }
            if (cell === null || cell === undefined || cell === '') {
              continue;
            }
            if (typeof cell === 'number') {
              nums.push(cell);
            } else if (typeof cell !== 'boolean' && !isNaN(Number(cell))) {
              nums.push(Number(cell));
            }
          }
        }
      } else if (Array.isArray(arg)) {
        nums.push(...this.flattenNumbers(arg));
      } else if (typeof arg === 'string' && ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(arg)) {
        throw new Error(arg);
      } else if (arg === null || arg === undefined || arg === '') {
        continue;
      } else if (typeof arg === 'number') {
        nums.push(arg);
      } else if (typeof arg !== 'boolean' && !isNaN(Number(arg))) {
        nums.push(Number(arg));
      }
    }
    return nums;
  }

  private flattenAll(args: any[]): any[] {
    const items: any[] = [];
    for (const arg of args) {
      if (arg && arg.isRange) {
        for (const row of arg.values as any[][]) {
          for (const cell of row) {
            items.push(cell);
          }
        }
      } else if (Array.isArray(arg)) {
        items.push(...this.flattenAll(arg));
      } else {
        items.push(arg);
      }
    }
    return items;
  }

  private executeFunction(name: string, args: any[]): any {
    switch (name) {
      // 1. Math and Stats
      case 'SUM': {
        const nums = this.flattenNumbers(args);
        return nums.reduce((acc, v) => acc + v, 0);
      }

      case 'AVERAGE': {
        const nums = this.flattenNumbers(args);
        if (nums.length === 0) throw new Error('#DIV/0!');
        return nums.reduce((acc, v) => acc + v, 0) / nums.length;
      }

      case 'COUNT': {
        const nums = this.flattenNumbers(args);
        return nums.length;
      }

      case 'COUNTA': {
        const items = this.flattenAll(args);
        return items.filter((x) => x !== null && x !== undefined && x !== '').length;
      }

      case 'MAX': {
        const nums = this.flattenNumbers(args);
        if (nums.length === 0) return 0;
        return Math.max(...nums);
      }

      case 'MIN': {
        const nums = this.flattenNumbers(args);
        if (nums.length === 0) return 0;
        return Math.min(...nums);
      }

      case 'ROUND': {
        if (args.length < 1) throw new Error('#VALUE!');
        const val = Number(args[0]);
        const digits = args.length > 1 ? Number(args[1]) : 0;
        if (isNaN(val) || isNaN(digits)) throw new Error('#VALUE!');
        const factor = Math.pow(10, digits);
        return Math.round(val * factor) / factor;
      }

      // 2. Conditional
      case 'IF': {
        if (args.length < 2) throw new Error('#VALUE!');
        const condition = !!args[0];
        if (condition) {
          return args[1];
        }
        return args.length > 2 ? args[2] : false;
      }

      case 'SUMIF': {
        if (args.length < 2) throw new Error('#VALUE!');
        const rangeArg = args[0];
        const criteria = args[1];
        const sumRangeArg = args.length > 2 ? args[2] : rangeArg;

        const checkCells = this.extractFlatCells(rangeArg);
        const sumCells = this.extractFlatCells(sumRangeArg);

        let sum = 0;
        for (let i = 0; i < checkCells.length; i++) {
          if (this.matchesCriteria(checkCells[i], criteria)) {
            const raw = i < sumCells.length ? sumCells[i] : null;
            if (raw !== null && raw !== undefined && raw !== '') {
              const valToSum = Number(raw);
              if (!isNaN(valToSum)) {
                sum += valToSum;
              }
            }
          }
        }
        return sum;
      }

      case 'COUNTIF': {
        if (args.length < 2) throw new Error('#VALUE!');
        const rangeArg = args[0];
        const criteria = args[1];
        const checkCells = this.extractFlatCells(rangeArg);

        let count = 0;
        for (const cell of checkCells) {
          if (this.matchesCriteria(cell, criteria)) {
            count++;
          }
        }
        return count;
      }

      case 'AVERAGEIF': {
        if (args.length < 2) throw new Error('#VALUE!');
        const rangeArg = args[0];
        const criteria = args[1];
        const avgRangeArg = args.length > 2 ? args[2] : rangeArg;

        const checkCells = this.extractFlatCells(rangeArg);
        const avgCells = this.extractFlatCells(avgRangeArg);

        let sum = 0;
        let count = 0;
        for (let i = 0; i < checkCells.length; i++) {
          if (this.matchesCriteria(checkCells[i], criteria)) {
            const raw = i < avgCells.length ? avgCells[i] : null;
            if (raw !== null && raw !== undefined && raw !== '') {
              const valToAvg = Number(raw);
              if (!isNaN(valToAvg)) {
                sum += valToAvg;
                count++;
              }
            }
          }
        }
        if (count === 0) throw new Error('#DIV/0!');
        return sum / count;
      }

      // 3. Lookup
      case 'VLOOKUP': {
        if (args.length < 3) throw new Error('#VALUE!');
        const lookupVal = args[0];
        const tableArg = args[1];
        const colIdx = Number(args[2]);
        // Default range_lookup is false (exact match) for our activities
        const exactMatch = args.length > 3 ? !args[3] : true;

        if (!tableArg || !tableArg.isRange) throw new Error('#VALUE!');
        const grid = tableArg.values as any[][];
        if (colIdx < 1 || colIdx > (grid[0]?.length || 0)) {
          throw new Error('#REF!');
        }

        let foundRow: any[] | null = null;
        for (const row of grid) {
          const firstCol = row[0];
          if (this.equalsLookup(firstCol, lookupVal, exactMatch)) {
            foundRow = row;
            break;
          }
        }

        if (!foundRow) {
          throw new Error('#N/A');
        }

        return foundRow[colIdx - 1];
      }

      // 4. Text and Date
      case 'TEXT': {
        if (args.length < 2) throw new Error('#VALUE!');
        const val = args[0];
        const fmt = String(args[1]).toLowerCase();
        if (fmt.includes('$') || fmt.includes('currency')) {
          const n = Number(val);
          return '$' + (isNaN(n) ? '0.00' : n.toFixed(2));
        }
        if (fmt.includes('%')) {
          const n = Number(val);
          return (isNaN(n) ? '0' : (n * 100).toFixed(0)) + '%';
        }
        return String(val);
      }

      case 'CONCATENATE': {
        const items = this.flattenAll(args);
        return items.map((x) => (x !== null && x !== undefined ? String(x) : '')).join('');
      }

      case 'DATE': {
        if (args.length < 3) throw new Error('#VALUE!');
        const y = Number(args[0]);
        const m = Number(args[1]);
        const d = Number(args[2]);
        if (isNaN(y) || isNaN(m) || isNaN(d)) throw new Error('#VALUE!');
        const monthStr = String(m).padStart(2, '0');
        const dayStr = String(d).padStart(2, '0');
        return `${y}-${monthStr}-${dayStr}`;
      }

      case 'TODAY': {
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const d = String(now.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }

      case 'YEAR': {
        if (args.length < 1) throw new Error('#VALUE!');
        const d = parseDateParts(args[0]);
        if (!d) throw new Error('#VALUE!');
        return d.year;
      }

      case 'MONTH': {
        if (args.length < 1) throw new Error('#VALUE!');
        const d = parseDateParts(args[0]);
        if (!d) throw new Error('#VALUE!');
        return d.month;
      }

      case 'DAY': {
        if (args.length < 1) throw new Error('#VALUE!');
        const d = parseDateParts(args[0]);
        if (!d) throw new Error('#VALUE!');
        return d.day;
      }

      default:
        throw new Error('#NAME?');
    }
  }

  private extractFlatCells(arg: any): any[] {
    if (arg && arg.isRange) {
      const flat: any[] = [];
      for (const row of arg.values as any[][]) {
        for (const cell of row) {
          flat.push(cell);
        }
      }
      return flat;
    }
    return [arg];
  }

  private matchesCriteria(cellVal: any, criteria: any): boolean {
    if (criteria === undefined || criteria === null) return false;
    const isBlank = cellVal === null || cellVal === undefined || cellVal === '';
    const critStr = String(criteria).trim();

    // Condition pattern e.g. ">50", "<=10", "<>Pencil", "=Apple"
    const m = critStr.match(/^(<=|>=|<>|<|>|=)(.*)$/);
    if (m) {
      const op = m[1];
      const targetStr = m[2].trim();

      if (isBlank) {
        if (op === '=') return targetStr === '';
        return false;
      }

      const d1 = parseDateParts(cellVal);
      const d2 = parseDateParts(targetStr);
      if (d1 && d2) {
        const t1 = d1.utcTime;
        const t2 = d2.utcTime;
        if (op === '=') return t1 === t2;
        if (op === '<>') return t1 !== t2;
        if (op === '<') return t1 < t2;
        if (op === '<=') return t1 <= t2;
        if (op === '>') return t1 > t2;
        if (op === '>=') return t1 >= t2;
      }

      const numCell = Number(cellVal);
      const numTarget = Number(targetStr);

      if (!isNaN(numCell) && !isNaN(numTarget)) {
        if (op === '=') return Math.abs(numCell - numTarget) < 0.0000001;
        if (op === '<>') return Math.abs(numCell - numTarget) >= 0.0000001;
        if (op === '<') return numCell < numTarget;
        if (op === '<=') return numCell <= numTarget;
        if (op === '>') return numCell > numTarget;
        if (op === '>=') return numCell >= numTarget;
      }

      // String comparison
      const s1 = String(cellVal ?? '').toLowerCase();
      const s2 = targetStr.toLowerCase();
      if (op === '=') return s1 === s2;
      if (op === '<>') return s1 !== s2;
      return false;
    }

    if (isBlank) {
      return critStr === '' || critStr === '=';
    }

    const d1 = parseDateParts(cellVal);
    const d2 = parseDateParts(criteria);
    if (d1 && d2) {
      return d1.utcTime === d2.utcTime;
    }

    // Direct value comparison
    const n1 = Number(cellVal);
    const n2 = Number(criteria);
    if (!isNaN(n1) && !isNaN(n2)) {
      return Math.abs(n1 - n2) < 0.0000001;
    }

    return String(cellVal ?? '').trim().toLowerCase() === critStr.toLowerCase();
  }

  private equalsLookup(cellVal: any, lookupVal: any, exact: boolean): boolean {
    if (cellVal === null || cellVal === undefined) {
      return lookupVal === null || lookupVal === undefined || lookupVal === '';
    }
    const s1 = String(cellVal ?? '').trim().toLowerCase();
    const s2 = String(lookupVal ?? '').trim().toLowerCase();
    if (s1 === s2) return true;

    const n1 = Number(cellVal);
    const n2 = Number(lookupVal);
    if (!isNaN(n1) && !isNaN(n2)) {
      return Math.abs(n1 - n2) < 0.0000001;
    }
    return false;
  }
}
