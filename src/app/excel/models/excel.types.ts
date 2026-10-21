export interface CellCoord {
  row: number; // 0-indexed
  col: number; // 0-indexed
}

export interface CellRangeCoord {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export type NumberFormatType = 'automatic' | 'general' | 'number' | 'currency' | 'percent' | 'date';
export type TextAlign = 'left' | 'center' | 'right';

export interface CellBorderEdges {
  top?: 'thin' | 'thick';
  bottom?: 'thin' | 'thick';
  left?: 'thin' | 'thick';
  right?: 'thin' | 'thick';
}

export type BorderAction =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'all'
  | 'outside'
  | 'thickOutside'
  | 'noBorder';

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  numberFormat?: NumberFormatType;
  decimals?: number;
  align?: TextAlign;
  textColor?: string;
  fillColor?: string;
  border?: CellBorderEdges;
}

export interface CellData {
  raw: string | number;
  formula?: string;
  value: any;
  displayValue: string;
  format?: CellFormat;
  error?: string; // '#REF!' | '#DIV/0!' | '#NAME?' | '#VALUE!' | '#N/A'
  comment?: string;
}

export type ChartType = 'column' | 'bar' | 'line' | 'pie';

export interface ChartItem {
  id: string;
  chartType: ChartType;
  range: string;
  title: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PivotTableItem {
  id: string;
  range: string;
  rowsField: string;
  valuesField: string;
  summarize: 'SUM' | 'COUNT' | 'AVERAGE';
  targetCell: string;
  headers: { rowLabel: string; valueLabel: string };
  rows: Array<{ rowLabel: string; value: number }>;
  grandTotal: number;
}

export type FilterConditionType =
  | 'equals'
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'contains'
  | 'isEmpty';

export interface FilterCondition {
  type: FilterConditionType;
  value?: any;
  value2?: any;
}

export interface ColumnFilter {
  colIndex: number;
  condition?: FilterCondition;
  selectedValues?: string[]; // values that are checked / visible
}

export interface SheetFilter {
  range: string;
  headerRow: number;
  startCol: number;
  endCol: number;
  startRow: number;
  endRow: number;
  columnFilters: Record<number, ColumnFilter>;
}

export interface ConditionalFormatRule {
  range: string;
  condition: FilterConditionType;
  value?: any;
  value2?: any;
  fillColor: string;
}

export interface DataValidationRule {
  range: string;
  type: 'list';
  allowed: string[];
  rejectInvalid?: boolean;
}

export type SortOrder = 'asc' | 'desc';

export type ExcelActionType =
  | 'undo'
  | 'redo'
  | 'copy'
  | 'cut'
  | 'paste'
  | 'selectRange'
  | 'goToCell';

export interface LoggedAction {
  action: ExcelActionType;
  range?: string;
  timestamp: number;
}

export interface SheetSnapshot {
  cells: Record<string, string | number>;
  formats: Record<string, CellFormat>;
  columnWidths: Record<string, number>;
  charts: ChartItem[];
  comments: Record<string, string>;
  pivotTables: PivotTableItem[];
  filter: SheetFilter | null;
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
}

export interface ExcelLessonTaskState {
  taskIndex: number;
  isPassed: boolean;
  isCompleted: boolean;
  feedbackMessage: string;
  isAssessment: boolean;
  totalTasks: number;
}
