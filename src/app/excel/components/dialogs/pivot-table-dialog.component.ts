import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { PivotTableItem } from '../../models/excel.types';
import { ExcelStateService } from '../../services/excel-state.service';
import { formatAddress, indexToColLetter, parseRange } from '../../services/formula-engine.service';

@Component({
  selector: 'app-pivot-table-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './pivot-table-dialog.component.html',
  styleUrls: ['./pivot-table-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PivotTableDialogComponent implements OnInit {
  @Input() initialRange: string = 'A1:C10';
  @Output() readonly addPivotTable = new EventEmitter<PivotTableItem>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly range = signal<string>('A1:C10');
  readonly availableFields = signal<string[]>([]);
  readonly selectedRowField = signal<string>('');
  readonly selectedValueField = signal<string>('');
  readonly selectedSummarize = signal<'SUM' | 'COUNT' | 'AVERAGE'>('SUM');

  constructor(private state: ExcelStateService) {}

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
    this.extractHeaders();
  }

  onRangeChange(newRange: string): void {
    this.range.set(newRange);
    this.extractHeaders();
  }

  private extractHeaders(): void {
    const parsed = parseRange(this.range());
    if (!parsed) return;

    const headers: string[] = [];
    const curCells = this.state.cells();

    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      const addr = formatAddress({ row: parsed.startRow, col: c });
      const val = curCells.get(addr)?.value;
      const label = val !== undefined && val !== null && val !== '' ? String(val).trim() : `Column ${indexToColLetter(c)}`;
      headers.push(label);
    }

    this.availableFields.set(headers);
    if (headers.length > 0) {
      this.selectedRowField.set(headers[0]);
      this.selectedValueField.set(headers.length > 1 ? headers[1] : headers[0]);
    }
  }

  onApply(): void {
    const parsed = parseRange(this.range());
    if (!parsed) return;

    const rowField = this.selectedRowField();
    const valField = this.selectedValueField();
    const summarize = this.selectedSummarize();

    const headers = this.availableFields();
    const rowColIdx = parsed.startCol + headers.indexOf(rowField);
    const valColIdx = parsed.startCol + headers.indexOf(valField);

    // Aggregate data
    const curCells = this.state.cells();
    const grouped = new Map<string, number[]>();

    for (let r = parsed.startRow + 1; r <= parsed.endRow; r++) {
      const rowAddr = formatAddress({ row: r, col: rowColIdx });
      const valAddr = formatAddress({ row: r, col: valColIdx });

      const rowVal = String(curCells.get(rowAddr)?.value ?? 'Blank').trim();
      const val = Number(curCells.get(valAddr)?.value);

      if (!grouped.has(rowVal)) {
        grouped.set(rowVal, []);
      }
      if (!isNaN(val)) {
        grouped.get(rowVal)!.push(val);
      }
    }

    const rows: Array<{ rowLabel: string; value: number }> = [];
    let grandTotal = 0;
    const allValues: number[] = [];

    for (const [key, vals] of grouped.entries()) {
      let calc = 0;
      if (summarize === 'SUM') {
        calc = vals.reduce((a, b) => a + b, 0);
      } else if (summarize === 'COUNT') {
        calc = vals.length;
      } else if (summarize === 'AVERAGE') {
        calc = vals.length > 0 ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
      }
      calc = Number.isInteger(calc) ? calc : Math.round(calc * 100) / 100;
      rows.push({ rowLabel: key, value: calc });
      if (summarize !== 'AVERAGE') {
        grandTotal += calc;
      }
      allValues.push(...vals);
    }

    if (summarize === 'AVERAGE') {
      grandTotal = allValues.length > 0 ? allValues.reduce((a, b) => a + b, 0) / allValues.length : 0;
    }
    grandTotal = Number.isInteger(grandTotal) ? grandTotal : Math.round(grandTotal * 100) / 100;

    const summarizeLabel = summarize === 'COUNT' ? 'Count' : summarize === 'AVERAGE' ? 'Average' : 'Sum';

    const ptItem: PivotTableItem = {
      id: 'pt_' + Date.now(),
      range: this.range().toUpperCase(),
      rowsField: rowField,
      valuesField: valField,
      summarize,
      targetCell: formatAddress({ row: parsed.startRow, col: parsed.endCol + 2 }),
      headers: {
        rowLabel: 'Row Labels',
        valueLabel: `${summarizeLabel} of ${valField}`,
      },
      rows,
      grandTotal,
    };

    this.addPivotTable.emit(ptItem);
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
