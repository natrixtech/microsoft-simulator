import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { colLetterToIndex, indexToColLetter, parseRange } from '../../services/formula-engine.service';
import { ExcelStateService } from '../../services/excel-state.service';
import { SortOrder } from '../../models/excel.types';

@Component({
  selector: 'app-sort-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './sort-dialog.component.html',
  styleUrls: ['./sort-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class SortDialogComponent implements OnInit {
  @Input() rangeStr: string = 'A1:C10';
  @Output() readonly applySort = new EventEmitter<{ byColumn: string; order: SortOrder; hasHeader: boolean }>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly hasHeader = signal<boolean>(true);
  readonly selectedColumn = signal<string>('A');
  readonly selectedOrder = signal<SortOrder>('asc');
  readonly availableColumns = signal<Array<{ letter: string; label: string }>>([]);

  constructor(private state: ExcelStateService) {}

  ngOnInit(): void {
    this.updateColumns();
  }

  onHasHeaderChange(checked: boolean): void {
    this.hasHeader.set(checked);
    this.updateColumns();
  }

  private updateColumns(): void {
    const parsed = parseRange(this.rangeStr);
    if (!parsed) return;

    const cols: Array<{ letter: string; label: string }> = [];
    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      const letter = indexToColLetter(c);
      let label = `Column ${letter}`;
      if (this.hasHeader()) {
        const headerAddr = `${letter}${parsed.startRow + 1}`;
        const headerVal = this.state.getCell(headerAddr)?.value;
        if (headerVal !== undefined && headerVal !== null && headerVal !== '') {
          label = `${headerVal} (${letter})`;
        }
      }
      cols.push({ letter, label });
    }
    this.availableColumns.set(cols);
    if (cols.length > 0 && !cols.some((c) => c.letter === this.selectedColumn())) {
      this.selectedColumn.set(cols[0].letter);
    }
  }

  onApply(): void {
    this.applySort.emit({
      byColumn: this.selectedColumn(),
      order: this.selectedOrder(),
      hasHeader: this.hasHeader(),
    });
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
