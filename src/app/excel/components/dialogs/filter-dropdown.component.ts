import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { FilterCondition, FilterConditionType, SortOrder } from '../../models/excel.types';
import { ExcelStateService } from '../../services/excel-state.service';
import { formatAddress, indexToColLetter } from '../../services/formula-engine.service';

@Component({
  selector: 'app-filter-dropdown',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './filter-dropdown.component.html',
  styleUrls: ['./filter-dropdown.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class FilterDropdownComponent implements OnInit {
  @Input() colIndex: number = 0;
  @Input() anchorX: number = 0;
  @Input() anchorY: number = 0;

  @Output() readonly applyFilter = new EventEmitter<{
    colIndex: number;
    condition?: FilterCondition;
    selectedValues?: string[];
  }>();
  @Output() readonly sortColumn = new EventEmitter<{ colIndex: number; order: SortOrder }>();
  @Output() readonly closeDropdown = new EventEmitter<void>();

  readonly allValues = signal<string[]>([]);
  readonly selectedValues = signal<Set<string>>(new Set());
  readonly selectAll = signal<boolean>(true);

  // Condition filter
  readonly showConditionForm = signal<boolean>(false);
  readonly conditionType = signal<FilterConditionType>('greaterThan');
  readonly conditionVal1 = signal<string>('');
  readonly conditionVal2 = signal<string>('');

  constructor(private state: ExcelStateService) {}

  ngOnInit(): void {
    this.extractUniqueValues();
  }

  private extractUniqueValues(): void {
    const f = this.state.filter();
    if (!f) return;

    const values = new Set<string>();
    const curCells = this.state.cells();

    for (let r = f.startRow + 1; r <= f.endRow; r++) {
      const addr = formatAddress({ row: r, col: this.colIndex });
      const cellVal = curCells.get(addr)?.value;
      if (cellVal !== undefined && cellVal !== null && cellVal !== '') {
        values.add(String(cellVal).trim());
      }
    }

    const valList = Array.from(values).sort();
    this.allValues.set(valList);

    // Existing filter state
    const existingColFilter = f.columnFilters[this.colIndex];
    if (existingColFilter?.selectedValues && existingColFilter.selectedValues.length > 0) {
      this.selectedValues.set(new Set(existingColFilter.selectedValues));
      this.selectAll.set(existingColFilter.selectedValues.length === valList.length);
    } else {
      this.selectedValues.set(new Set(valList));
      this.selectAll.set(true);
    }

    if (existingColFilter?.condition) {
      this.conditionType.set(existingColFilter.condition.type);
      this.conditionVal1.set(existingColFilter.condition.value !== undefined ? String(existingColFilter.condition.value) : '');
      this.conditionVal2.set(existingColFilter.condition.value2 !== undefined ? String(existingColFilter.condition.value2) : '');
      this.showConditionForm.set(true);
    }
  }

  onToggleSelectAll(checked: boolean): void {
    this.selectAll.set(checked);
    if (checked) {
      this.selectedValues.set(new Set(this.allValues()));
    } else {
      this.selectedValues.set(new Set());
    }
  }

  onToggleValue(val: string, checked: boolean): void {
    const cur = new Set(this.selectedValues());
    if (checked) {
      cur.add(val);
    } else {
      cur.delete(val);
    }
    this.selectedValues.set(cur);
    this.selectAll.set(cur.size === this.allValues().length);
  }

  onSort(order: SortOrder): void {
    this.sortColumn.emit({ colIndex: this.colIndex, order });
    this.closeDropdown.emit();
  }

  onApply(): void {
    let cond: FilterCondition | undefined;
    if (this.showConditionForm() && (this.conditionVal1() || this.conditionType() === 'isEmpty')) {
      cond = {
        type: this.conditionType(),
        value: this.conditionVal1() ? (!isNaN(Number(this.conditionVal1())) ? Number(this.conditionVal1()) : this.conditionVal1()) : undefined,
        value2: this.conditionVal2() ? (!isNaN(Number(this.conditionVal2())) ? Number(this.conditionVal2()) : this.conditionVal2()) : undefined,
      };
    }

    this.applyFilter.emit({
      colIndex: this.colIndex,
      condition: cond,
      selectedValues: Array.from(this.selectedValues()),
    });
    this.closeDropdown.emit();
  }

  onClearFilter(): void {
    this.applyFilter.emit({
      colIndex: this.colIndex,
      condition: undefined,
      selectedValues: undefined,
    });
    this.closeDropdown.emit();
  }

  onCancel(): void {
    this.closeDropdown.emit();
  }
}
