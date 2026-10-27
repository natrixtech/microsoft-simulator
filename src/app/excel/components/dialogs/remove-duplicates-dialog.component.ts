import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelStateService } from '../../services/excel-state.service';
import { ExcelHistoryService } from '../../services/excel-history.service';

@Component({
  selector: 'app-remove-duplicates-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './remove-duplicates-dialog.component.html',
  styleUrls: ['./remove-duplicates-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class RemoveDuplicatesDialogComponent implements OnInit {
  @Input() initialRange: string = 'A1:C20';
  @Output() readonly duplicatesRemoved = new EventEmitter<{ removedCount: number; remainingCount: number }>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly range = signal<string>('A1:C20');
  readonly hasHeader = signal<boolean>(true);
  readonly resultMessage = signal<string | null>(null);

  constructor(
    private state: ExcelStateService,
    private history: ExcelHistoryService
  ) {}

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
  }

  onRemove(): void {
    this.history.push(this.state.getSnapshot());
    const res = this.state.removeDuplicates(this.range().toUpperCase(), this.hasHeader());
    this.resultMessage.set(
      `${res.removedCount} duplicate values found and removed; ${res.remainingCount} unique values remain.`
    );
    this.duplicatesRemoved.emit(res);
  }

  onDismissResult(): void {
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
