import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DataValidationRule } from '../../models/excel.types';

@Component({
  selector: 'app-data-validation-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './data-validation-dialog.component.html',
  styleUrls: ['./data-validation-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class DataValidationDialogComponent implements OnInit {
  @Input() initialRange: string = 'B2:B10';
  @Output() readonly addRule = new EventEmitter<DataValidationRule>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly range = signal<string>('B2:B10');
  readonly itemsText = signal<string>('Yes, No');
  readonly rejectInvalid = signal<boolean>(true);

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
  }

  onApply(): void {
    const raw = this.itemsText();
    const list = raw
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

    this.addRule.emit({
      range: this.range().toUpperCase(),
      type: 'list',
      allowed: list,
      rejectInvalid: this.rejectInvalid(),
    });
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
