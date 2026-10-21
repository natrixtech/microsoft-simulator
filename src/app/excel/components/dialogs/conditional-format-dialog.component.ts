import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ConditionalFormatRule, FilterConditionType } from '../../models/excel.types';
import { COLORS } from 'shared';

@Component({
  selector: 'app-conditional-format-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './conditional-format-dialog.component.html',
  styleUrls: ['./conditional-format-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ConditionalFormatDialogComponent implements OnInit {
  @Input() initialRange: string = 'A1:C10';
  @Output() readonly addRule = new EventEmitter<ConditionalFormatRule>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly range = signal<string>('A1:C10');
  readonly condition = signal<FilterConditionType>('greaterThan');
  readonly value = signal<string>('100');
  readonly value2 = signal<string>('');
  readonly selectedColorKey = signal<string>('lightGreen');

  readonly availableColors = [
    { key: 'lightGreen', name: 'Light Green', hex: COLORS['lightGreen'] || '#d9ead3' },
    { key: 'lightRed', name: 'Light Red', hex: COLORS['lightRed'] || '#f4cccc' },
    { key: 'lightYellow', name: 'Light Yellow', hex: COLORS['lightYellow'] || '#fff2cc' },
    { key: 'lightOrange', name: 'Light Orange', hex: COLORS['lightOrange'] || '#fce5cd' },
    { key: 'lightBlue', name: 'Light Blue', hex: COLORS['lightBlue'] || '#cfe2f3' },
    { key: 'lightCyan', name: 'Light Cyan', hex: COLORS['lightCyan'] || '#d0e0e3' },
    { key: 'green', name: 'Green', hex: COLORS['green'] || '#6aa84f' },
    { key: 'yellow', name: 'Yellow', hex: COLORS['yellow'] || '#f1c232' },
    { key: 'red', name: 'Red', hex: COLORS['red'] || '#cc0000' },
  ];

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
  }

  onApply(): void {
    const v = this.value();
    const v2 = this.value2();
    const parsedVal = !isNaN(Number(v)) && v !== '' ? Number(v) : v;
    const parsedVal2 = !isNaN(Number(v2)) && v2 !== '' ? Number(v2) : v2;

    this.addRule.emit({
      range: this.range().toUpperCase(),
      condition: this.condition(),
      value: this.condition() === 'isEmpty' ? undefined : parsedVal,
      value2: this.condition() === 'between' ? parsedVal2 : undefined,
      fillColor: this.selectedColorKey(),
    });
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
