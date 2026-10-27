import { ChangeDetectionStrategy, Component, EventEmitter, HostListener, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ConditionalFormatRule, FilterConditionType } from '../../models/excel.types';
import { COLORS } from 'shared';
import { parseRange } from '../../services/formula-engine.service';

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
  readonly value = signal<string>('');
  readonly value2 = signal<string>('');
  readonly selectedColorKey = signal<string>('');

  readonly availableColors = [
    { key: 'black', name: 'Black', hex: COLORS['black'] || '#000000' },
    { key: 'darkGray', name: 'Dark Gray', hex: COLORS['darkGray'] || '#434343' },
    { key: 'gray', name: 'Gray', hex: COLORS['gray'] || '#999999' },
    { key: 'white', name: 'White', hex: COLORS['white'] || '#ffffff' },
    { key: 'red', name: 'Red', hex: COLORS['red'] || '#cc0000' },
    { key: 'orange', name: 'Orange', hex: COLORS['orange'] || '#e69138' },
    { key: 'yellow', name: 'Yellow', hex: COLORS['yellow'] || '#f1c232' },
    { key: 'green', name: 'Green', hex: COLORS['green'] || '#6aa84f' },
    { key: 'cyan', name: 'Cyan', hex: COLORS['cyan'] || '#45818e' },
    { key: 'blue', name: 'Blue', hex: COLORS['blue'] || '#3d85c6' },
    { key: 'purple', name: 'Purple', hex: COLORS['purple'] || '#674ea7' },
    { key: 'magenta', name: 'Magenta', hex: COLORS['magenta'] || '#a64d79' },
    { key: 'lightRed', name: 'Light Red', hex: COLORS['lightRed'] || '#f4cccc' },
    { key: 'lightOrange', name: 'Light Orange', hex: COLORS['lightOrange'] || '#fce5cd' },
    { key: 'lightYellow', name: 'Light Yellow', hex: COLORS['lightYellow'] || '#fff2cc' },
    { key: 'lightGreen', name: 'Light Green', hex: COLORS['lightGreen'] || '#d9ead3' },
    { key: 'lightCyan', name: 'Light Cyan', hex: COLORS['lightCyan'] || '#d0e0e3' },
    { key: 'lightBlue', name: 'Light Blue', hex: COLORS['lightBlue'] || '#cfe2f3' },
    { key: 'lightPurple', name: 'Light Purple', hex: COLORS['lightPurple'] || '#d9d2e9' },
    { key: 'lightMagenta', name: 'Light Magenta', hex: COLORS['lightMagenta'] || '#ead1dc' },
  ];

  @HostListener('window:keydown.escape')
  onEscape(): void {
    this.onCancel();
  }

  get isRangeValid(): boolean {
    const r = this.range().trim();
    if (!r) return false;
    return parseRange(r) !== null;
  }

  get canApply(): boolean {
    if (!this.isRangeValid) return false;
    if (!this.selectedColorKey()) return false;
    const cond = this.condition();
    if (cond === 'isEmpty') return true;
    if (cond === 'between') {
      return this.value().trim() !== '' && this.value2().trim() !== '';
    }
    return this.value().trim() !== '';
  }

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
  }

  onApply(): void {
    if (!this.canApply) return;

    const v = this.value().trim();
    const v2 = this.value2().trim();
    const parsedVal = !isNaN(Number(v)) && v !== '' ? Number(v) : v;
    const parsedVal2 = !isNaN(Number(v2)) && v2 !== '' ? Number(v2) : v2;

    this.addRule.emit({
      range: this.range().trim().toUpperCase(),
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
