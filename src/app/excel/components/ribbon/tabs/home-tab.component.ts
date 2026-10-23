import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  Output,
  inject,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { BorderAction, CellFormat, NumberFormatType, TextAlign } from '../../../models/excel.types';
import { COLORS } from 'shared';
import { ExcelIconComponent } from '../excel-icon.component';
import { ExcelStateService } from '../../../services/excel-state.service';
import { FormulaEngineService } from '../../../services/formula-engine.service';

@Component({
  selector: 'app-home-tab',
  standalone: true,
  imports: [CommonModule, FormsModule, ExcelIconComponent],
  templateUrl: './home-tab.component.html',
  styleUrls: ['./home-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class HomeTabComponent {
  @Input() currentFormat: CellFormat = {};
  @Input() canUndo: boolean = false;
  @Input() canRedo: boolean = false;

  @Output() readonly undo = new EventEmitter<void>();
  @Output() readonly redo = new EventEmitter<void>();
  @Output() readonly formatChange = new EventEmitter<Partial<CellFormat>>();
  @Output() readonly borderChange = new EventEmitter<BorderAction>();
  @Output() readonly cut = new EventEmitter<void>();
  @Output() readonly copy = new EventEmitter<void>();
  @Output() readonly paste = new EventEmitter<void>();
  @Output() readonly openConditionalFormat = new EventEmitter<void>();
  @Output() readonly autoSum = new EventEmitter<string | void>();
  @Output() readonly openInsertFunction = new EventEmitter<void>();
  @Output() readonly quickSort = new EventEmitter<'asc' | 'desc'>();
  @Output() readonly openSortDialog = new EventEmitter<void>();
  @Output() readonly toggleFilter = new EventEmitter<void>();
  @Output() readonly insertRow = new EventEmitter<void>();
  @Output() readonly insertColumn = new EventEmitter<void>();
  @Output() readonly deleteRow = new EventEmitter<void>();
  @Output() readonly deleteColumn = new EventEmitter<void>();

  private state = inject(ExcelStateService);
  private formulaEngine = inject(FormulaEngineService);
  private elementRef = inject(ElementRef);

  // Pickers state
  readonly showTextColorPicker = signal<boolean>(false);
  readonly showFillColorPicker = signal<boolean>(false);
  readonly showSortMenu = signal<boolean>(false);
  readonly showBordersMenu = signal<boolean>(false);
  readonly showAutoSumMenu = signal<boolean>(false);
  readonly showInsertMenu = signal<boolean>(false);
  readonly showDeleteMenu = signal<boolean>(false);

  readonly fillPickerPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly textPickerPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly sortMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly bordersMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly autoSumMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly insertMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });
  readonly deleteMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });

  readonly fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
  readonly fontFamilies = [
    'Calibri',
    'Arial',
    'Segoe UI',
    'Times New Roman',
    'Courier New',
    'Georgia',
    'Verdana',
    'Comic Sans MS',
  ];

  readonly colorPalette = [
    { key: 'black', hex: COLORS['black'] || '#000000' },
    { key: 'darkGray', hex: COLORS['darkGray'] || '#434343' },
    { key: 'gray', hex: COLORS['gray'] || '#999999' },
    { key: 'white', hex: COLORS['white'] || '#ffffff' },
    { key: 'red', hex: COLORS['red'] || '#cc0000' },
    { key: 'orange', hex: COLORS['orange'] || '#e69138' },
    { key: 'yellow', hex: COLORS['yellow'] || '#f1c232' },
    { key: 'green', hex: COLORS['green'] || '#6aa84f' },
    { key: 'cyan', hex: COLORS['cyan'] || '#45818e' },
    { key: 'blue', hex: COLORS['blue'] || '#3d85c6' },
    { key: 'purple', hex: COLORS['purple'] || '#674ea7' },
    { key: 'magenta', hex: COLORS['magenta'] || '#a64d79' },
    { key: 'lightRed', hex: COLORS['lightRed'] || '#f4cccc' },
    { key: 'lightOrange', hex: COLORS['lightOrange'] || '#fce5cd' },
    { key: 'lightYellow', hex: COLORS['lightYellow'] || '#fff2cc' },
    { key: 'lightGreen', hex: COLORS['lightGreen'] || '#d9ead3' },
    { key: 'lightCyan', hex: COLORS['lightCyan'] || '#d0e0e3' },
    { key: 'lightBlue', hex: COLORS['lightBlue'] || '#cfe2f3' },
    { key: 'lightPurple', hex: COLORS['lightPurple'] || '#d9d2e9' },
    { key: 'lightMagenta', hex: COLORS['lightMagenta'] || '#ead1dc' },
  ];

  get isBold(): boolean {
    return !!this.currentFormat.bold;
  }

  get isItalic(): boolean {
    return !!this.currentFormat.italic;
  }

  get isUnderline(): boolean {
    return !!this.currentFormat.underline;
  }

  get fontFamily(): string {
    return this.currentFormat.fontFamily || 'Calibri';
  }

  setFontFamily(fontFamily: string): void {
    this.formatChange.emit({ fontFamily });
  }

  get fontSize(): number {
    return this.currentFormat.fontSize || 10;
  }

  get currentNumberFormat(): string {
    const f = this.currentFormat.numberFormat;
    if (!f || f === 'automatic' || f === 'number' || (f !== 'currency' && f !== 'percent' && f !== 'date')) {
      return 'automatic';
    }
    return f;
  }

  get currentAlign(): TextAlign {
    return this.currentFormat.align || 'left';
  }

  toggleBold(): void {
    this.formatChange.emit({ bold: !this.isBold });
  }

  toggleItalic(): void {
    this.formatChange.emit({ italic: !this.isItalic });
  }

  toggleUnderline(): void {
    this.formatChange.emit({ underline: !this.isUnderline });
  }

  setFontSize(size: number): void {
    this.formatChange.emit({ fontSize: Number(size) });
  }

  increaseFontSize(): void {
    const cur = this.fontSize;
    const next = this.fontSizes.find((s) => s > cur) || cur + 2;
    this.setFontSize(next);
  }

  decreaseFontSize(): void {
    const cur = this.fontSize;
    const prev = [...this.fontSizes].reverse().find((s) => s < cur) || Math.max(6, cur - 2);
    this.setFontSize(prev);
  }

  setAlign(align: TextAlign): void {
    this.formatChange.emit({ align });
  }

  setNumberFormat(numFmt: NumberFormatType): void {
    this.formatChange.emit({ numberFormat: numFmt });
  }

  onNumberFormatChange(val: string): void {
    if (val === 'automatic') {
      this.formatChange.emit({ numberFormat: 'automatic' });
    } else {
      this.formatChange.emit({ numberFormat: val as NumberFormatType });
    }
  }

  setCommaStyle(): void {
    this.formatChange.emit({ numberFormat: 'number', decimals: 2 });
  }

  increaseDecimal(): void {
    const activeCell = this.state.getCell(this.state.activeCell());
    const curDecimals = this.formulaEngine.getDecimals(activeCell?.value, this.currentFormat);
    this.formatChange.emit({ decimals: curDecimals + 1 });
  }

  decreaseDecimal(): void {
    const activeCell = this.state.getCell(this.state.activeCell());
    const curDecimals = this.formulaEngine.getDecimals(activeCell?.value, this.currentFormat);
    this.formatChange.emit({ decimals: Math.max(0, curDecimals - 1) });
  }

  toggleBordersMenu(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.bordersMenuPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showBordersMenu.update((v) => !v);
    this.showFillColorPicker.set(false);
    this.showTextColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  applyBorder(action: BorderAction): void {
    this.borderChange.emit(action);
    this.showBordersMenu.set(false);
  }

  toggleFillColorPicker(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.fillPickerPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showFillColorPicker.update((v) => !v);
    this.showTextColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showBordersMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  toggleTextColorPicker(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.textPickerPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showTextColorPicker.update((v) => !v);
    this.showFillColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showBordersMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  toggleSortMenu(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.sortMenuPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showSortMenu.update((v) => !v);
    this.showFillColorPicker.set(false);
    this.showTextColorPicker.set(false);
    this.showBordersMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  toggleAutoSumMenu(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.autoSumMenuPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showAutoSumMenu.update((v) => !v);
    this.showFillColorPicker.set(false);
    this.showTextColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showBordersMenu.set(false);
  }

  onSelectAutoSum(fn: string): void {
    this.autoSum.emit(fn);
    this.showAutoSumMenu.set(false);
  }

  onSelectMoreFunctions(): void {
    this.openInsertFunction.emit();
    this.showAutoSumMenu.set(false);
  }

  setTextColor(colorKey: string): void {
    this.formatChange.emit({ textColor: colorKey });
    this.showTextColorPicker.set(false);
  }

  setFillColor(colorKey: string): void {
    this.formatChange.emit({ fillColor: colorKey });
    this.showFillColorPicker.set(false);
  }

  toggleInsertMenu(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.insertMenuPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showInsertMenu.update((v) => !v);
    this.showDeleteMenu.set(false);
    this.showFillColorPicker.set(false);
    this.showTextColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showBordersMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  toggleDeleteMenu(event: MouseEvent): void {
    event.stopPropagation();
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    this.deleteMenuPos.set({ top: rect.bottom + 2, left: rect.left });
    this.showDeleteMenu.update((v) => !v);
    this.showInsertMenu.set(false);
    this.showFillColorPicker.set(false);
    this.showTextColorPicker.set(false);
    this.showSortMenu.set(false);
    this.showBordersMenu.set(false);
    this.showAutoSumMenu.set(false);
  }

  onInsertRowClick(): void {
    this.insertRow.emit();
    this.showInsertMenu.set(false);
  }

  onInsertColClick(): void {
    this.insertColumn.emit();
    this.showInsertMenu.set(false);
  }

  onDeleteRowClick(): void {
    this.deleteRow.emit();
    this.showDeleteMenu.set(false);
  }

  onDeleteColClick(): void {
    this.deleteColumn.emit();
    this.showDeleteMenu.set(false);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (!this.elementRef.nativeElement.contains(target)) {
      this.showFillColorPicker.set(false);
      this.showTextColorPicker.set(false);
      this.showSortMenu.set(false);
      this.showBordersMenu.set(false);
      this.showAutoSumMenu.set(false);
      this.showInsertMenu.set(false);
      this.showDeleteMenu.set(false);
    }
  }
}
