import { ChangeDetectionStrategy, Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelIconComponent } from '../ribbon/excel-icon.component';

@Component({
  selector: 'app-excel-replica-ribbon',
  standalone: true,
  imports: [CommonModule, ExcelIconComponent],
  templateUrl: './excel-replica-ribbon.component.html',
  styleUrls: ['./excel-replica-ribbon.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ExcelReplicaRibbonComponent {
  @Input() activeTab: 'home' | 'insert' | 'formulas' | 'data' = 'home';
  @Input() isBold: boolean = false;
  @Input() isItalic: boolean = false;
  @Input() isUnderline: boolean = false;
  @Input() currentAlign: string = 'left';
  @Input() fontFamily: string = 'Calibri';
  @Input() fontSize: number | string = 11;
  @Input() currentNumberFormat: string = 'automatic';
  @Input() hasFilterActive: boolean = false;
  @Input() isFillColorOpen: boolean = false;
  @Input() isTextColorOpen: boolean = false;

  @Output() readonly tabChange = new EventEmitter<'home' | 'insert' | 'formulas' | 'data'>();

  readonly fontFamilies: string[] = [
    'Calibri',
    'Arial',
    'Segoe UI',
    'Times New Roman',
    'Consolas',
    'Courier New',
    'Georgia',
    'Verdana',
    'Tahoma',
    'Trebuchet MS',
    'Impact',
    'Comic Sans MS',
  ];

  readonly fontSizes: number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  readonly standardColors: Array<{ name: string; hex: string }> = [
    { name: 'darkRed', hex: '#c00000' },
    { name: 'red', hex: '#ff0000' },
    { name: 'orange', hex: '#ffc000' },
    { name: 'yellow', hex: '#ffff00' },
    { name: 'lightGreen', hex: '#92d050' },
    { name: 'green', hex: '#00b050' },
    { name: 'lightBlue', hex: '#00b0f0' },
    { name: 'blue', hex: '#0070c0' },
    { name: 'darkBlue', hex: '#002060' },
    { name: 'purple', hex: '#7030a0' },
  ];

  readonly themeColors: Array<{ name: string; hex: string }> = [
    { name: 'white', hex: '#ffffff' },
    { name: 'black', hex: '#000000' },
    { name: 'gray', hex: '#7f7f7f' },
    { name: 'lightGray', hex: '#d9d9d9' },
    { name: 'accentBlue', hex: '#4472c4' },
    { name: 'accentOrange', hex: '#ed7d31' },
    { name: 'accentGold', hex: '#ffc000' },
    { name: 'accentGreen', hex: '#70ad47' },
  ];

  selectTab(tab: 'home' | 'insert' | 'formulas' | 'data'): void {
    this.tabChange.emit(tab);
  }
}
