import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { COLORS, FONTS } from '../../utils/editor.utils';

@Component({
  selector: 'app-word-replica-ribbon',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './word-replica-ribbon.component.html',
  styleUrls: ['./word-replica-ribbon.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class WordReplicaRibbonComponent {
  @Input() activeTab: string = 'home';
  @Input() activeAction: string = '';
  @Input() hasTable: boolean = false;
  @Input() hasChart: boolean = false;
  @Input() font: string = 'Arial';
  @Input() fontSize: string = '11';
  @Input() fontColor: string = '#000000';
  @Input() highlightColor: string = '#facc15';
  @Input() selectedStyle: string = 'normal';
  @Input() customStyles: Array<{
    name: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    font?: string;
    size?: number;
    color?: string;
    [key: string]: any;
  }> = [];
  @Input() orientation: 'portrait' | 'landscape' = 'portrait';

  @Input() isFontDropdownOpen: boolean = false;
  @Input() isSizeDropdownOpen: boolean = false;
  @Input() isFontColorOpen: boolean = false;
  @Input() isHighlightOpen: boolean = false;
  @Input() isLineSpacingOpen: boolean = false;
  @Input() isStylesMenuOpen: boolean = false;
  @Input() isOrientationOpen: boolean = false;
  @Input() isShadingOpen: boolean = false;
  @Input() isTableInsertOpen: boolean = false;
  @Input() isToolsOpen: boolean = false;
  @Input() extraColors: Array<{ key: string; color: string }> = [];

  @Output() readonly tabChange = new EventEmitter<string>();

  readonly fontList = FONTS;
  readonly sizeList = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  readonly COLORS = COLORS;

  readonly baseColors: Array<{ key: string; name: string; hex: string }> = [
    { key: 'black', name: 'Black', hex: COLORS['black'] },
    { key: 'darkGray', name: 'Dark Gray', hex: COLORS['darkGray'] },
    { key: 'gray', name: 'Gray', hex: COLORS['gray'] },
    { key: 'white', name: 'White', hex: COLORS['white'] },
    { key: 'red', name: 'Red', hex: COLORS['red'] },
    { key: 'orange', name: 'Orange', hex: COLORS['orange'] },
    { key: 'yellow', name: 'Yellow', hex: COLORS['yellow'] },
    { key: 'green', name: 'Green', hex: COLORS['green'] },
    { key: 'cyan', name: 'Cyan', hex: COLORS['cyan'] },
    { key: 'blue', name: 'Blue', hex: COLORS['blue'] },
    { key: 'purple', name: 'Purple', hex: COLORS['purple'] },
    { key: 'magenta', name: 'Magenta', hex: COLORS['magenta'] },
  ];

  readonly lightColors: Array<{ key: string; name: string; hex: string }> = [
    { key: 'lightRed', name: 'Light Red', hex: COLORS['lightRed'] },
    { key: 'lightOrange', name: 'Light Orange', hex: COLORS['lightOrange'] },
    { key: 'lightYellow', name: 'Light Yellow', hex: COLORS['lightYellow'] },
    { key: 'lightGreen', name: 'Light Green', hex: COLORS['lightGreen'] },
    { key: 'lightCyan', name: 'Light Cyan', hex: COLORS['lightCyan'] },
    { key: 'lightBlue', name: 'Light Blue', hex: COLORS['lightBlue'] },
    { key: 'lightPurple', name: 'Light Purple', hex: COLORS['lightPurple'] },
    { key: 'lightMagenta', name: 'Light Magenta', hex: COLORS['lightMagenta'] },
  ];

  setTab(tab: string): void {
    this.tabChange.emit(tab);
  }

  getShadingColors(): Array<{ key: string; color: string }> {
    const list: Array<{ key: string; color: string }> = [
      { key: 'white', color: '#ffffff' },
      { key: 'black', color: '#000000' },
      { key: 'darkGray', color: '#434343' },
      { key: 'gray', color: '#999999' },
      { key: 'blue', color: '#3d85c6' },
      { key: 'orange', color: '#e69138' },
      { key: 'yellow', color: '#f1c232' },
      { key: 'green', color: '#6aa84f' },
      { key: 'cyan', color: '#45818e' },
      { key: 'purple', color: '#674ea7' },
      { key: 'lightRed', color: '#f4cccc' },
      { key: 'red', color: '#cc0000' },
      { key: 'lightOrange', color: '#fce5cd' },
      { key: 'lightYellow', color: '#fff2cc' },
      { key: 'lightGreen', color: '#d9ead3' },
      { key: 'lightCyan', color: '#d0e0e3' },
      { key: 'lightBlue', color: '#cfe2f3' },
      { key: 'lightPurple', color: '#d9d2e9' },
      { key: 'magenta', color: '#a64d79' },
      { key: 'lightMagenta', color: '#ead1dc' },
    ];
    if (this.extraColors && this.extraColors.length > 0) {
      for (const ec of this.extraColors) {
        if (!list.some((item) => item.key === ec.key)) {
          list.push(ec);
        }
      }
    }
    return list;
  }

  hasAnyPopoverOpen(): boolean {
    return (
      this.isFontDropdownOpen ||
      this.isSizeDropdownOpen ||
      this.isFontColorOpen ||
      this.isHighlightOpen ||
      this.isLineSpacingOpen ||
      this.isStylesMenuOpen ||
      this.isOrientationOpen ||
      this.isShadingOpen ||
      this.isTableInsertOpen ||
      this.isToolsOpen
    );
  }
}
