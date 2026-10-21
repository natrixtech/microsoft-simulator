import { ChangeDetectionStrategy, Component, EventEmitter, Input, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { BorderAction, CellFormat, ChartType, NumberFormatType, SortOrder, TextAlign } from '../../models/excel.types';
import { ExcelIconComponent } from './excel-icon.component';
import { HomeTabComponent } from './tabs/home-tab.component';
import { InsertTabComponent } from './tabs/insert-tab.component';
import { FormulasTabComponent } from './tabs/formulas-tab.component';
import { DataTabComponent } from './tabs/data-tab.component';

export type ActiveRibbonTab = 'home' | 'insert' | 'formulas' | 'data';

export interface CommandItem {
  id: string;
  label: string;
  icon?: string;
  keywords?: string[];
  run: () => void;
}

@Component({
  selector: 'app-excel-ribbon',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    ExcelIconComponent,
    HomeTabComponent,
    InsertTabComponent,
    FormulasTabComponent,
    DataTabComponent,
  ],
  templateUrl: './ribbon.component.html',
  styleUrls: ['./ribbon.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class RibbonComponent {
  @Input() activeTab: ActiveRibbonTab = 'home';
  @Input() currentFormat: CellFormat = {};
  @Input() hasFilterActive: boolean = false;
  @Input() canUndo: boolean = false;
  @Input() canRedo: boolean = false;

  @Output() readonly undo = new EventEmitter<void>();
  @Output() readonly redo = new EventEmitter<void>();
  @Output() readonly tabChange = new EventEmitter<ActiveRibbonTab>();
  @Output() readonly formatChange = new EventEmitter<Partial<CellFormat>>();
  @Output() readonly borderChange = new EventEmitter<BorderAction>();
  @Output() readonly cut = new EventEmitter<void>();
  @Output() readonly copy = new EventEmitter<void>();
  @Output() readonly paste = new EventEmitter<void>();
  @Output() readonly openConditionalFormat = new EventEmitter<void>();
  @Output() readonly openPivotTable = new EventEmitter<void>();
  @Output() readonly openChart = new EventEmitter<ChartType>();
  @Output() readonly openComment = new EventEmitter<void>();
  @Output() readonly openInsertFunction = new EventEmitter<void>();
  @Output() readonly insertNamedFunction = new EventEmitter<string>();
  @Output() readonly autoSum = new EventEmitter<string | void>();
  @Output() readonly sortAsc = new EventEmitter<void>();
  @Output() readonly sortDesc = new EventEmitter<void>();
  @Output() readonly openSortDialog = new EventEmitter<void>();
  @Output() readonly toggleFilter = new EventEmitter<void>();
  @Output() readonly openDataValidation = new EventEmitter<void>();
  @Output() readonly openRemoveDuplicates = new EventEmitter<void>();
  @Output() readonly focusGrid = new EventEmitter<void>();

  searchQuery: string = '';
  isSearchOpen: boolean = false;
  selectedCommandIndex: number = 0;

  readonly commands: CommandItem[] = [
    {
      id: 'undo',
      label: 'Undo',
      icon: 'undo',
      keywords: ['undo', 'revert'],
      run: () => this.undo.emit(),
    },
    {
      id: 'redo',
      label: 'Redo',
      icon: 'redo',
      keywords: ['redo', 'repeat'],
      run: () => this.redo.emit(),
    },
    {
      id: 'cut',
      label: 'Cut',
      icon: 'cut',
      keywords: ['cut', 'clip'],
      run: () => this.cut.emit(),
    },
    {
      id: 'copy',
      label: 'Copy',
      icon: 'copy',
      keywords: ['copy', 'duplicate'],
      run: () => this.copy.emit(),
    },
    {
      id: 'paste',
      label: 'Paste',
      icon: 'paste',
      keywords: ['paste', 'insert'],
      run: () => this.paste.emit(),
    },
    {
      id: 'bold',
      label: 'Bold',
      icon: 'bold',
      keywords: ['bold', 'font bold', 'heavy'],
      run: () => this.formatChange.emit({ bold: !this.currentFormat.bold }),
    },
    {
      id: 'italic',
      label: 'Italic',
      icon: 'italic',
      keywords: ['italic', 'oblique', 'slant'],
      run: () => this.formatChange.emit({ italic: !this.currentFormat.italic }),
    },
    {
      id: 'underline',
      label: 'Underline',
      icon: 'underline',
      keywords: ['underline'],
      run: () => this.formatChange.emit({ underline: !this.currentFormat.underline }),
    },
    {
      id: 'borders',
      label: 'Borders',
      icon: 'borders',
      keywords: ['borders', 'border', 'all borders', 'gridlines'],
      run: () => this.borderChange.emit('all'),
    },
    {
      id: 'fillColor',
      label: 'Fill Color',
      icon: 'fillColor',
      keywords: ['fill color', 'fill', 'background color', 'shading'],
      run: () => this.formatChange.emit({ fillColor: 'yellow' }),
    },
    {
      id: 'fontColor',
      label: 'Font Color',
      icon: 'fontColor',
      keywords: ['font color', 'text color', 'color'],
      run: () => this.formatChange.emit({ textColor: 'red' }),
    },
    {
      id: 'alignLeft',
      label: 'Align Left',
      icon: 'alignLeft',
      keywords: ['align left', 'left align', 'text align left'],
      run: () => this.formatChange.emit({ align: 'left' }),
    },
    {
      id: 'alignCenter',
      label: 'Align Center',
      icon: 'alignCenter',
      keywords: ['align center', 'center align', 'center text'],
      run: () => this.formatChange.emit({ align: 'center' }),
    },
    {
      id: 'alignRight',
      label: 'Align Right',
      icon: 'alignRight',
      keywords: ['align right', 'right align', 'text align right'],
      run: () => this.formatChange.emit({ align: 'right' }),
    },
    {
      id: 'currency',
      label: 'Currency',
      icon: 'accounting',
      keywords: ['currency', 'accounting', 'dollar', 'money', '$', 'format currency'],
      run: () => this.formatChange.emit({ numberFormat: 'currency' }),
    },
    {
      id: 'percent',
      label: 'Percent',
      icon: 'percent',
      keywords: ['percent', 'percentage', '%', 'format percent'],
      run: () => this.formatChange.emit({ numberFormat: 'percent' }),
    },
    {
      id: 'date',
      label: 'Date',
      icon: 'date',
      keywords: ['date', 'short date', 'format date'],
      run: () => this.formatChange.emit({ numberFormat: 'date' }),
    },
    {
      id: 'conditionalFormatting',
      label: 'Conditional Formatting',
      icon: 'conditionalFormatting',
      keywords: ['conditional formatting', 'highlight cells', 'rules'],
      run: () => this.openConditionalFormat.emit(),
    },
    {
      id: 'autoSum',
      label: 'AutoSum',
      icon: 'autoSum',
      keywords: ['autosum', 'sum', 'total', 'add', 'sigma', 'functions'],
      run: () => this.autoSum.emit('SUM'),
    },
    {
      id: 'sortAsc',
      label: 'Sort A to Z',
      icon: 'sortAsc',
      keywords: ['sort a to z', 'sort ascending', 'sort asc', 'ascending'],
      run: () => this.sortAsc.emit(),
    },
    {
      id: 'sortDesc',
      label: 'Sort Z to A',
      icon: 'sortDesc',
      keywords: ['sort z to a', 'sort descending', 'sort desc', 'descending'],
      run: () => this.sortDesc.emit(),
    },
    {
      id: 'customSort',
      label: 'Custom Sort',
      icon: 'sortDialog',
      keywords: ['sort', 'custom sort', 'sort dialog', 'sort range'],
      run: () => this.openSortDialog.emit(),
    },
    {
      id: 'filter',
      label: 'Filter',
      icon: 'filter',
      keywords: ['filter', 'autofilter', 'toggle filter'],
      run: () => this.toggleFilter.emit(),
    },
    {
      id: 'insertChart',
      label: 'Insert Chart',
      icon: 'columnChart',
      keywords: ['insert chart', 'chart', 'graph', 'column chart'],
      run: () => this.openChart.emit('column'),
    },
    {
      id: 'insertPivotTable',
      label: 'Insert PivotTable',
      icon: 'pivotTable',
      keywords: ['insert pivottable', 'pivottable', 'pivot table'],
      run: () => this.openPivotTable.emit(),
    },
    {
      id: 'newComment',
      label: 'New Comment',
      icon: 'comment',
      keywords: ['new comment', 'comment', 'add comment', 'note'],
      run: () => this.openComment.emit(),
    },
    {
      id: 'dataValidation',
      label: 'Data Validation',
      icon: 'dataValidation',
      keywords: ['data validation', 'validation', 'list validation', 'dropdown in cell'],
      run: () => this.openDataValidation.emit(),
    },
    {
      id: 'removeDuplicates',
      label: 'Remove Duplicates',
      icon: 'removeDuplicates',
      keywords: ['remove duplicates', 'duplicates', 'unique'],
      run: () => this.openRemoveDuplicates.emit(),
    },
    {
      id: 'insertFunction',
      label: 'Insert Function',
      icon: 'insertFunction',
      keywords: ['insert function', 'function', 'fx', 'formula'],
      run: () => this.openInsertFunction.emit(),
    },
  ];

  get filteredCommands(): CommandItem[] {
    const q = this.searchQuery.trim().toLowerCase();
    if (!q) return [];
    return this.commands
      .filter((cmd) => {
        if (cmd.label.toLowerCase().includes(q)) return true;
        return cmd.keywords?.some((k) => k.toLowerCase().includes(q));
      })
      .slice(0, 8);
  }

  onSearchInput(event: Event): void {
    const val = (event.target as HTMLInputElement).value;
    this.searchQuery = val;
    this.isSearchOpen = val.trim().length > 0;
    this.selectedCommandIndex = 0;
  }

  onSearchFocus(): void {
    if (this.searchQuery.trim().length > 0) {
      this.isSearchOpen = true;
    }
  }

  onSearchBlur(): void {
    setTimeout(() => {
      this.isSearchOpen = false;
    }, 200);
  }

  onSearchKeyDown(event: KeyboardEvent): void {
    const list = this.filteredCommands;
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      if (list.length > 0) {
        this.selectedCommandIndex = (this.selectedCommandIndex + 1) % list.length;
      }
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      if (list.length > 0) {
        this.selectedCommandIndex = (this.selectedCommandIndex - 1 + list.length) % list.length;
      }
    } else if (event.key === 'Enter') {
      event.preventDefault();
      if (list.length > 0) {
        const cmd = list[this.selectedCommandIndex] || list[0];
        if (cmd) {
          this.executeCommand(cmd);
        }
      }
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.searchQuery = '';
      this.isSearchOpen = false;
      this.focusGrid.emit();
    }
  }

  executeCommand(cmd: CommandItem): void {
    cmd.run();
    this.searchQuery = '';
    this.isSearchOpen = false;
    this.selectedCommandIndex = 0;
    this.focusGrid.emit();
  }

  selectTab(tab: ActiveRibbonTab): void {
    this.activeTab = tab;
    this.tabChange.emit(tab);
  }
}
