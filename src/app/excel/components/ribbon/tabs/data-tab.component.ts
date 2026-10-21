import { ChangeDetectionStrategy, Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelIconComponent } from '../excel-icon.component';

@Component({
  selector: 'app-data-tab',
  standalone: true,
  imports: [CommonModule, ExcelIconComponent],
  templateUrl: './data-tab.component.html',
  styleUrls: ['./data-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class DataTabComponent {
  @Input() hasFilterActive: boolean = false;

  @Output() readonly sortAsc = new EventEmitter<void>();
  @Output() readonly sortDesc = new EventEmitter<void>();
  @Output() readonly openSortDialog = new EventEmitter<void>();
  @Output() readonly toggleFilter = new EventEmitter<void>();
  @Output() readonly openDataValidation = new EventEmitter<void>();
  @Output() readonly openRemoveDuplicates = new EventEmitter<void>();
}
