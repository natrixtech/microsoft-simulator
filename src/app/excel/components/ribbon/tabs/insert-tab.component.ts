import { ChangeDetectionStrategy, Component, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ChartType } from '../../../models/excel.types';
import { ExcelIconComponent } from '../excel-icon.component';

@Component({
  selector: 'app-insert-tab',
  standalone: true,
  imports: [CommonModule, ExcelIconComponent],
  templateUrl: './insert-tab.component.html',
  styleUrls: ['./insert-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class InsertTabComponent {
  @Output() readonly openPivotTableDialog = new EventEmitter<void>();
  @Output() readonly openChartDialog = new EventEmitter<ChartType>();
  @Output() readonly openCommentDialog = new EventEmitter<void>();
}
