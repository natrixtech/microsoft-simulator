import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ChartItem, ChartType } from '../../models/excel.types';

@Component({
  selector: 'app-chart-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './chart-dialog.component.html',
  styleUrls: ['./chart-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ChartDialogComponent implements OnInit {
  @Input() initialRange: string = 'A1:B5';
  @Input() initialChartType: ChartType = 'column';

  @Output() readonly addChart = new EventEmitter<ChartItem>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly chartType = signal<ChartType>('column');
  readonly range = signal<string>('A1:B5');
  readonly title = signal<string>('Chart Title');

  readonly chartTypes: Array<{ type: ChartType; label: string; icon: string }> = [
    { type: 'column', label: 'Column', icon: '📊' },
    { type: 'bar', label: 'Bar', icon: '📉' },
    { type: 'line', label: 'Line', icon: '📈' },
    { type: 'pie', label: 'Pie', icon: '🥧' },
  ];

  ngOnInit(): void {
    if (this.initialRange) {
      this.range.set(this.initialRange);
    }
    if (this.initialChartType) {
      this.chartType.set(this.initialChartType);
    }
  }

  onApply(): void {
    const item: ChartItem = {
      id: 'chart_' + Date.now(),
      chartType: this.chartType(),
      range: this.range().toUpperCase(),
      title: this.title().trim(),
      x: 180,
      y: 120,
      width: 320,
      height: 220,
    };

    this.addChart.emit(item);
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
