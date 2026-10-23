import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  OnInit,
  Output,
  SimpleChanges,
  computed,
  inject,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import {
  DEFAULT_CHART_DATA,
  WordChartData,
  WordChartType,
  renderChartSvg,
} from './chart-svg.helper';

export interface ChartDialogResult {
  chartType: WordChartType;
  title: string;
  data: WordChartData;
}

@Component({
  selector: 'app-word-chart-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './chart-dialog.component.html',
  styleUrls: ['./chart-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ChartDialogComponent implements OnInit, OnChanges {
  private readonly sanitizer = inject(DomSanitizer);

  @Input() isOpen = false;
  @Input() mode: 'insert' | 'change' = 'insert';
  @Input() initialChartType: WordChartType = 'column';
  @Input() initialTitle: string = 'Chart Title';
  @Input() initialData?: WordChartData | null = null;

  @Output() readonly insertChart = new EventEmitter<ChartDialogResult>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly selectedType = signal<WordChartType>('column');
  readonly title = signal<string>('Chart Title');
  readonly data = signal<WordChartData>(JSON.parse(JSON.stringify(DEFAULT_CHART_DATA)));

  // All categories listed in Word
  readonly categoriesList = [
    { type: 'recent', label: 'Recent', enabled: false, icon: 'clock' },
    { type: 'templates', label: 'Templates', enabled: false, icon: 'folder' },
    { type: 'column', label: 'Column', enabled: true, icon: 'column', target: 'chart-type-column' },
    { type: 'line', label: 'Line', enabled: true, icon: 'line', target: 'chart-type-line' },
    { type: 'pie', label: 'Pie', enabled: true, icon: 'pie', target: 'chart-type-pie' },
    { type: 'bar', label: 'Bar', enabled: true, icon: 'bar', target: 'chart-type-bar' },
    { type: 'area', label: 'Area', enabled: false, icon: 'area' },
    { type: 'scatter', label: 'X Y (Scatter)', enabled: false, icon: 'scatter' },
    { type: 'map', label: 'Map', enabled: false, icon: 'map' },
    { type: 'stock', label: 'Stock', enabled: false, icon: 'stock' },
    { type: 'surface', label: 'Surface', enabled: false, icon: 'surface' },
    { type: 'radar', label: 'Radar', enabled: false, icon: 'radar' },
    { type: 'treemap', label: 'Treemap', enabled: false, icon: 'treemap' },
    { type: 'sunburst', label: 'Sunburst', enabled: false, icon: 'sunburst' },
    { type: 'histogram', label: 'Histogram', enabled: false, icon: 'histogram' },
    { type: 'boxwhisker', label: 'Box & Whisker', enabled: false, icon: 'boxwhisker' },
    { type: 'waterfall', label: 'Waterfall', enabled: false, icon: 'waterfall' },
    { type: 'funnel', label: 'Funnel', enabled: false, icon: 'funnel' },
    { type: 'combo', label: 'Combo', enabled: false, icon: 'combo' },
  ];

  readonly subTypeName = computed(() => {
    switch (this.selectedType()) {
      case 'column':
        return 'Clustered Column';
      case 'line':
        return 'Line';
      case 'pie':
        return 'Pie';
      case 'bar':
        return 'Clustered Bar';
      default:
        return 'Clustered Column';
    }
  });

  readonly previewSvg = computed<SafeHtml>(() => {
    const svgStr = renderChartSvg(this.selectedType(), this.title(), this.data(), 450, 240);
    return this.sanitizer.bypassSecurityTrustHtml(svgStr);
  });

  ngOnInit(): void {
    this.resetState();
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isOpen'] && this.isOpen) {
      this.resetState();
    }
  }

  private resetState(): void {
    if (this.initialChartType) {
      this.selectedType.set(this.initialChartType);
    } else {
      this.selectedType.set('column');
    }

    if (this.initialTitle !== undefined && this.initialTitle !== null) {
      this.title.set(this.initialTitle);
    } else {
      this.title.set('Chart Title');
    }

    if (this.initialData && this.initialData.categories?.length && this.initialData.series?.length) {
      this.data.set(JSON.parse(JSON.stringify(this.initialData)));
    } else {
      this.data.set(JSON.parse(JSON.stringify(DEFAULT_CHART_DATA)));
    }
  }

  selectType(type: string): void {
    if (type === 'column' || type === 'line' || type === 'pie' || type === 'bar') {
      this.selectedType.set(type);
    }
  }

  onTitleInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.title.set(input.value);
  }

  onCellInput(sIdx: number, cIdx: number, event: Event): void {
    const input = event.target as HTMLInputElement;
    const raw = input.value.trim();
    const num = parseFloat(raw);
    if (!isNaN(num) && isFinite(num) && num >= 0) {
      this.data.update((curr) => {
        const copy: WordChartData = JSON.parse(JSON.stringify(curr));
        if (copy.series[sIdx]?.values) {
          copy.series[sIdx].values[cIdx] = num;
        }
        return copy;
      });
    } else if (event.type === 'blur') {
      // Invalid number on blur: reset to current stored value
      const old = this.data().series[sIdx]?.values[cIdx] ?? 0;
      input.value = String(old);
    }
  }

  onInsert(): void {
    this.insertChart.emit({
      chartType: this.selectedType(),
      title: this.title().trim() || 'Chart Title',
      data: JSON.parse(JSON.stringify(this.data())),
    });
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }

  @HostListener('window:keydown.escape')
  onEscape(): void {
    if (this.isOpen) {
      this.onCancel();
    }
  }
}
