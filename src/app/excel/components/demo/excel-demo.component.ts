import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnDestroy,
  OnInit,
  Output,
  ViewChild,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { TaskGoal } from 'shared';
import { SheetSnapshot } from '../../models/excel.types';
import { colLetterToIndex, formatAddress, indexToColLetter, parseAddress, parseRange } from '../../services/formula-engine.service';
import { COLORS } from 'shared';

export interface DemoStepItem {
  label: string;
  state: 'pending' | 'active' | 'done';
}

@Component({
  selector: 'app-excel-demo',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './excel-demo.component.html',
  styleUrls: ['./excel-demo.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ExcelDemoComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('demoReplicaArea') replicaAreaRef?: ElementRef<HTMLElement>;

  @Input() title: string = '';
  @Input() goals: TaskGoal[] = [];
  @Input() startingSheetSnapshot: SheetSnapshot | null = null;
  @Input() showMeHow?: boolean | { description?: string; sampleSheet?: any };
  @Input() isPaused: boolean = false;

  get description(): string {
    if (typeof this.showMeHow === 'object' && this.showMeHow?.description) {
      return this.showMeHow.description;
    }
    return '';
  }

  @Output() readonly close = new EventEmitter<void>();

  // Demo playback state
  readonly steps = signal<DemoStepItem[]>([]);
  readonly showToast = signal<boolean>(false);
  readonly cursorPixelX = signal<number>(120);
  readonly cursorPixelY = signal<number>(100);
  readonly isClicking = signal<boolean>(false);

  // Scaled replica sheet state
  readonly replicaActiveTab = signal<'home' | 'insert' | 'formulas' | 'data'>('home');
  readonly replicaActiveCell = signal<string>('A1');
  readonly replicaFormulaText = signal<string>('');
  readonly replicaCells = signal<Map<string, { value: any; display: string }>>(new Map());
  readonly replicaFormats = signal<Map<string, any>>(new Map());
  readonly replicaSelectedRange = signal<string>('A1');

  private isCancelled = false;
  private activeTimers: any[] = [];

  constructor(
    private elRef: ElementRef,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.setupSteps();
  }

  ngAfterViewInit(): void {
    this.resetReplicaState();
    this.play();
  }

  ngOnDestroy(): void {
    this.cancelAll();
  }

  onClose(): void {
    this.cancelAll();
    this.close.emit();
  }

  replay(): void {
    this.cancelAll();
    this.resetReplicaState();
    this.setupSteps();
    this.play();
  }

  private cancelAll(): void {
    this.isCancelled = true;
    for (const t of this.activeTimers) {
      clearTimeout(t);
    }
    this.activeTimers = [];
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => {
      let timerId: any = null;

      const finish = () => {
        const idx = this.activeTimers.indexOf(timerId);
        if (idx >= 0) this.activeTimers.splice(idx, 1);
        if (this.isCancelled) {
          resolve();
          return;
        }
        if (this.isPaused) {
          timerId = setTimeout(finish, 50);
          this.activeTimers.push(timerId);
          return;
        }
        resolve();
      };

      timerId = setTimeout(finish, ms);
      this.activeTimers.push(timerId);
    });
  }

  private setupSteps(): void {
    const list: DemoStepItem[] = (this.goals || []).map((g) => ({
      label: this.getStepLabelForGoal(g),
      state: 'pending',
    }));
    this.steps.set(list);
  }

  private getStepLabelForGoal(g: any): string {
    switch (g?.type) {
      case 'cellValue':
        return 'Type the values';
      case 'cellFormula':
        return 'Enter the formula';
      case 'cellFormat':
        return 'Select the range, then Apply the format';
      case 'sorted':
        return 'Sort the data';
      case 'filter':
        return 'Create a filter, then Set the condition';
      case 'chart':
        return 'Insert a chart';
      case 'pivotTable':
        return 'Insert a pivot table';
      case 'comment':
        return 'Add a comment';
      case 'conditionalFormat':
        return 'Add a rule';
      case 'dataValidation':
        return 'Add data validation';
      case 'noDuplicates':
        return 'Remove duplicates';
      case 'action':
        return 'Do the action';
      default:
        return 'Complete task';
    }
  }

  private convertSampleSheet(sheet: any): SheetSnapshot {
    const cellsObj: Record<string, string | number> = {};
    const formatsObj: Record<string, any> = {};
    const widthsObj: Record<string, number> = {};

    if (sheet?.columnWidths && typeof sheet.columnWidths === 'object') {
      for (const [col, w] of Object.entries(sheet.columnWidths)) {
        widthsObj[col.toUpperCase()] = Number(w);
      }
    }

    if (sheet?.formats && typeof sheet.formats === 'object') {
      for (const [rangeKey, fmt] of Object.entries(sheet.formats)) {
        const parsed = parseRange(rangeKey);
        if (parsed) {
          for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
              const addr = formatAddress({ row: r, col: c });
              const existing = formatsObj[addr] || {};
              formatsObj[addr] = { ...existing, ...(fmt as any) };
            }
          }
        }
      }
    }

    if (sheet?.cells && typeof sheet.cells === 'object') {
      for (const [addr, raw] of Object.entries(sheet.cells)) {
        const norm = addr.toUpperCase().replace(/\$/g, '');
        cellsObj[norm] = raw as string | number;
      }
    }

    return {
      cells: cellsObj,
      formats: formatsObj,
      columnWidths: widthsObj,
      charts: [],
      comments: {},
      pivotTables: [],
      filter: null,
      conditionalFormats: [],
      dataValidations: [],
    };
  }

  private resetReplicaState(): void {
    this.isCancelled = false;
    this.showToast.set(false);

    const cellMap = new Map<string, { value: any; display: string }>();
    const fmtMap = new Map<string, any>();

    const sampleSheet =
      typeof this.showMeHow === 'object' && this.showMeHow?.sampleSheet
        ? this.showMeHow.sampleSheet
        : null;

    const snapshot = sampleSheet
      ? this.convertSampleSheet(sampleSheet)
      : this.startingSheetSnapshot;

    if (snapshot) {
      if (snapshot.cells) {
        for (const [addr, raw] of Object.entries(snapshot.cells)) {
          cellMap.set(addr, { value: raw, display: String(raw) });
        }
      }
      if (snapshot.formats) {
        for (const [addr, f] of Object.entries(snapshot.formats)) {
          fmtMap.set(addr, { ...f });
        }
      }
    }

    this.replicaCells.set(cellMap);
    this.replicaFormats.set(fmtMap);
    this.replicaActiveTab.set('home');
    this.replicaActiveCell.set('A1');
    this.replicaSelectedRange.set('A1');
    this.replicaFormulaText.set('');
  }

  private setStepState(idx: number, state: 'pending' | 'active' | 'done'): void {
    this.steps.update((s) => s.map((step, i) => (i === idx ? { ...step, state } : step)));
    this.cdr.markForCheck();
  }

  async play(): Promise<void> {
    await this.delay(400);
    if (this.isCancelled) return;

    const goalsList = this.goals || [];
    if (goalsList.length === 0) {
      this.showToast.set(true);
      return;
    }

    for (let i = 0; i < goalsList.length; i++) {
      if (this.isCancelled) return;
      this.setStepState(i, 'active');
      const g = goalsList[i];

      await this.executeDemoGoal(g);
      if (this.isCancelled) return;

      this.setStepState(i, 'done');
      await this.delay(350);
    }

    if (this.isCancelled) return;
    this.showToast.set(true);
    this.cdr.markForCheck();
  }

  private async moveToTarget(targetName: string): Promise<boolean> {
    if (this.isCancelled) return false;

    const host = this.elRef.nativeElement;
    let target = host.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;

    if (!target) {
      for (let i = 0; i < 5; i++) {
        await this.delay(80);
        if (this.isCancelled) return false;
        target = host.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;
        if (target) break;
      }
    }
    if (!target) return false;

    const replicaArea = this.replicaAreaRef?.nativeElement || host;
    const areaRect = replicaArea.getBoundingClientRect();
    const targetRect = target.getBoundingClientRect();

    const x = targetRect.left - areaRect.left + targetRect.width / 2;
    const y = targetRect.top - areaRect.top + targetRect.height / 2;

    this.cursorPixelX.set(Math.round(x));
    this.cursorPixelY.set(Math.round(y));
    this.cdr.markForCheck();

    await this.delay(350);
    if (this.isCancelled) return false;

    this.isClicking.set(true);
    await this.delay(120);
    this.isClicking.set(false);
    return true;
  }

  private async executeDemoGoal(g: any): Promise<void> {
    const type = g?.type;

    // 1. cellValue
    if (type === 'cellValue') {
      const cells = g.cells as Record<string, any>;
      if (cells) {
        for (const [addr, val] of Object.entries(cells)) {
          this.replicaActiveCell.set(addr);
          this.replicaSelectedRange.set(addr);
          await this.moveToTarget(`replica-cell-${addr}`);

          // Type value character by character into replica cell
          const strVal = String(val);
          let cur = '';
          for (const ch of strVal) {
            if (this.isCancelled) return;
            cur += ch;
            this.replicaFormulaText.set(cur);
            const m = new Map(this.replicaCells());
            m.set(addr, { value: cur, display: cur });
            this.replicaCells.set(m);
            await this.delay(60);
          }
          await this.delay(200);
        }
      }
      return;
    }

    // 2. cellFormula
    if (type === 'cellFormula') {
      const addr = g.cell || 'B5';
      this.replicaActiveCell.set(addr);
      this.replicaSelectedRange.set(addr);
      await this.moveToTarget(`replica-cell-${addr}`);

      const formulaText = g.equals || (g.function ? `=${g.function}(${g.references || 'B2:B4'})` : '=SUM(B2:B4)');
      let cur = '';
      for (const ch of formulaText) {
        if (this.isCancelled) return;
        cur += ch;
        this.replicaFormulaText.set(cur);
        const m = new Map(this.replicaCells());
        m.set(addr, { value: cur, display: cur });
        this.replicaCells.set(m);
        await this.delay(50);
      }
      await this.delay(250);
      return;
    }

    // 3. cellFormat
    if (type === 'cellFormat') {
      this.replicaActiveTab.set('home');
      const range = g.range || 'B2:B4';
      this.replicaSelectedRange.set(range);
      await this.moveToTarget('tab-home');

      if (g.bold) {
        await this.moveToTarget('bold-btn');
        this.applyReplicaFormat(range, { bold: true });
      }
      if (g.italic) {
        await this.moveToTarget('italic-btn');
        this.applyReplicaFormat(range, { italic: true });
      }
      if (g.underline) {
        await this.moveToTarget('underline-btn');
        this.applyReplicaFormat(range, { underline: true });
      }
      if (g.numberFormat === 'currency' || g.numberFormat === '$') {
        await this.moveToTarget('currency-btn');
        this.applyReplicaFormat(range, { numberFormat: 'currency' });
      }
      if (g.numberFormat === 'percent' || g.numberFormat === '%') {
        await this.moveToTarget('percent-btn');
        this.applyReplicaFormat(range, { numberFormat: 'percent' });
      }
      if (g.align) {
        await this.moveToTarget(`align-${g.align}-btn`);
        this.applyReplicaFormat(range, { align: g.align });
      }
      if (g.fillColor) {
        await this.moveToTarget('fill-color-btn');
        this.applyReplicaFormat(range, { fillColor: g.fillColor });
      }
      return;
    }

    // 4. sorted
    if (type === 'sorted') {
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      const order = g.order || 'asc';
      await this.moveToTarget(order === 'desc' ? 'data-sort-desc-btn' : 'data-sort-asc-btn');
      return;
    }

    // 5. filter
    if (type === 'filter') {
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      await this.moveToTarget('data-filter-btn');
      return;
    }

    // 6. chart
    if (type === 'chart') {
      this.replicaActiveTab.set('insert');
      await this.moveToTarget('tab-insert');
      const cType = g.chartType || 'column';
      await this.moveToTarget(`insert-chart-${cType}-btn`);
      return;
    }

    // 7. pivotTable
    if (type === 'pivotTable') {
      this.replicaActiveTab.set('insert');
      await this.moveToTarget('tab-insert');
      await this.moveToTarget('insert-pivottable-btn');
      return;
    }

    // 8. comment
    if (type === 'comment') {
      this.replicaActiveTab.set('insert');
      await this.moveToTarget('tab-insert');
      await this.moveToTarget('insert-comment-btn');
      return;
    }

    // 9. conditionalFormat
    if (type === 'conditionalFormat') {
      this.replicaActiveTab.set('home');
      await this.moveToTarget('tab-home');
      await this.moveToTarget('conditional-formatting-btn');
      return;
    }

    // 10. dataValidation
    if (type === 'dataValidation') {
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      await this.moveToTarget('data-validation-btn');
      return;
    }

    // 11. noDuplicates
    if (type === 'noDuplicates') {
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      await this.moveToTarget('data-remove-duplicates-btn');
      return;
    }

    // 12. action
    if (type === 'action') {
      const act = g.action;
      if (act === 'undo') {
        await this.moveToTarget('home-undo-btn');
      } else if (act === 'copy') {
        this.replicaActiveTab.set('home');
        await this.moveToTarget('copy-btn');
      } else if (act === 'paste') {
        this.replicaActiveTab.set('home');
        await this.moveToTarget('paste-btn');
      } else if (act === 'selectRange' || act === 'goToCell') {
        const range = g.range || 'B2:B4';
        this.replicaSelectedRange.set(range);
        await this.moveToTarget(`replica-cell-${range.split(':')[0]}`);
      }
      return;
    }
  }

  private applyReplicaFormat(rangeStr: string, updates: any): void {
    const parsed = parseRange(rangeStr);
    if (!parsed) return;
    const curFmt = new Map(this.replicaFormats());
    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const addr = formatAddress({ row: r, col: c });
        const existing = curFmt.get(addr) || {};
        curFmt.set(addr, { ...existing, ...updates });
      }
    }
    this.replicaFormats.set(curFmt);
  }

  getToastTitle(): string {
    const t = (this.title || '').trim();
    if (!t) return 'complete this task';
    return t.toLowerCase();
  }
}
