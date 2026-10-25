import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnDestroy,
  OnInit,
  Output,
  ViewChild,
  computed,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { COLORS, TaskGoal } from 'shared';
import { SheetSnapshot } from '../../models/excel.types';
import {
  colLetterToIndex,
  formatAddress,
  indexToColLetter,
  parseAddress,
  parseDateParts,
  parseRange,
  FormulaEngineService,
} from '../../services/formula-engine.service';
import { ExcelReplicaRibbonComponent } from './excel-replica-ribbon.component';

export interface DemoStepItem {
  label: string;
  state: 'pending' | 'active' | 'done';
}

// ============================================================================
// Show Me How Timing Constants (Excel replica demo playback)
// ============================================================================
export const CURSOR_MOVE_DURATION_MS = 1100;     // ~2x slower glide to target
export const PRE_CLICK_PAUSE_MS = 500;           // 400–600 ms pause before clicking
export const CLICK_EFFECT_DURATION_MS = 400;     // Visible click pulse effect duration
export const POST_OPEN_PAUSE_MS = 600;           // Pause after menu / dropdown / dialog opens
export const STEP_TRANSITION_PAUSE_MS = 500;     // Pause between actions / steps
export const REPLICA_BASE_WIDTH = 1160;          // Base ribbon width to fit all ribbon groups (Word match)

@Component({
  selector: 'app-excel-demo',
  standalone: true,
  imports: [CommonModule, ExcelReplicaRibbonComponent],
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

  // Scaling state
  readonly replicaReady = signal<boolean>(false);
  readonly scaleTransform = signal<string>('none');
  readonly scalerWidth = signal<string>('100%');
  readonly scalerHeight = signal<string>('100%');
  private resizeObserver?: ResizeObserver;
  private rafId?: number;

  // Scaled replica sheet state
  readonly replicaActiveTab = signal<'home' | 'insert' | 'formulas' | 'data'>('home');
  readonly replicaActiveCell = signal<string>('A1');
  readonly replicaFormulaText = signal<string>('');
  readonly formulaBarText = computed<string>(() => {
    const override = this.replicaFormulaText();
    if (override) return override;
    const active = this.replicaActiveCell();
    const cellVal = this.replicaCells().get(active)?.value;
    if (cellVal !== undefined && cellVal !== null) {
      return String(cellVal);
    }
    return '';
  });
  readonly replicaCells = signal<Map<string, { value: any; calculatedValue?: any; display: string }>>(new Map());
  readonly replicaFormats = signal<Map<string, any>>(new Map());
  readonly replicaSelectedRange = signal<string>('A1');

  // Ribbon state in replica
  readonly replicaFontFamily = signal<string>('Calibri');
  readonly replicaFontSize = signal<number | string>(11);
  readonly replicaNumberFormat = signal<string>('automatic');
  readonly replicaAlign = signal<string>('');
  readonly replicaIsBold = signal<boolean>(false);
  readonly replicaIsItalic = signal<boolean>(false);
  readonly replicaIsUnderline = signal<boolean>(false);
  readonly replicaHasFilter = signal<boolean>(false);
  readonly replicaFillColorOpen = signal<boolean>(false);
  readonly replicaTextColorOpen = signal<boolean>(false);

  // In-replica dialogs & overlays
  readonly replicaDialog = signal<
    'sort' | 'filter' | 'chart' | 'pivotTable' | 'comment' | 'conditionalFormat' | 'dataValidation' | 'removeDuplicates' | null
  >(null);

  // Dialog parameters
  readonly sortColumn = signal<string>('Column B');
  readonly sortOrder = signal<string>('Z to A');
  readonly filterCondition = signal<string>('greaterThan');
  readonly filterValue = signal<string>('50');
  readonly filterColumn = signal<string>('C');
  readonly chartTitle = signal<string>('Sales Report');
  readonly chartType = signal<string>('column');
  readonly pivotRows = signal<string>('Category');
  readonly pivotValues = signal<string>('Amount');
  readonly pivotSummarize = signal<string>('SUM');
  readonly commentText = signal<string>('check pricing');
  readonly cfCondition = signal<string>('greaterThan');
  readonly cfValue = signal<string>('100');
  readonly cfValue2 = signal<string>('');
  readonly cfColor = signal<string>('lightGreen');
  readonly dvType = signal<string>('List');
  readonly dvSource = signal<string>('Yes, No');
  readonly rdHeaders = signal<boolean>(false);
  readonly keyBadge = signal<string | null>(null);
  readonly replicaCopiedRange = signal<string | null>(null);
  readonly commentBalloonPos = signal<{ left: number; top: number } | null>(null);
  readonly pivotCardData = signal<{
    headerValue: string;
    rows: Array<{ label: string; value: string | number }>;
    grandTotal: string | number;
  } | null>(null);
  readonly COLORS = COLORS;
  readonly colorKeys = Object.keys(COLORS);

  private demoUndoStack: Array<{
    cells: Map<string, { value: any; calculatedValue?: any; display: string }>;
    formats: Map<string, any>;
    hiddenRows: Set<number>;
  }> = [];
  private demoRedoStack: Array<{
    cells: Map<string, { value: any; calculatedValue?: any; display: string }>;
    formats: Map<string, any>;
    hiddenRows: Set<number>;
  }> = [];
  private demoClipboard: {
    range: string;
    isCut: boolean;
    cells: Map<string, { value: any; display: string }>;
  } | null = null;

  // Grid visual additions
  readonly hasChartObject = signal<{ title: string; type: string } | null>(null);
  readonly hasPivotTableObject = signal<boolean>(false);
  readonly hasCommentOnCell = signal<string | null>(null);
  readonly hiddenRows = signal<Set<number>>(new Set());

  // Grid structure
  readonly columns = ['A', 'B', 'C', 'D', 'E'];
  rows: number[] = Array.from({ length: 26 }, (_, i) => i + 1);

  private formulaEngine = new FormulaEngineService();
  private isCancelled = false;
  private activeTimers: any[] = [];

  constructor(
    private elRef: ElementRef,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.resetReplicaState();
    this.setupSteps();
  }

  ngAfterViewInit(): void {
    if (typeof ResizeObserver !== 'undefined' && this.replicaAreaRef) {
      this.resizeObserver = new ResizeObserver(() => this.updateScaling());
      this.resizeObserver.observe(this.replicaAreaRef.nativeElement);
    }

    let frameCount = 0;
    const checkScale = () => {
      if (this.isCancelled) return;
      frameCount++;
      const scaled = this.updateScaling();
      if (scaled) {
        this.replicaReady.set(true);
        this.cdr.detectChanges();
        this.play();
      } else if (frameCount < 20) {
        this.rafId = requestAnimationFrame(checkScale);
      } else {
        this.replicaReady.set(true);
        this.cdr.detectChanges();
        this.play();
      }
    };

    checkScale();
  }

  ngOnDestroy(): void {
    this.cancelAll();
    this.resizeObserver?.disconnect();
  }

  onClose(): void {
    this.cancelAll();
    this.close.emit();
  }

  replay(): void {
    this.cancelAll();
    this.isCancelled = false;
    this.showToast.set(false);
    this.resetReplicaState();
    this.setupSteps();
    this.updateScaling();
    this.play();
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    this.updateScaling();
  }

  updateScaling(): boolean {
    if (!this.replicaAreaRef) return false;
    const el = this.replicaAreaRef.nativeElement;
    const w = el.clientWidth;
    const h = el.clientHeight;
    if (w <= 0 || h <= 0) return false;

    let scaledH = h;
    if (w >= REPLICA_BASE_WIDTH) {
      this.scaleTransform.set('none');
      this.scalerWidth.set('100%');
      this.scalerHeight.set('100%');
    } else {
      const scale = w / REPLICA_BASE_WIDTH;
      scaledH = Math.round(h / scale);
      this.scaleTransform.set(`scale(${scale})`);
      this.scalerWidth.set(`${REPLICA_BASE_WIDTH}px`);
      this.scalerHeight.set(`${scaledH}px`);
    }

    // Chrome inside replica: titlebar (30) + ribbon (129) + formula bar (27) = 186px
    // Grid header: 24px. Each data row: 26px.
    const neededRows = Math.max(26, Math.ceil((scaledH - 210) / 26) + 1);
    if (this.rows.length !== neededRows) {
      this.rows = Array.from({ length: neededRows }, (_, i) => i + 1);
    }
    this.cdr.markForCheck();
    return true;
  }

  private cancelAll(): void {
    this.isCancelled = true;
    if (this.rafId) {
      cancelAnimationFrame(this.rafId);
      this.rafId = undefined;
    }
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
        return 'Type values into cells';
      case 'cellFormula':
        return g.function ? `Enter =${g.function}(${g.references || ''}) formula` : 'Enter the formula';
      case 'cellFormat':
        if (g.numberFormat === 'currency' || g.numberFormat === '$') return 'Format range as currency ($)';
        if (g.numberFormat === 'percent' || g.numberFormat === '%') return 'Format range as percent (%)';
        if (g.fontFamily) return `Set font family to ${g.fontFamily}`;
        if (g.bold) return 'Apply bold formatting';
        if (g.fillColor) return 'Apply fill color';
        return 'Format cells on Home tab';
      case 'sorted':
        return `Sort range by column ${g.byColumn || 'B'} (${g.order === 'desc' ? 'Z to A' : 'A to Z'})`;
      case 'filter':
        return `Filter column ${g.column || 'C'} (${g.condition || 'condition'} ${g.value ?? ''})`;
      case 'chart':
        return `Insert ${g.chartType || 'column'} chart`;
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
    this.hasChartObject.set(null);
    this.hasPivotTableObject.set(false);
    this.pivotCardData.set(null);
    this.hasCommentOnCell.set(null);
    this.commentBalloonPos.set(null);
    this.hiddenRows.set(new Set());
    this.replicaDialog.set(null);
    this.replicaFillColorOpen.set(false);
    this.replicaTextColorOpen.set(false);
    this.replicaHasFilter.set(false);
    this.rdHeaders.set(false);
    this.dvType.set('List');
    this.dvSource.set('');
    this.cfCondition.set('');
    this.cfValue.set('');
    this.cfValue2.set('');
    this.cfColor.set('');
    this.pivotRows.set('');
    this.pivotValues.set('');
    this.pivotSummarize.set('');
    this.keyBadge.set(null);
    this.replicaCopiedRange.set(null);
    this.demoUndoStack = [];
    this.demoRedoStack = [];
    this.demoClipboard = null;

    const cellMap = new Map<string, { value: any; calculatedValue?: any; display: string }>();
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

    // Evaluate formulas and format displays initially
    this.evaluateAndFormatCells();
  }

  private evaluateAndFormatCells(): void {
    const cur = new Map(this.replicaCells());
    const fmts = this.replicaFormats();

    const getVal = (addr: string): any => {
      const norm = addr.toUpperCase().replace(/\$/g, '');
      const c = cur.get(norm);
      if (!c) return 0;
      if (typeof c.value === 'string' && c.value.startsWith('=')) {
        return c.calculatedValue ?? 0;
      }
      return c.value ?? 0;
    };

    const getRange = (rangeStr: string): any[][] => {
      const parsed = parseRange(rangeStr);
      if (!parsed) return [['#REF!']];
      const res: any[][] = [];
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        const rowVals: any[] = [];
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          const a = formatAddress({ row: r, col: c });
          const norm = a.toUpperCase().replace(/\$/g, '');
          const cell = cur.get(norm);
          if (!cell || cell.value === '' || cell.value === undefined || cell.value === null) {
            rowVals.push(null);
          } else {
            rowVals.push(getVal(a));
          }
        }
        res.push(rowVals);
      }
      return res;
    };

    // 1. Evaluate formula cells
    for (const [addr, cell] of cur.entries()) {
      if (typeof cell.value === 'string' && cell.value.startsWith('=')) {
        const evalRes = this.formulaEngine.evaluate(cell.value, addr, getVal, getRange);
        cell.calculatedValue = evalRes.value;
      }
    }

    // 2. Format display values
    for (const [addr, cell] of cur.entries()) {
      const fmt = fmts.get(addr);
      const valToFormat =
        typeof cell.value === 'string' && cell.value.startsWith('=')
          ? cell.calculatedValue
          : cell.value;
      cell.display = this.formulaEngine.formatDisplayValue(
        valToFormat,
        fmt?.numberFormat,
        fmt?.decimals
      );
    }

    this.replicaCells.set(new Map(cur));
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
    this.evaluateAndFormatCells();
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
      await this.delay(STEP_TRANSITION_PAUSE_MS);
    }

    if (this.isCancelled) return;
    this.showToast.set(true);
    this.cdr.markForCheck();
  }

  private positionCursorOver(targetRect: DOMRect): void {
    const area = this.replicaAreaRef?.nativeElement;
    if (!area) return;
    const areaRect = area.getBoundingClientRect();
    const x = targetRect.left - areaRect.left + (targetRect.width > 0 ? targetRect.width / 2 : 10);
    const y = targetRect.top - areaRect.top + (targetRect.height > 0 ? targetRect.height / 2 : 10);

    this.cursorPixelX.set(Math.round(x));
    this.cursorPixelY.set(Math.round(y));
    this.cdr.markForCheck();
  }

  private async moveToTarget(targetName: string): Promise<boolean> {
    if (this.isCancelled) return false;

    // Switch ribbon tab automatically if targeting tab buttons or tab controls
    if (targetName === 'tab-insert' || targetName.startsWith('insert-')) {
      this.replicaActiveTab.set('insert');
    } else if (targetName === 'tab-formulas' || targetName.startsWith('formulas-')) {
      this.replicaActiveTab.set('formulas');
    } else if (targetName === 'tab-data' || targetName.startsWith('data-')) {
      this.replicaActiveTab.set('data');
    } else if (
      targetName === 'tab-home' ||
      [
        'bold-btn',
        'italic-btn',
        'underline-btn',
        'borders-btn',
        'fill-color-btn',
        'text-color-btn',
        'font-color-btn',
        'font-family-select',
        'font-size-select',
        'increase-font-size-btn',
        'decrease-font-size-btn',
        'align-left-btn',
        'align-center-btn',
        'align-right-btn',
        'format-currency-btn',
        'currency-btn',
        'format-percent-btn',
        'percent-btn',
        'conditional-format-btn',
        'conditional-formatting-btn',
        'paste-btn',
        'cut-btn',
        'copy-btn',
        'autosum-btn',
        'sort-filter-btn',
        'qa-undo-btn',
        'qa-redo-btn',
        'home-undo-btn',
      ].includes(targetName)
    ) {
      this.replicaActiveTab.set('home');
    }
    this.cdr.markForCheck();

    const host = this.elRef.nativeElement;
    let target = (host.querySelector(`[data-demo-target="${targetName}"]`) ||
      host.querySelector(`[data-demo-target-alt="${targetName}"]`)) as HTMLElement;

    if (!target) {
      for (let i = 0; i < 8; i++) {
        await this.delay(60);
        if (this.isCancelled) return false;
        target = (host.querySelector(`[data-demo-target="${targetName}"]`) ||
          host.querySelector(`[data-demo-target-alt="${targetName}"]`)) as HTMLElement;
        if (target) break;
      }
    }
    if (!target) return false;

    this.positionCursorOver(target.getBoundingClientRect());

    // Glide smoothly to target (~2x slower pace)
    await this.delay(CURSOR_MOVE_DURATION_MS);
    if (this.isCancelled) return false;

    // Short pause (400-600ms) before clicking
    await this.delay(PRE_CLICK_PAUSE_MS);
    if (this.isCancelled) return false;

    // Visible click effect
    this.isClicking.set(true);
    await this.delay(CLICK_EFFECT_DURATION_MS);
    this.isClicking.set(false);
    return true;
  }

  selectRange(rangeStr: string): void {
    this.replicaSelectedRange.set(rangeStr);
    const first = rangeStr.split(':')[0];
    this.replicaActiveCell.set(first);
    this.cdr.markForCheck();
  }

  isCellInRange(addr: string, rangeStr: string | null | undefined): boolean {
    if (!rangeStr) return false;
    const parsed = parseRange(rangeStr);
    const cellCoord = parseAddress(addr);
    if (!parsed || !cellCoord) return addr === rangeStr;
    return (
      cellCoord.row >= parsed.startRow &&
      cellCoord.row <= parsed.endRow &&
      cellCoord.col >= parsed.startCol &&
      cellCoord.col <= parsed.endCol
    );
  }

  private async showKeyBadge(shortcut: string, durationMs = 1000): Promise<void> {
    if (this.isCancelled) return;
    this.keyBadge.set(shortcut);
    await this.delay(durationMs);
    this.keyBadge.set(null);
  }

  private captureDemoSnapshot(): {
    cells: Map<string, { value: any; calculatedValue?: any; display: string }>;
    formats: Map<string, any>;
    hiddenRows: Set<number>;
  } {
    const cellsCopy = new Map<string, { value: any; calculatedValue?: any; display: string }>();
    for (const [k, v] of this.replicaCells().entries()) {
      cellsCopy.set(k, { ...v });
    }
    const formatsCopy = new Map<string, any>();
    for (const [k, v] of this.replicaFormats().entries()) {
      formatsCopy.set(k, { ...v });
    }
    return {
      cells: cellsCopy,
      formats: formatsCopy,
      hiddenRows: new Set(this.hiddenRows()),
    };
  }

  private restoreDemoSnapshot(snap: {
    cells: Map<string, { value: any; calculatedValue?: any; display: string }>;
    formats: Map<string, any>;
    hiddenRows: Set<number>;
  }): void {
    this.replicaCells.set(snap.cells);
    this.replicaFormats.set(snap.formats);
    this.hiddenRows.set(snap.hiddenRows);
    this.evaluateAndFormatCells();
    this.cdr.markForCheck();
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

          // Type value character by character into formula bar and cell
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
          await this.delay(PRE_CLICK_PAUSE_MS);
          this.evaluateAndFormatCells();
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
      await this.delay(PRE_CLICK_PAUSE_MS);
      // On Enter commit: calculate formula result and format
      const m = new Map(this.replicaCells());
      m.set(addr, { value: formulaText, display: formulaText });
      this.replicaCells.set(m);
      this.evaluateAndFormatCells();
      await this.delay(250);
      return;
    }

    // 3. cellFormat
    if (type === 'cellFormat') {
      const range = g.range || 'B2:B4';
      this.selectRange(range);
      await this.moveToTarget(`replica-cell-${range.split(':')[0]}`);
      await this.moveToTarget('tab-home');

      if (g.numberFormat === 'currency' || g.numberFormat === '$') {
        await this.moveToTarget('format-currency-btn');
        this.replicaNumberFormat.set('currency');
        this.applyReplicaFormat(range, { numberFormat: 'currency' });
      } else if (g.numberFormat === 'percent' || g.numberFormat === '%') {
        await this.moveToTarget('format-percent-btn');
        this.replicaNumberFormat.set('percent');
        this.applyReplicaFormat(range, { numberFormat: 'percent' });
      } else if (g.numberFormat === 'date') {
        await this.moveToTarget('number-format-select');
        this.replicaNumberFormat.set('date');
        await this.delay(POST_OPEN_PAUSE_MS);
        this.applyReplicaFormat(range, { numberFormat: 'date' });
      }

      if (g.bold) {
        await this.moveToTarget('bold-btn');
        this.replicaIsBold.set(true);
        this.applyReplicaFormat(range, { bold: true });
      }
      if (g.italic) {
        await this.moveToTarget('italic-btn');
        this.replicaIsItalic.set(true);
        this.applyReplicaFormat(range, { italic: true });
      }
      if (g.underline) {
        await this.moveToTarget('underline-btn');
        this.replicaIsUnderline.set(true);
        this.applyReplicaFormat(range, { underline: true });
      }
      if (g.align) {
        await this.moveToTarget(`align-${g.align}-btn`);
        this.replicaAlign.set(g.align);
        this.applyReplicaFormat(range, { align: g.align });
      }
      if (g.fontFamily) {
        await this.moveToTarget('font-family-select');
        this.replicaFontFamily.set(g.fontFamily);
        await this.delay(POST_OPEN_PAUSE_MS);
        this.applyReplicaFormat(range, { fontFamily: g.fontFamily });
      }
      if (g.fontSize) {
        await this.moveToTarget('font-size-select');
        this.replicaFontSize.set(g.fontSize);
        await this.delay(POST_OPEN_PAUSE_MS);
        this.applyReplicaFormat(range, { fontSize: g.fontSize });
      }
      if (g.fillColor) {
        await this.moveToTarget('fill-color-btn');
        this.replicaFillColorOpen.set(true);
        await this.delay(POST_OPEN_PAUSE_MS);
        await this.moveToTarget(`demo-swatch-${g.fillColor || 'lightGreen'}`);
        this.replicaFillColorOpen.set(false);
        this.applyReplicaFormat(range, { fillColor: g.fillColor });
      }
      if (g.color || g.textColor) {
        await this.moveToTarget('text-color-btn');
        this.replicaTextColorOpen.set(true);
        await this.delay(POST_OPEN_PAUSE_MS);
        await this.moveToTarget(`demo-text-swatch-${g.color || 'red'}`);
        this.replicaTextColorOpen.set(false);
        this.applyReplicaFormat(range, { color: g.color || g.textColor });
      }
      return;
    }

    // 4. sorted
    if (type === 'sorted') {
      const range = g.range || 'A1:B5';
      this.selectRange(range);
      await this.moveToTarget(`replica-cell-${range.split(':')[0]}`);
      await this.moveToTarget('tab-data');

      // Open Sort dialog
      await this.moveToTarget('data-sort-dialog-btn');
      this.sortColumn.set('Column ' + (g.byColumn || 'B'));
      this.sortOrder.set(g.order === 'desc' ? 'Z to A' : 'A to Z');
      this.replicaDialog.set('sort');
      await this.delay(POST_OPEN_PAUSE_MS);

      await this.moveToTarget('sort-dialog-ok-btn');
      this.replicaDialog.set(null);

      // Reorder rows descending in replica grid
      const m = new Map(this.replicaCells());
      if (g.order === 'desc') {
        m.set('A2', { value: 'Desk', display: 'Desk' });
        m.set('B2', { value: 150, display: '150' });
        m.set('A3', { value: 'Chair', display: 'Chair' });
        m.set('B3', { value: 75, display: '75' });
        m.set('A4', { value: 'Lamp', display: 'Lamp' });
        m.set('B4', { value: 35, display: '35' });
        m.set('A5', { value: 'Pen', display: 'Pen' });
        m.set('B5', { value: 2, display: '2' });
      }
      this.replicaCells.set(m);
      return;
    }

    // 5. filter
    if (type === 'filter') {
      const range = g.range || 'A1:C5';
      this.selectRange(range);
      await this.moveToTarget(`replica-cell-${range.split(':')[0]}`);
      await this.moveToTarget('tab-data');

      // Click filter toggle button
      await this.moveToTarget('data-filter-btn');
      this.replicaHasFilter.set(true);
      await this.delay(POST_OPEN_PAUSE_MS);

      // Click column filter arrow
      const col = g.column || 'C';
      await this.moveToTarget(`col-filter-btn-${col}`);
      this.filterColumn.set(col);
      this.filterCondition.set(g.condition || 'greaterThan');
      this.filterValue.set(String(g.value ?? 50));
      this.replicaDialog.set('filter');
      await this.delay(POST_OPEN_PAUSE_MS);

      // Click OK on filter dialog
      await this.moveToTarget('filter-dialog-ok-btn');
      this.replicaDialog.set(null);

      // Hide rows where col C <= 50 (Row 3: 40, Row 5: 30)
      const h = new Set<number>();
      h.add(3);
      h.add(5);
      this.hiddenRows.set(h);
      return;
    }

    // 6. chart
    if (type === 'chart') {
      const range = g.range || 'A1:B5';
      this.selectRange(range);
      await this.moveToTarget(`replica-cell-${range.split(':')[0]}`);
      await this.moveToTarget('tab-insert');

      const cType = g.chartType || 'column';
      await this.moveToTarget(`insert-chart-${cType}-btn`);
      this.chartType.set(cType);
      this.chartTitle.set(g.title || 'Sales Report');
      this.replicaDialog.set('chart');
      await this.delay(POST_OPEN_PAUSE_MS);

      await this.moveToTarget('chart-dialog-insert-btn');
      this.replicaDialog.set(null);
      this.hasChartObject.set({ title: g.title || 'Sales Report', type: cType });
      return;
    }

    // 7. pivotTable
    if (type === 'pivotTable') {
      const range = g.range || 'A1:C6';
      const parts = range.split(':');
      const start = parts[0];
      const end = parts[1] || start;

      // Selection appears as cursor moves to start, then drags to end
      await this.moveToTarget(`replica-cell-${start}`);
      this.selectRange(start);
      if (end !== start) {
        await this.moveToTarget(`replica-cell-${end}`);
        this.selectRange(range);
      }

      // Compute pivot table data dynamically from replica sheet
      const pData = this.computePivotTableData(range, g.rowsField, g.valuesField, g.summarize);

      // Insert tab -> PivotTable
      this.replicaActiveTab.set('insert');
      await this.moveToTarget('tab-insert');
      await this.moveToTarget('insert-pivottable-btn');

      // Dialog opens with empty fields
      this.pivotRows.set('');
      this.pivotValues.set('');
      this.pivotSummarize.set('');
      this.replicaDialog.set('pivotTable');
      await this.delay(POST_OPEN_PAUSE_MS);

      // Click pivot-rows-select and pick rowsField
      await this.moveToTarget('pivot-rows-select');
      this.pivotRows.set(pData.resolvedRowsField);
      this.cdr.detectChanges();
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click pivot-values-select and pick valuesField
      await this.moveToTarget('pivot-values-select');
      this.pivotValues.set(pData.resolvedValuesField);
      this.cdr.detectChanges();
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click pivot-summarize-select and pick summarize (default SUM)
      await this.moveToTarget('pivot-summarize-select');
      this.pivotSummarize.set(pData.resolvedSummarize);
      this.cdr.detectChanges();
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click OK
      await this.moveToTarget('pivottable-dialog-ok-btn');
      this.replicaDialog.set(null);

      // Result card computed from replica sheet
      this.pivotCardData.set(pData);
      this.hasPivotTableObject.set(true);
      return;
    }

    // 8. comment
    if (type === 'comment') {
      const cellAddr = g.cell || 'C4';

      // Cursor moves to cell FIRST, then cell is selected
      await this.moveToTarget(`replica-cell-${cellAddr}`);
      this.selectRange(cellAddr);

      // Insert tab -> Comment
      this.replicaActiveTab.set('insert');
      await this.moveToTarget('tab-insert');
      await this.moveToTarget('insert-comment-btn');

      // Dialog opens EMPTY
      this.commentText.set('');
      this.replicaDialog.set('comment');
      await this.delay(POST_OPEN_PAUSE_MS);

      // Click comment-input and type text char by char
      await this.moveToTarget('comment-input');
      const targetText = g.contains || 'This value needs checking.';
      let cur = '';
      for (const ch of targetText) {
        if (this.isCancelled) return;
        cur += ch;
        this.commentText.set(cur);
        await this.delay(60);
      }
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click comment-post-btn
      await this.moveToTarget('comment-post-btn');
      this.replicaDialog.set(null);

      // Position balloon to the right of that cell, top-aligned with it
      const ws = this.elRef.nativeElement.querySelector('.replica-grid-workspace') as HTMLElement | null;
      const cellEl = this.elRef.nativeElement.querySelector(`[data-demo-target="replica-cell-${cellAddr}"]`) as HTMLElement | null;
      let top = 0;
      let left = 0;
      if (cellEl && ws) {
        let el: HTMLElement | null = cellEl;
        let foundWs = false;
        while (el) {
          if (el === ws) {
            foundWs = true;
            break;
          }
          top += el.offsetTop;
          left += el.offsetLeft;
          el = el.offsetParent as HTMLElement | null;
        }
        if (foundWs) {
          left += cellEl.offsetWidth + 6;
        } else {
          const wsRect = ws.getBoundingClientRect();
          const cellRect = cellEl.getBoundingClientRect();
          const scale = ws.offsetWidth > 0 ? wsRect.width / ws.offsetWidth : 1;
          top = (cellRect.top - wsRect.top) / scale;
          left = (cellRect.right - wsRect.left) / scale + 6;
        }
      } else {
        const rowNum = parseInt(cellAddr.replace(/^[A-Z]+/i, ''), 10) || 1;
        const colLetter = cellAddr.match(/^[A-Z]+/i)?.[0]?.toUpperCase() || 'A';
        const colIdx = colLetter.charCodeAt(0) - 65;
        top = 24 + (rowNum - 1) * 26;
        left = 38 + (colIdx + 1) * 90 + 6;
      }

      this.commentBalloonPos.set({ left, top });
      this.hasCommentOnCell.set(cellAddr);
      return;
    }

    // 9. conditionalFormat
    if (type === 'conditionalFormat') {
      const range = g.range || 'C2:C8';
      const parts = range.split(':');
      const start = parts[0];
      const end = parts[1] || start;

      // Selection appears as cursor moves to start, then drags to end
      await this.moveToTarget(`replica-cell-${start}`);
      this.selectRange(start);
      if (end !== start) {
        await this.moveToTarget(`replica-cell-${end}`);
        this.selectRange(range);
      }

      // Home tab -> Conditional Formatting
      this.replicaActiveTab.set('home');
      await this.moveToTarget('tab-home');
      await this.moveToTarget('conditional-format-btn');

      // Dialog opens with empty/default fields
      this.cfCondition.set('');
      this.cfValue.set('');
      this.cfValue2.set('');
      this.cfColor.set('');
      this.replicaDialog.set('conditionalFormat');
      await this.delay(POST_OPEN_PAUSE_MS);

      // Click cf-condition-select and pick condition
      const targetCondition = g.condition || 'greaterThan';
      await this.moveToTarget('cf-condition-select');
      this.cfCondition.set(targetCondition);
      this.cdr.detectChanges();
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click cf-value-input and type value
      const targetValue = g.value !== undefined ? g.value : 100;
      if (targetCondition !== 'isEmpty' && targetValue !== undefined) {
        await this.moveToTarget('cf-value-input');
        let cur = '';
        for (const ch of String(targetValue)) {
          if (this.isCancelled) return;
          cur += ch;
          this.cfValue.set(cur);
          await this.delay(60);
        }
        await this.delay(PRE_CLICK_PAUSE_MS);
      }

      // Click cf-value2-input if between
      if (targetCondition === 'between' && g.value2 !== undefined) {
        await this.moveToTarget('cf-value2-input');
        let cur2 = '';
        for (const ch of String(g.value2)) {
          if (this.isCancelled) return;
          cur2 += ch;
          this.cfValue2.set(cur2);
          await this.delay(60);
        }
        await this.delay(PRE_CLICK_PAUSE_MS);
      }

      // Click swatch cf-swatch-<goal.fillColor>
      const targetColorKey = g.fillColor || 'lightGreen';
      await this.moveToTarget(`cf-swatch-${targetColorKey}`);
      this.cfColor.set(targetColorKey);
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Click Done
      await this.moveToTarget('cf-dialog-done-btn');
      this.replicaDialog.set(null);

      // Colour ONLY the cells in goal.range that match the condition
      const parsed = parseRange(range);
      if (parsed) {
        const fmts = new Map(this.replicaFormats());
        const exactColor = COLORS[targetColorKey] || targetColorKey;
        for (let r = parsed.startRow; r <= parsed.endRow; r++) {
          for (let c = parsed.startCol; c <= parsed.endCol; c++) {
            const addr = formatAddress({ row: r, col: c });
            const cell = this.replicaCells().get(addr);
            const cellVal = cell?.calculatedValue !== undefined ? cell.calculatedValue : cell?.value;
            if (this.evaluateDemoCondition(cellVal, targetCondition, targetValue, g.value2)) {
              const ex = fmts.get(addr) || {};
              fmts.set(addr, { ...ex, fillColor: exactColor });
            }
          }
        }
        this.replicaFormats.set(fmts);
      }
      return;
    }

    // 10. dataValidation
    if (type === 'dataValidation') {
      const range = g.range || 'B2:B6';
      const parts = range.split(':');
      const start = parts[0];
      const end = parts[1] || start;

      // Selection appears as cursor drags, not before
      await this.moveToTarget(`replica-cell-${start}`);
      this.selectRange(start);
      if (end !== start) {
        await this.moveToTarget(`replica-cell-${end}`);
        this.selectRange(range);
      }

      // Data tab -> Data Validation button
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      await this.moveToTarget('data-validation-btn');

      // Dialog opens with criteria empty or default
      this.dvType.set('Any value');
      this.dvSource.set('');
      this.replicaDialog.set('dataValidation');
      await this.delay(POST_OPEN_PAUSE_MS);

      // The cursor clicks criteria control and chooses "List"
      await this.moveToTarget('dv-allow-select');
      this.dvType.set('List');
      await this.delay(PRE_CLICK_PAUSE_MS);

      // The cursor clicks values field and types goal.allowed joined with ", " char by char (fallback "Yes, No")
      const allowedText = Array.isArray(g.allowed)
        ? g.allowed.join(', ')
        : (g.allowed ? String(g.allowed) : 'Yes, No');

      await this.moveToTarget('dv-source-input');
      let cur = '';
      for (const ch of allowedText) {
        if (this.isCancelled) return;
        cur += ch;
        this.dvSource.set(cur);
        await this.delay(60);
      }
      await this.delay(PRE_CLICK_PAUSE_MS);

      // Clicks OK / Done
      await this.moveToTarget('dv-dialog-done-btn');
      this.replicaDialog.set(null);

      // Dropdown arrows on the range
      const fmts = new Map(this.replicaFormats());
      const parsed = parseRange(range);
      if (parsed) {
        for (let r = parsed.startRow; r <= parsed.endRow; r++) {
          for (let c = parsed.startCol; c <= parsed.endCol; c++) {
            const a = formatAddress({ row: r, col: c });
            const ex = fmts.get(a) || {};
            fmts.set(a, { ...ex, hasDataValidation: true });
          }
        }
      }
      this.replicaFormats.set(fmts);
      return;
    }

    // 11. noDuplicates
    if (type === 'noDuplicates') {
      const rawRange = g.range || 'A1:C6';
      const parsed = parseRange(rawRange);

      let hasHeaderIncluded = false;
      let selectRangeStr = rawRange;

      if (parsed) {
        if (parsed.startRow > 0) {
          // Check if row directly above has values in replica sheet (a header row)
          let hasAbove = false;
          for (let c = parsed.startCol; c <= parsed.endCol; c++) {
            const a = formatAddress({ row: parsed.startRow - 1, col: c });
            const val = this.replicaCells().get(a)?.value;
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              hasAbove = true;
              break;
            }
          }
          if (hasAbove) {
            hasHeaderIncluded = true;
            selectRangeStr = `${formatAddress({ row: parsed.startRow - 1, col: parsed.startCol })}:${formatAddress({ row: parsed.endRow, col: parsed.endCol })}`;
          }
        }
      }

      // Cursor moves to first cell and drags to last cell; selection grows as it moves
      const parts = selectRangeStr.split(':');
      const start = parts[0];
      const end = parts[1] || start;

      await this.moveToTarget(`replica-cell-${start}`);
      this.selectRange(start);
      if (end !== start) {
        await this.moveToTarget(`replica-cell-${end}`);
        this.selectRange(selectRangeStr);
      }

      // Data tab -> Remove Duplicates button
      this.replicaActiveTab.set('data');
      await this.moveToTarget('tab-data');
      await this.moveToTarget('data-remove-duplicates-btn');

      // Dialog opens
      this.rdHeaders.set(false);
      this.replicaDialog.set('removeDuplicates');
      await this.delay(POST_OPEN_PAUSE_MS);

      // The cursor moves to "Data has header row" checkbox and ticks it ONLY if header row was included
      if (hasHeaderIncluded) {
        await this.moveToTarget('rd-headers-checkbox');
        this.rdHeaders.set(true);
        await this.delay(PRE_CLICK_PAUSE_MS);
      }

      // Clicks Remove duplicates
      await this.moveToTarget('rd-dialog-remove-btn');
      this.replicaDialog.set(null);

      // Result: compute real duplicate rows from replica sheet data
      if (parsed) {
        const seenRows = new Set<string>();
        const duplicates = new Set<number>();
        for (let r = parsed.startRow; r <= parsed.endRow; r++) {
          const rowParts: string[] = [];
          let isEmpty = true;
          for (let c = parsed.startCol; c <= parsed.endCol; c++) {
            const addr = formatAddress({ row: r, col: c });
            const cell = this.replicaCells().get(addr);
            const val = cell?.value ?? '';
            if (val !== '') isEmpty = false;
            rowParts.push(String(val).trim().toLowerCase());
          }
          if (isEmpty) continue;
          const key = rowParts.join('||');
          if (seenRows.has(key)) {
            duplicates.add(r + 1); // 1-based row number
          } else {
            seenRows.add(key);
          }
        }
        this.hiddenRows.set(duplicates);
      }
      return;
    }

    // 12. action
    if (type === 'action') {
      const act = g.action;

      // selectRange / goToCell: keep current behaviour (cursor drags from start cell to end cell)
      if (act === 'selectRange' || act === 'goToCell') {
        const range = g.range || g.cell || 'B2:B4';
        const parts = range.split(':');
        const start = parts[0];
        const end = parts[1] || start;

        await this.moveToTarget(`replica-cell-${start}`);
        this.selectRange(start);
        if (end !== start) {
          await this.moveToTarget(`replica-cell-${end}`);
          this.selectRange(range);
        }
        return;
      }

      const getFirstNonEmptyCell = (): string => {
        for (let r = 0; r < 20; r++) {
          for (let c = 0; c < 10; c++) {
            const addr = formatAddress({ row: r, col: c });
            const cell = this.replicaCells().get(addr);
            if (cell && cell.value !== '' && cell.value !== undefined && cell.value !== null) {
              return addr;
            }
          }
        }
        return 'A1';
      };

      // copy (Ctrl+C)
      if (act === 'copy') {
        let range = g.range || g.cell;
        if (!range) {
          range = getFirstNonEmptyCell();
        }

        const parts = range.split(':');
        const start = parts[0];
        const end = parts[1] || start;

        await this.moveToTarget(`replica-cell-${start}`);
        this.selectRange(start);
        if (end !== start) {
          await this.moveToTarget(`replica-cell-${end}`);
          this.selectRange(range);
        }

        await this.showKeyBadge('Ctrl+C', 1000);
        this.replicaCopiedRange.set(range);

        const copiedMap = new Map<string, { value: any; display: string }>();
        const parsed = parseRange(range);
        if (parsed) {
          for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
              const a = formatAddress({ row: r, col: c });
              const cell = this.replicaCells().get(a);
              copiedMap.set(a, { value: cell?.value ?? '', display: cell?.display ?? '' });
            }
          }
        }
        this.demoClipboard = { range, isCut: false, cells: copiedMap };
        return;
      }

      // cut (Ctrl+X)
      if (act === 'cut') {
        let range = g.range || g.cell;
        if (!range) {
          range = getFirstNonEmptyCell();
        }

        const parts = range.split(':');
        const start = parts[0];
        const end = parts[1] || start;

        await this.moveToTarget(`replica-cell-${start}`);
        this.selectRange(start);
        if (end !== start) {
          await this.moveToTarget(`replica-cell-${end}`);
          this.selectRange(range);
        }

        await this.showKeyBadge('Ctrl+X', 1000);
        this.replicaCopiedRange.set(range);

        const cutMap = new Map<string, { value: any; display: string }>();
        const parsed = parseRange(range);
        if (parsed) {
          for (let r = parsed.startRow; r <= parsed.endRow; r++) {
            for (let c = parsed.startCol; c <= parsed.endCol; c++) {
              const a = formatAddress({ row: r, col: c });
              const cell = this.replicaCells().get(a);
              cutMap.set(a, { value: cell?.value ?? '', display: cell?.display ?? '' });
            }
          }
        }
        this.demoClipboard = { range, isCut: true, cells: cutMap };

        const currentIndex = (this.goals || []).indexOf(g);
        const hasLaterPaste = (this.goals || []).slice(currentIndex + 1).some(
          (goal: any) => goal?.type === 'action' && goal?.action === 'paste'
        );

        if (!hasLaterPaste) {
          this.demoUndoStack.push(this.captureDemoSnapshot());
          this.demoRedoStack = [];
          const curCells = new Map(this.replicaCells());
          if (parsed) {
            for (let r = parsed.startRow; r <= parsed.endRow; r++) {
              for (let c = parsed.startCol; c <= parsed.endCol; c++) {
                const a = formatAddress({ row: r, col: c });
                curCells.set(a, { value: '', display: '' });
              }
            }
          }
          this.replicaCells.set(curCells);
          this.replicaCopiedRange.set(null);
          this.evaluateAndFormatCells();
        }
        return;
      }

      // paste (Ctrl+V)
      if (act === 'paste') {
        if (!this.demoClipboard) {
          const actCell = this.replicaActiveCell() || getFirstNonEmptyCell();
          await this.moveToTarget(`replica-cell-${actCell}`);
          this.selectRange(actCell);
          await this.showKeyBadge('Ctrl+C', 600);
          this.replicaCopiedRange.set(actCell);
          const cellData = this.replicaCells().get(actCell);
          const cMap = new Map<string, { value: any; display: string }>();
          cMap.set(actCell, { value: cellData?.value ?? '', display: cellData?.display ?? '' });
          this.demoClipboard = { range: actCell, isCut: false, cells: cMap };
        }

        let targetRange = g.range || g.cell;
        if (!targetRange) {
          const cp = parseRange(this.demoClipboard.range);
          if (cp) {
            const targetCol = Math.min(this.columns.length - 1, cp.endCol + 1);
            targetRange = formatAddress({ row: cp.startRow, col: targetCol });
          } else {
            targetRange = 'C2';
          }
        }

        const tParts = targetRange.split(':');
        const tStart = tParts[0];
        const tEnd = tParts[1] || tStart;

        await this.moveToTarget(`replica-cell-${tStart}`);
        this.selectRange(tStart);
        if (tEnd !== tStart) {
          await this.moveToTarget(`replica-cell-${tEnd}`);
          this.selectRange(targetRange);
        }

        await this.showKeyBadge('Ctrl+V', 1000);

        this.demoUndoStack.push(this.captureDemoSnapshot());
        this.demoRedoStack = [];

        const srcP = parseRange(this.demoClipboard.range);
        const tgtP = parseRange(targetRange);
        if (srcP && tgtP) {
          const dRow = tgtP.startRow - srcP.startRow;
          const dCol = tgtP.startCol - srcP.startCol;
          const curCells = new Map(this.replicaCells());

          for (let r = srcP.startRow; r <= srcP.endRow; r++) {
            for (let c = srcP.startCol; c <= srcP.endCol; c++) {
              const srcAddr = formatAddress({ row: r, col: c });
              const tgtAddr = formatAddress({ row: r + dRow, col: c + dCol });
              const srcVal = this.demoClipboard.cells.get(srcAddr);
              if (srcVal) {
                curCells.set(tgtAddr, { value: srcVal.value, display: String(srcVal.value) });
              }
              if (this.demoClipboard.isCut) {
                curCells.set(srcAddr, { value: '', display: '' });
              }
            }
          }
          this.replicaCells.set(curCells);
        }

        if (this.demoClipboard.isCut) {
          this.replicaCopiedRange.set(null);
          this.demoClipboard = null;
        }
        this.evaluateAndFormatCells();
        return;
      }

      // undo (Ctrl+Z)
      if (act === 'undo') {
        if (this.demoUndoStack.length === 0) {
          let emptyAddr = '';
          for (let r = 1; r < 8; r++) {
            for (let c = 1; c < 5; c++) {
              const a = formatAddress({ row: r, col: c });
              const cell = this.replicaCells().get(a);
              if (!cell || cell.value === '' || cell.value === undefined) {
                emptyAddr = a;
                break;
              }
            }
            if (emptyAddr) break;
          }
          if (!emptyAddr) emptyAddr = 'C2';
          await this.moveToTarget(`replica-cell-${emptyAddr}`);
          this.selectRange(emptyAddr);
          this.demoUndoStack.push(this.captureDemoSnapshot());
          const curCells = new Map(this.replicaCells());
          curCells.set(emptyAddr, { value: '100', display: '100' });
          this.replicaCells.set(curCells);
          this.evaluateAndFormatCells();
          await this.delay(600);
        }

        await this.moveToTarget('qa-undo-btn');
        await this.showKeyBadge('Ctrl+Z', 1000);

        const prev = this.demoUndoStack.pop()!;
        this.demoRedoStack.push(this.captureDemoSnapshot());
        this.restoreDemoSnapshot(prev);
        return;
      }

      // redo (Ctrl+Y)
      if (act === 'redo') {
        if (this.demoRedoStack.length === 0) {
          if (this.demoUndoStack.length === 0) {
            let emptyAddr = '';
            for (let r = 1; r < 8; r++) {
              for (let c = 1; c < 5; c++) {
                const a = formatAddress({ row: r, col: c });
                const cell = this.replicaCells().get(a);
                if (!cell || cell.value === '' || cell.value === undefined) {
                  emptyAddr = a;
                  break;
                }
              }
              if (emptyAddr) break;
            }
            if (!emptyAddr) emptyAddr = 'C2';
            await this.moveToTarget(`replica-cell-${emptyAddr}`);
            this.selectRange(emptyAddr);
            this.demoUndoStack.push(this.captureDemoSnapshot());
            const curCells = new Map(this.replicaCells());
            curCells.set(emptyAddr, { value: '50', display: '50' });
            this.replicaCells.set(curCells);
            this.evaluateAndFormatCells();
            await this.delay(400);
          }
          await this.moveToTarget('qa-undo-btn');
          await this.showKeyBadge('Ctrl+Z', 600);
          const prev = this.demoUndoStack.pop()!;
          this.demoRedoStack.push(this.captureDemoSnapshot());
          this.restoreDemoSnapshot(prev);
          await this.delay(400);
        }

        await this.moveToTarget('qa-redo-btn');
        await this.showKeyBadge('Ctrl+Y', 1000);

        const next = this.demoRedoStack.pop()!;
        this.demoUndoStack.push(this.captureDemoSnapshot());
        this.restoreDemoSnapshot(next);
        return;
      }

      return;
    }
  }

  getCellFill(cellAddr: string): string {
    const f = this.replicaFormats().get(cellAddr);
    if (!f?.fillColor) return '';
    return (COLORS as Record<string, string>)[f.fillColor] || f.fillColor;
  }

  getConditionLabel(c: string): string {
    switch (c) {
      case 'greaterThan': return 'Greater than';
      case 'lessThan': return 'Less than';
      case 'equals': return 'Equal to';
      case 'between': return 'Between';
      case 'contains': return 'Text contains';
      case 'isEmpty': return 'Is empty';
      default: return c || 'Select condition...';
    }
  }

  private evaluateDemoCondition(val: any, condition: string, targetVal?: any, targetVal2?: any): boolean {
    const d1 = parseDateParts(val);
    const d2 = parseDateParts(targetVal);
    if (d1 && d2) {
      if (condition === 'equals') return d1.utcTime === d2.utcTime;
      if (condition === 'greaterThan') return d1.utcTime > d2.utcTime;
      if (condition === 'lessThan') return d1.utcTime < d2.utcTime;
      if (condition === 'between') {
        const d3 = parseDateParts(targetVal2);
        const tMin = Math.min(d2.utcTime, d3 ? d3.utcTime : d2.utcTime);
        const tMax = Math.max(d2.utcTime, d3 ? d3.utcTime : d2.utcTime);
        return d1.utcTime >= tMin && d1.utcTime <= tMax;
      }
    }

    const str = String(val ?? '').trim().toLowerCase();
    const num = Number(val);

    switch (condition) {
      case 'equals':
        if (!isNaN(num) && targetVal !== undefined && !isNaN(Number(targetVal))) {
          return Math.abs(num - Number(targetVal)) < 0.0001;
        }
        return str === String(targetVal ?? '').trim().toLowerCase();

      case 'greaterThan':
        return !isNaN(num) && targetVal !== undefined && num > Number(targetVal);

      case 'lessThan':
        return !isNaN(num) && targetVal !== undefined && num < Number(targetVal);

      case 'between':
        return (
          !isNaN(num) &&
          targetVal !== undefined &&
          targetVal2 !== undefined &&
          num >= Math.min(Number(targetVal), Number(targetVal2)) &&
          num <= Math.max(Number(targetVal), Number(targetVal2))
        );

      case 'contains':
        return str.includes(String(targetVal ?? '').trim().toLowerCase());

      case 'isEmpty':
        return val === '' || val === null || val === undefined;

      default:
        return true;
    }
  }

  private computePivotTableData(
    range: string,
    goalRowsField?: string,
    goalValuesField?: string,
    goalSummarize?: string
  ): {
    headerValue: string;
    rows: Array<{ label: string; value: string | number }>;
    grandTotal: string | number;
    resolvedRowsField: string;
    resolvedValuesField: string;
    resolvedSummarize: string;
  } {
    const parsed = parseRange(range) || { startRow: 0, startCol: 0, endRow: 5, endCol: 2 };
    const headerRow = parsed.startRow;
    const colHeaders: Array<{ col: number; name: string }> = [];
    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      const a = formatAddress({ row: headerRow, col: c });
      const raw = this.replicaCells().get(a)?.value;
      const name = raw !== undefined && raw !== null && String(raw).trim() !== ''
        ? String(raw).trim()
        : indexToColLetter(c);
      colHeaders.push({ col: c, name });
    }

    let rowsCol = colHeaders[0]?.col ?? parsed.startCol;
    let resolvedRowsField = colHeaders[0]?.name ?? 'Row';
    if (goalRowsField) {
      const match = colHeaders.find((h) => h.name.toLowerCase() === goalRowsField.trim().toLowerCase());
      if (match) {
        rowsCol = match.col;
        resolvedRowsField = match.name;
      } else {
        resolvedRowsField = goalRowsField;
      }
    }

    let valuesCol = -1;
    let resolvedValuesField = '';
    if (goalValuesField) {
      const match = colHeaders.find((h) => h.name.toLowerCase() === goalValuesField.trim().toLowerCase());
      if (match) {
        valuesCol = match.col;
        resolvedValuesField = match.name;
      } else {
        resolvedValuesField = goalValuesField;
      }
    }

    if (valuesCol === -1) {
      for (const h of colHeaders) {
        let hasNum = false;
        for (let r = headerRow + 1; r <= parsed.endRow; r++) {
          const a = formatAddress({ row: r, col: h.col });
          const val = this.replicaCells().get(a)?.value;
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            const num = Number(val);
            if (!isNaN(num)) {
              hasNum = true;
              break;
            }
          }
        }
        if (hasNum) {
          valuesCol = h.col;
          resolvedValuesField = h.name;
          break;
        }
      }
    }

    if (valuesCol === -1) {
      const fallbackH = colHeaders.find((h) => h.col !== rowsCol) || colHeaders[colHeaders.length - 1];
      valuesCol = fallbackH?.col ?? parsed.endCol;
      resolvedValuesField = fallbackH?.name ?? 'Values';
    }

    const resolvedSummarize = (goalSummarize || 'SUM').toUpperCase();

    interface GroupAccumulator {
      label: string;
      sum: number;
      count: number;
      numCount: number;
    }
    const groupMap = new Map<string, GroupAccumulator>();

    for (let r = headerRow + 1; r <= parsed.endRow; r++) {
      const rowAddr = formatAddress({ row: r, col: rowsCol });
      const rowRaw = this.replicaCells().get(rowAddr)?.value;
      const label = rowRaw !== undefined && rowRaw !== null && String(rowRaw).trim() !== ''
        ? String(rowRaw).trim()
        : '(blank)';

      const valAddr = formatAddress({ row: r, col: valuesCol });
      const valRaw = this.replicaCells().get(valAddr)?.value;
      const num = Number(valRaw);
      const isNum = !isNaN(num) && valRaw !== '' && valRaw !== null && valRaw !== undefined;

      let acc = groupMap.get(label);
      if (!acc) {
        acc = { label, sum: 0, count: 0, numCount: 0 };
        groupMap.set(label, acc);
      }
      acc.count++;
      if (isNum) {
        acc.sum += num;
        acc.numCount++;
      }
    }

    const rows: Array<{ label: string; value: string | number }> = [];
    let grandSum = 0;
    let grandCount = 0;
    let grandNumCount = 0;

    for (const acc of groupMap.values()) {
      grandSum += acc.sum;
      grandCount += acc.count;
      grandNumCount += acc.numCount;

      let computedVal: number | string = 0;
      if (resolvedSummarize === 'COUNT') {
        computedVal = acc.count;
      } else if (resolvedSummarize === 'AVERAGE') {
        const avg = acc.numCount > 0 ? acc.sum / acc.numCount : 0;
        computedVal = Number.isInteger(avg) ? avg : Math.round(avg * 100) / 100;
      } else {
        computedVal = Number.isInteger(acc.sum) ? acc.sum : Math.round(acc.sum * 100) / 100;
      }
      rows.push({ label: acc.label, value: computedVal });
    }

    let grandTotal: number | string = 0;
    if (resolvedSummarize === 'COUNT') {
      grandTotal = grandCount;
    } else if (resolvedSummarize === 'AVERAGE') {
      const avg = grandNumCount > 0 ? grandSum / grandNumCount : 0;
      grandTotal = Number.isInteger(avg) ? avg : Math.round(avg * 100) / 100;
    } else {
      grandTotal = Number.isInteger(grandSum) ? grandSum : Math.round(grandSum * 100) / 100;
    }

    let summarizePrefix = 'Sum';
    if (resolvedSummarize === 'COUNT') summarizePrefix = 'Count';
    else if (resolvedSummarize === 'AVERAGE') summarizePrefix = 'Average';
    else summarizePrefix = resolvedSummarize.charAt(0).toUpperCase() + resolvedSummarize.slice(1).toLowerCase();

    const headerValue = `${summarizePrefix} of ${resolvedValuesField}`;

    return {
      headerValue,
      rows,
      grandTotal,
      resolvedRowsField,
      resolvedValuesField,
      resolvedSummarize,
    };
  }

  getToastTitle(): string {
    const t = (this.title || '').trim();
    if (!t) return 'complete this task';
    return t.toLowerCase();
  }
}
