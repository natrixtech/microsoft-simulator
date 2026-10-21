import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  HostListener,
  OnDestroy,
  OnInit,
  ViewChild,
  computed,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  ActivityPayload,
  ConfettiService,
  LessonTask,
  PluginBridgeService,
  TaskCardComponent,
  TaskGoal,
  ThemePreference,
} from 'shared';
import {
  BorderAction,
  CellData,
  CellFormat,
  ChartItem,
  ChartType,
  NumberFormatType,
  PivotTableItem,
  SheetSnapshot,
  SortOrder,
  TextAlign,
} from './models/excel.types';
import { ExcelStateService } from './services/excel-state.service';
import { ExcelClipboardService } from './services/excel-clipboard.service';
import { ExcelHistoryService } from './services/excel-history.service';
import { ExcelActionLoggerService } from './services/excel-action-logger.service';
import { ExcelEvaluatorService } from './services/excel-evaluator.service';
import {
  FormulaEngineService,
  formatAddress,
  indexToColLetter,
  parseAddress,
  parseDateParts,
  parseRange,
} from './services/formula-engine.service';
import { ActiveRibbonTab, RibbonComponent } from './components/ribbon/ribbon.component';
import { FormulaBarComponent } from './components/formula-bar/formula-bar.component';
import { GridComponent } from './components/grid/grid.component';
import { SortDialogComponent } from './components/dialogs/sort-dialog.component';
import { FilterDropdownComponent } from './components/dialogs/filter-dropdown.component';
import { ConditionalFormatDialogComponent } from './components/dialogs/conditional-format-dialog.component';
import { DataValidationDialogComponent } from './components/dialogs/data-validation-dialog.component';
import { RemoveDuplicatesDialogComponent } from './components/dialogs/remove-duplicates-dialog.component';
import { ChartDialogComponent } from './components/dialogs/chart-dialog.component';
import { PivotTableDialogComponent } from './components/dialogs/pivot-table-dialog.component';
import { CommentDialogComponent } from './components/dialogs/comment-dialog.component';
import { FunctionPickerDialogComponent } from './components/dialogs/function-picker-dialog.component';
import { ExcelDemoComponent } from './components/demo/excel-demo.component';
import { EXCEL_LESSONS } from './excel-lessons.data';

@Component({
  selector: 'app-excel-editor',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    TaskCardComponent,
    RibbonComponent,
    FormulaBarComponent,
    GridComponent,
    SortDialogComponent,
    FilterDropdownComponent,
    ConditionalFormatDialogComponent,
    DataValidationDialogComponent,
    RemoveDuplicatesDialogComponent,
    ChartDialogComponent,
    PivotTableDialogComponent,
    CommentDialogComponent,
    FunctionPickerDialogComponent,
    ExcelDemoComponent,
  ],
  templateUrl: './excel-editor.component.html',
  styleUrls: ['./excel-editor.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ExcelEditorComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('editorContainer') containerRef?: ElementRef<HTMLElement>;
  @ViewChild('workLayout') workLayoutRef?: ElementRef<HTMLElement>;
  @ViewChild('movableGroup') movableGroupRef?: ElementRef<HTMLElement>;
  @ViewChild(GridComponent) gridComponent?: GridComponent;
  @ViewChild(FormulaBarComponent) formulaBarComponent?: FormulaBarComponent;

  commitPendingEdit(): void {
    const gridInput = this.gridComponent?.cellEditorInputRef?.nativeElement;
    if (gridInput && this.state.isEditing()) {
      this.state.editInput.set(gridInput.value);
    }
    const formulaInput = this.formulaBarComponent?.formulaInputRef?.nativeElement;
    if (formulaInput && document.activeElement === formulaInput) {
      this.state.editInput.set(formulaInput.value);
      this.state.editingAddress.set(this.state.activeCell());
      this.state.isEditing.set(true);
    }

    if (!this.state.isEditing()) return;

    const addr = this.state.editingAddress() || this.state.activeCell();
    const val = this.state.editInput();

    // Data validation reject check
    const dv = this.state.getDataValidation(addr);
    if (dv && dv.rejectInvalid && val !== '') {
      const isAllowed = dv.allowed.some((a) => a.trim().toLowerCase() === val.trim().toLowerCase());
      if (!isAllowed) {
        alert(
          `The value you entered doesn't match the data validation restrictions defined for this cell. Allowed: ${dv.allowed.join(', ')}`
        );
        this.state.isEditing.set(false);
        this.cdr.detectChanges();
        return;
      }
    }

    this.history.push(this.state.getSnapshot());
    this.state.setCellValue(addr, val);
    this.state.isEditing.set(false);
    this.cdr.detectChanges();
  }

  focusGrid(): void {
    this.gridComponent?.focus();
    setTimeout(() => {
      this.gridComponent?.focus();
    }, 0);
  }

  // Activity & Task state
  readonly currentActivity = signal<ActivityPayload | null>(null);
  readonly taskIndex = signal<number>(0);
  readonly passedTasks = signal<Set<number>>(new Set());
  readonly isPassed = signal<boolean>(false);
  readonly isCompleted = signal<boolean>(false);
  readonly feedbackMsg = signal<string>('');
  readonly isDemoOpen = signal<boolean>(false);

  // Initial Task Snapshot for Demo & reset
  readonly currentTaskSnapshot = signal<SheetSnapshot | null>(null);

  // Task Card positioning state (Top-right of worksheet viewport)
  readonly groupX = signal<number | null>(null);
  readonly groupY = signal<number | null>(null);
  readonly isDragging = signal<boolean>(false);
  readonly hasUserMoved = signal<boolean>(false);
  private dragStartPointerX = 0;
  private dragStartPointerY = 0;
  private dragStartGroupX = 0;
  private dragStartGroupY = 0;
  private resizeObserver?: ResizeObserver;

  // Active Ribbon Tab
  readonly activeTab = signal<ActiveRibbonTab>('home');

  // Active Dialogs
  readonly showSortDialog = signal<boolean>(false);
  readonly showFilterDropdown = signal<boolean>(false);
  readonly filterDropdownParams = signal<{ colIndex: number; anchorX: number; anchorY: number } | null>(null);
  readonly showConditionalFormatDialog = signal<boolean>(false);
  readonly showDataValidationDialog = signal<boolean>(false);
  readonly showRemoveDuplicatesDialog = signal<boolean>(false);
  readonly showChartDialog = signal<boolean>(false);
  readonly chartDialogInitialType = signal<ChartType>('column');
  readonly showPivotTableDialog = signal<boolean>(false);
  readonly showCommentDialog = signal<boolean>(false);
  readonly commentTargetCell = signal<string>('A1');
  readonly showFunctionPicker = signal<boolean>(false);

  // Subscriptions & Timers
  private bridgeUnsubscribes: Array<() => void> = [];
  private advanceTimer: any = null;

  constructor(
    public bridge: PluginBridgeService,
    public state: ExcelStateService,
    public clipboard: ExcelClipboardService,
    public history: ExcelHistoryService,
    private actionLogger: ExcelActionLoggerService,
    private evaluator: ExcelEvaluatorService,
    private formulaEngine: FormulaEngineService,
    private confetti: ConfettiService,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    const act = this.bridge.activity();
    if (act && act.app === 'excel') {
      this.loadActivity(act);
    } else {
      // Fallback default lesson if none provided yet
      this.loadActivity(EXCEL_LESSONS['cellValue']);
    }

    this.bridgeUnsubscribes.push(
      this.bridge.onActivity((newAct) => {
        if (newAct && newAct.app === 'excel') {
          this.loadActivity(newAct);
          this.cdr.markForCheck();
        }
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onForceFinish(() => {
        this.handleForceFinish();
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onThemeChange(() => {
        this.cdr.detectChanges();
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onPauseChange(() => {
        this.cdr.detectChanges();
      })
    );
  }

  ngAfterViewInit(): void {
    requestAnimationFrame(() => {
      this.computeInitialPlacement();
    });

    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => {
        this.handleViewportOrGroupResize();
      });
      if (this.workLayoutRef?.nativeElement) {
        this.resizeObserver.observe(this.workLayoutRef.nativeElement);
      }
      if (this.movableGroupRef?.nativeElement) {
        this.resizeObserver.observe(this.movableGroupRef.nativeElement);
      }
    }
  }

  ngOnDestroy(): void {
    for (const unsub of this.bridgeUnsubscribes) {
      unsub();
    }
    this.bridgeUnsubscribes = [];
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = undefined;
    }
  }

  get isDarkTheme(): boolean {
    return this.bridge.theme() === 'DARK';
  }

  get isFrozen(): boolean {
    return this.bridge.isPaused() || this.isCompleted();
  }

  closeAllDialogs(): void {
    this.state.isEditing.set(false);
    this.showSortDialog.set(false);
    this.showFilterDropdown.set(false);
    this.filterDropdownParams.set(null);
    this.showConditionalFormatDialog.set(false);
    this.showDataValidationDialog.set(false);
    this.showRemoveDuplicatesDialog.set(false);
    this.showChartDialog.set(false);
    this.showPivotTableDialog.set(false);
    this.showCommentDialog.set(false);
    this.showFunctionPicker.set(false);
    this.focusGrid();
  }

  closeDialog(dialogName: string): void {
    switch (dialogName) {
      case 'sort': this.showSortDialog.set(false); break;
      case 'filter': this.showFilterDropdown.set(false); this.filterDropdownParams.set(null); break;
      case 'condFormat': this.showConditionalFormatDialog.set(false); break;
      case 'dataVal': this.showDataValidationDialog.set(false); break;
      case 'removeDup': this.showRemoveDuplicatesDialog.set(false); break;
      case 'chart': this.showChartDialog.set(false); break;
      case 'pivot': this.showPivotTableDialog.set(false); break;
      case 'comment': this.showCommentDialog.set(false); break;
      case 'fnPicker': this.showFunctionPicker.set(false); break;
    }
    this.focusGrid();
  }

  openConditionalFormatDialog(): void {
    this.commitPendingEdit();
    this.showConditionalFormatDialog.set(true);
  }

  openPivotTableDialog(): void {
    this.commitPendingEdit();
    this.showPivotTableDialog.set(true);
  }

  openChartDialog(type: ChartType): void {
    this.commitPendingEdit();
    this.chartDialogInitialType.set(type);
    this.showChartDialog.set(true);
  }

  openCommentDialog(): void {
    this.commitPendingEdit();
    this.commentTargetCell.set(this.state.activeCell());
    this.showCommentDialog.set(true);
  }

  openFunctionPickerDialog(): void {
    this.commitPendingEdit();
    this.showFunctionPicker.set(true);
  }

  openSortDialog(): void {
    this.commitPendingEdit();
    this.showSortDialog.set(true);
  }

  openDataValidationDialog(): void {
    this.commitPendingEdit();
    this.showDataValidationDialog.set(true);
  }

  openRemoveDuplicatesDialog(): void {
    this.commitPendingEdit();
    this.showRemoveDuplicatesDialog.set(true);
  }

  onOpenCommentPopup(cell: string): void {
    this.commitPendingEdit();
    this.commentTargetCell.set(cell);
    this.showCommentDialog.set(true);
  }

  onShowDemo(): void {
    this.commitPendingEdit();
    this.isDemoOpen.set(true);
  }

  onTabChange(tab: ActiveRibbonTab): void {
    this.commitPendingEdit();
    this.activeTab.set(tab);
  }

  get hasShowMeHow(): boolean {
    const smh = this.currentTask?.showMeHow ?? (this.currentActivity() as any)?.showMeHow;
    return smh === true || (typeof smh === 'object' && smh !== null);
  }

  get currentTask(): LessonTask | null {
    const act = this.currentActivity();
    if (!act?.tasks || act.tasks.length === 0) return null;
    return act.tasks[this.taskIndex()] || null;
  }

  get isAssessment(): boolean {
    return this.currentActivity()?.mode === 'assessment';
  }

  get totalTasks(): number {
    return this.currentActivity()?.tasks?.length || 1;
  }

  get currentCellFormat(): CellFormat {
    const addr = this.state.activeCell();
    return this.state.cellFormats().get(addr) || {};
  }

  get activeCellRaw(): string {
    if (this.state.isEditing()) {
      return this.state.editInput();
    }
    const cell = this.state.getCell(this.state.activeCell());
    if (!cell) return '';
    return cell.formula || (cell.raw !== undefined ? String(cell.raw) : '');
  }

  loadActivity(act: ActivityPayload): void {
    this.hasUserMoved.set(false);
    this.currentActivity.set(act);
    this.taskIndex.set(0);
    this.passedTasks.set(new Set());
    this.isPassed.set(false);
    this.isCompleted.set(false);
    this.closeAllDialogs();
    this.feedbackMsg.set('');
    this.isDemoOpen.set(false);

    if (act.tasks && act.tasks.length > 0) {
      this.setupTask(0, true);
    }

    requestAnimationFrame(() => {
      this.computeInitialPlacement();
    });
  }

  private setupTask(index: number, isFirst: boolean): void {
    const act = this.currentActivity();
    if (!act?.tasks) return;
    const task = act.tasks[index];
    if (!task) return;

    this.taskIndex.set(index);
    this.isPassed.set(false);
    this.feedbackMsg.set('');

    // Clear task-local undo history and action log
    this.history.clear();
    this.actionLogger.clear();

    if (task.startingSheet !== null && task.startingSheet !== undefined) {
      // Explicit starting sheet
      this.state.initFromStartingSheet(task.startingSheet);
    } else if (isFirst) {
      // First task must have a starting sheet
      this.state.initFromStartingSheet({ cells: {}, formats: {} });
    }
    // If startingSheet === null on later assessment tasks, preserve existing sheet state!

    // Capture task-start snapshot for demo / reset
    this.currentTaskSnapshot.set(this.state.getSnapshot());
    this.cdr.markForCheck();
  }

  // Task checking
  checkWork(): void {
    this.commitPendingEdit();
    if (this.bridge.isPaused() || this.isCompleted()) return;
    const task = this.currentTask;
    if (!task) return;

    const res = this.evaluator.evaluateGoals(task.goals || []);

    if (res.passed) {
      this.isPassed.set(true);
      this.feedbackMsg.set('');
      const newPassed = new Set(this.passedTasks()).add(this.taskIndex());
      this.passedTasks.set(newPassed);

      this.confetti.celebrate();

      const total = this.totalTasks;
      const isLast = this.taskIndex() === total - 1;

      if (this.isAssessment) {
        if (!isLast) {
          // Auto-advance in assessment after ~1s
          if (this.advanceTimer) clearTimeout(this.advanceTimer);
          this.advanceTimer = setTimeout(() => {
            this.setupTask(this.taskIndex() + 1, false);
          }, 1100);
        } else {
          // Completed all assessment tasks
          this.isCompleted.set(true);
          this.closeAllDialogs();
          const passedCount = newPassed.size;
          const allPassed = passedCount === total;
          this.bridge.sendComplete(allPassed, passedCount, total);
        }
      } else {
        // Lesson mode
        if (total > 1 && !isLast) {
          // Lesson mode with more than one task: show "Well done!" for ~1.5s, then move to next task automatically
          if (this.advanceTimer) clearTimeout(this.advanceTimer);
          this.advanceTimer = setTimeout(() => {
            this.setupTask(this.taskIndex() + 1, false);
          }, 1500);
        } else {
          // Single-task lesson or last task: complete immediately!
          this.isCompleted.set(true);
          this.closeAllDialogs();
          const passedCount = newPassed.size;
          const allPassed = passedCount === total;
          this.bridge.sendComplete(allPassed, passedCount, total);
        }
      }
    } else {
      this.isPassed.set(false);
      const hint = task.hint || 'Not quite right. Try again.';
      this.feedbackMsg.set(hint);
    }
    this.cdr.markForCheck();
  }

  onStartAgain(): void {
    this.commitPendingEdit();
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }
    this.closeAllDialogs();
    this.isDemoOpen.set(false);
    this.isPassed.set(false);
    this.isCompleted.set(false);
    this.feedbackMsg.set('');
    this.passedTasks.set(new Set());
    this.taskIndex.set(0);
    this.bridge.resetCompleted();

    // Clear selection and cursor
    this.state.isEditing.set(false);
    this.state.editInput.set('');
    this.state.activeCell.set('A1');
    this.state.selectedRange.set('A1');
    this.state.selectedChartId.set(null);
    window.getSelection()?.removeAllRanges();
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }

    // Restore starting sheet of task 1 (task index 0)
    const act = this.currentActivity();
    if (act?.tasks && act.tasks.length > 0) {
      this.setupTask(0, true);
    }
    this.cdr.markForCheck();
  }

  private handleForceFinish(): void {
    this.commitPendingEdit();
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }
    this.isDemoOpen.set(false);
    this.isCompleted.set(true);
    this.closeAllDialogs();
    const passedCount = this.passedTasks().size;
    const allPassed = passedCount === this.totalTasks;
    this.bridge.sendComplete(allPassed, passedCount, this.totalTasks);
    this.cdr.markForCheck();
  }

  // Cell reference click in instructions
  selectReference(ref: string): void {
    this.commitPendingEdit();
    const clean = ref.trim().toUpperCase();
    if (clean.includes(':')) {
      this.state.selectRange(clean);
      this.actionLogger.log('selectRange', clean);
    } else {
      this.state.selectCell(clean);
      this.actionLogger.log('goToCell', clean);
    }
  }

  // Formula bar and cell editing
  onFormulaValueChange(val: string): void {
    if (this.state.isEditing()) {
      this.state.editInput.set(val);
    } else {
      this.state.editingAddress.set(this.state.activeCell());
      this.state.editInput.set(val);
      this.state.isEditing.set(true);
    }
  }

  onFormulaCommit(payload: string | { value: string; dRow?: number; dCol?: number }): void {
    const val = typeof payload === 'string' ? payload : payload.value;
    const dRow = typeof payload === 'string' ? 1 : (payload.dRow ?? 1);
    const dCol = typeof payload === 'string' ? 0 : (payload.dCol ?? 0);

    const addr = this.state.activeCell();
    this.history.push(this.state.getSnapshot());
    this.state.setCellValue(addr, val);
    this.state.isEditing.set(false);

    if (dRow !== 0 || dCol !== 0) {
      const p = parseAddress(addr);
      if (p) {
        const r = Math.max(0, Math.min(this.state.maxRows - 1, p.row + dRow));
        const c = Math.max(0, Math.min(this.state.maxCols - 1, p.col + dCol));
        this.state.selectCell(formatAddress({ row: r, col: c }));
      }
    }

    this.focusGrid();
    this.gridComponent?.scrollToActiveCell();
  }

  onFormulaCancel(): void {
    this.state.isEditing.set(false);
    this.focusGrid();
  }

  onFormulaNavigate(addr: string): void {
    this.selectReference(addr);
  }

  // Ribbon Actions
  onUndo(): void {
    this.commitPendingEdit();
    const prev = this.history.undo(this.state.getSnapshot());
    if (prev) {
      this.state.restoreSnapshot(prev);
      this.cdr.markForCheck();
    }
    this.focusGrid();
  }

  onRedo(): void {
    this.commitPendingEdit();
    const next = this.history.redo(this.state.getSnapshot());
    if (next) {
      this.state.restoreSnapshot(next);
      this.cdr.markForCheck();
    }
    this.focusGrid();
  }

  onFormatChange(updates: Partial<CellFormat>): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.formatSelection(updates);
    this.focusGrid();
  }

  onApplyBorder(action: BorderAction): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.applyBorderToSelection(action);
    this.focusGrid();
  }

  onCut(): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.clipboard.cut(this.state.selectedRange());
    this.focusGrid();
  }

  onCopy(): void {
    this.commitPendingEdit();
    this.clipboard.copy(this.state.selectedRange());
    this.focusGrid();
  }

  onPaste(): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.clipboard.paste(this.state.activeCell());
    this.focusGrid();
  }

  private isNumericCell(cell: CellData | undefined): boolean {
    if (!cell) return false;
    return this.isNumericValue(cell.value);
  }

  private isNumericValue(val: any): boolean {
    if (val === undefined || val === null || val === '') return false;
    if (typeof val === 'number' && !isNaN(val)) return true;
    if (typeof val === 'string') {
      const trimmed = val.trim();
      if (trimmed === '') return false;
      if (parseDateParts(trimmed)) return false;
      if (['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(trimmed)) return false;
      return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(trimmed) && !isNaN(Number(trimmed));
    }
    return false;
  }

  private detectHeader(region: { startRow: number; endRow: number; startCol: number; endCol: number }): boolean {
    if (region.startRow >= region.endRow) return false;

    for (let c = region.startCol; c <= region.endCol; c++) {
      const topAddr = formatAddress({ row: region.startRow, col: c });
      const topCell = this.state.getCell(topAddr);
      const topVal = topCell?.value;
      const topIsText = topVal !== undefined && topVal !== null && String(topVal).trim() !== '' && !this.isNumericValue(topVal);

      if (topIsText) {
        for (let r = region.startRow + 1; r <= region.endRow; r++) {
          const belowAddr = formatAddress({ row: r, col: c });
          const belowCell = this.state.getCell(belowAddr);
          if (belowCell && this.isNumericValue(belowCell.value)) {
            return true;
          }
        }
      }
    }

    const topFmt = this.state.cellFormats().get(formatAddress({ row: region.startRow, col: region.startCol }));
    if (topFmt?.bold) {
      return true;
    }

    return false;
  }

  onAutoSum(fnName: any = 'SUM'): void {
    const fn = (typeof fnName === 'string' && fnName.trim() ? fnName.trim() : 'SUM').toUpperCase();
    this.commitPendingEdit();
    const active = this.state.activeCell();
    const p = parseAddress(active);
    if (!p) return;

    let targetRow = p.row;
    let targetCol = p.col;

    const activeCellData = this.state.getCell(active);
    const activeHasValue =
      !!activeCellData &&
      activeCellData.value !== '' &&
      activeCellData.value !== null &&
      activeCellData.value !== undefined &&
      String(activeCellData.raw || '').trim() !== '';

    if (activeHasValue) {
      // Active cell is not empty: put SUM in first empty cell below the selected range
      const sel = this.state.selectionCoords();
      targetCol = p.col;
      let r = sel.endRow + 1;
      while (r < this.state.maxRows) {
        const addr = formatAddress({ row: r, col: targetCol });
        const cell = this.state.getCell(addr);
        const hasVal =
          !!cell &&
          cell.value !== '' &&
          cell.value !== null &&
          cell.value !== undefined &&
          String(cell.raw || '').trim() !== '';
        if (!hasVal) {
          break;
        }
        r++;
      }
      targetRow = Math.min(this.state.maxRows - 1, r);
    }

    // Now find contiguous numbers directly above target cell
    let aboveStartRow = -1;
    let aboveEndRow = -1;
    let r = targetRow - 1;

    while (r >= 0) {
      const addr = formatAddress({ row: r, col: targetCol });
      const cell = this.state.getCell(addr);
      if (this.isNumericCell(cell)) {
        if (aboveEndRow === -1) aboveEndRow = r;
        aboveStartRow = r;
        r--;
      } else {
        break; // Stop at first text or empty cell
      }
    }

    let formula = '';
    if (aboveStartRow !== -1 && aboveEndRow !== -1) {
      const startAddr = formatAddress({ row: aboveStartRow, col: targetCol });
      const endAddr = formatAddress({ row: aboveEndRow, col: targetCol });
      formula = `=${fn}(${startAddr}:${endAddr})`;
    } else {
      // No numbers above, check to the left
      let leftStartCol = -1;
      let leftEndCol = -1;
      let c = targetCol - 1;

      while (c >= 0) {
        const addr = formatAddress({ row: targetRow, col: c });
        const cell = this.state.getCell(addr);
        if (this.isNumericCell(cell)) {
          if (leftEndCol === -1) leftEndCol = c;
          leftStartCol = c;
          c--;
        } else {
          break; // Stop at first text or empty cell
        }
      }

      if (leftStartCol !== -1 && leftEndCol !== -1) {
        const startAddr = formatAddress({ row: targetRow, col: leftStartCol });
        const endAddr = formatAddress({ row: targetRow, col: leftEndCol });
        formula = `=${fn}(${startAddr}:${endAddr})`;
      } else {
        const fallbackRow = Math.max(0, targetRow - 1);
        const fallbackAddr = formatAddress({ row: fallbackRow, col: targetCol });
        formula = `=${fn}(${fallbackAddr}:${fallbackAddr})`;
      }
    }

    const targetAddr = formatAddress({ row: targetRow, col: targetCol });
    this.history.push(this.state.getSnapshot());
    this.state.setCellValue(targetAddr, formula);
    this.state.selectCell(targetAddr);
    this.focusGrid();
  }

  onQuickSort(order: SortOrder): void {
    this.commitPendingEdit();
    const sel = this.state.selectedRange();
    const p = parseAddress(this.state.activeCell()) || { row: 0, col: 0 };
    const parsedSel = parseRange(sel);
    const isSingleCell =
      !parsedSel || (parsedSel.startRow === parsedSel.endRow && parsedSel.startCol === parsedSel.endCol);

    let sortRangeStr = sel;
    let hasHeader = true;

    if (isSingleCell) {
      const activeAddr = formatAddress(p);
      sortRangeStr = this.state.getCurrentRegion(activeAddr);
      const parsedRegion = parseRange(sortRangeStr);
      if (parsedRegion) {
        hasHeader = this.detectHeader(parsedRegion);
      }
    } else {
      // With a range selected, sort that range as now
      hasHeader = true;
    }

    const byCol = indexToColLetter(p.col);
    this.history.push(this.state.getSnapshot());
    this.state.sortRange(sortRangeStr, byCol, order, hasHeader);
    this.focusGrid();
  }

  onApplySort(data: { byColumn: string; order: SortOrder; hasHeader: boolean }): void {
    this.commitPendingEdit();
    let rangeStr = this.state.selectedRange();
    const parsed = parseRange(rangeStr);
    if (!parsed || (parsed.startRow === parsed.endRow && parsed.startCol === parsed.endCol)) {
      rangeStr = this.state.getCurrentRegion(this.state.activeCell());
    }
    this.history.push(this.state.getSnapshot());
    this.state.sortRange(rangeStr, data.byColumn, data.order, data.hasHeader);
    this.closeDialog('sort');
  }

  onToggleFilter(): void {
    this.commitPendingEdit();
    const curFilter = this.state.filter();
    if (curFilter) {
      this.state.setFilter(null);
    } else {
      const sel = this.state.selectionCoords();
      const newFilter = {
        range: this.state.selectedRange(),
        headerRow: sel.startRow,
        startRow: sel.startRow,
        endRow: sel.endRow,
        startCol: sel.startCol,
        endCol: sel.endCol,
        columnFilters: {},
      };
      this.state.setFilter(newFilter);
    }
    this.focusGrid();
  }

  onOpenFilterDropdown(params: { colIndex: number; anchorX: number; anchorY: number }): void {
    this.commitPendingEdit();
    this.filterDropdownParams.set(params);
    this.showFilterDropdown.set(true);
  }

  onApplyColumnFilter(data: { colIndex: number; condition?: any; selectedValues?: string[] }): void {
    this.commitPendingEdit();
    const f = this.state.filter();
    if (!f) return;
    const updated = { ...f, columnFilters: { ...f.columnFilters } };
    updated.columnFilters[data.colIndex] = {
      colIndex: data.colIndex,
      condition: data.condition,
      selectedValues: data.selectedValues,
    };
    this.state.setFilter(updated);
    this.closeDialog('filter');
  }

  onSortFromFilter(data: { colIndex: number; order: SortOrder }): void {
    this.commitPendingEdit();
    const f = this.state.filter();
    if (!f) return;
    const colLetter = indexToColLetter(data.colIndex);
    this.history.push(this.state.getSnapshot());
    this.state.sortRange(f.range, colLetter, data.order, true);
    this.closeDialog('filter');
  }

  onAddChart(chart: ChartItem): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());

    const parsed = parseRange(chart.range);
    let targetX = 180;
    let targetY = 120;
    if (parsed) {
      let xPos = 44; // Row header width
      for (let c = 0; c <= parsed.endCol; c++) {
        const letter = indexToColLetter(c);
        xPos += this.state.columnWidths().get(letter) || 84;
      }
      targetX = xPos + 20;
      targetY = Math.max(30, parsed.startRow * 24 + 10);
    }

    const existingCharts = this.state.charts();
    while (existingCharts.some((ec) => Math.abs(ec.x - targetX) < 25 && Math.abs(ec.y - targetY) < 25)) {
      targetX += 30;
      targetY += 30;
    }

    chart.x = targetX;
    chart.y = targetY;

    this.state.addChart(chart);
    this.closeDialog('chart');
  }

  onAddPivotTable(pt: PivotTableItem): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.addPivotTable(pt);
    this.closeDialog('pivot');
  }

  onSaveComment(data: { cell: string; text: string }): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.setComment(data.cell, data.text);
    this.closeDialog('comment');
  }

  onDeleteComment(cell: string): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.deleteComment(cell);
    this.closeDialog('comment');
  }

  onAddConditionalFormat(rule: any): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.addConditionalFormat(rule);
    this.closeDialog('condFormat');
  }

  onAddDataValidation(rule: any): void {
    this.commitPendingEdit();
    this.history.push(this.state.getSnapshot());
    this.state.addDataValidation(rule);
    this.closeDialog('dataVal');
  }

  onInsertNamedFunction(fnName: string): void {
    this.commitPendingEdit();
    const active = this.state.activeCell();
    this.history.push(this.state.getSnapshot());
    this.state.setCellValue(active, `=${fnName}()`);
    this.closeDialog('fnPicker');
  }

  // Initial placement of Task Card + Mascot group at top-right of worksheet viewport
  computeInitialPlacement(): void {
    if (!this.workLayoutRef?.nativeElement) return;
    const layoutEl = this.workLayoutRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    if (layoutRect.width === 0 || layoutRect.height === 0) return;

    const groupEl = this.movableGroupRef?.nativeElement;
    const groupRect = groupEl?.getBoundingClientRect();
    const groupWidth = groupRect?.width || groupEl?.offsetWidth || 300;
    const groupHeight = groupRect?.height || groupEl?.offsetHeight || 480;

    // Report if worksheet viewport is physically too short to fit the complete group
    if (layoutRect.height > 0 && layoutRect.height < groupHeight + 16) {
      console.warn(
        `Excel worksheet viewport height (${Math.round(layoutRect.height)}px) is too short to fit the task-card and mascot group (${Math.round(groupHeight)}px).`
      );
    }

    // Top-right placement: ~8-12px spacing from top and right of the worksheet viewport
    const targetY = 10;
    const targetX = layoutRect.width - groupWidth - 12;

    this.groupX.set(Math.round(this.clampX(targetX)));
    this.groupY.set(Math.round(this.clampY(targetY)));
    this.cdr.markForCheck();
  }

  clampX(x: number): number {
    if (!this.workLayoutRef?.nativeElement) return Math.max(8, x);
    const layoutEl = this.workLayoutRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupRect = groupEl?.getBoundingClientRect();
    const groupWidth = groupRect?.width || groupEl?.offsetWidth || 300;

    // Left limit accounts for assessment sidebar if present
    const sidebarEl = layoutEl.querySelector('.assessment-progress-sidebar') as HTMLElement | null;
    const minX = sidebarEl ? sidebarEl.offsetWidth + 8 : 8;
    const maxX = Math.max(minX, layoutRect.width - groupWidth - 10);
    return Math.min(Math.max(x, minX), maxX);
  }

  clampY(y: number): number {
    if (!this.workLayoutRef?.nativeElement) return Math.max(8, y);
    const layoutEl = this.workLayoutRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupRect = groupEl?.getBoundingClientRect();
    const groupHeight = groupRect?.height || groupEl?.offsetHeight || 480;

    // Top boundary: ~8px below the formula bar (red line boundary)
    const minY = 8;
    // Bottom limit: keep entire group (including mascot) visible above sheet tabs / status bar
    const maxY = Math.max(minY, layoutRect.height - groupHeight - 8);

    return Math.min(Math.max(y, minY), maxY);
  }

  onGroupPointerDown(event: PointerEvent): void {
    if (this.bridge.isPaused() || this.isCompleted()) return;
    const target = event.target as HTMLElement | null;
    const isDragHandle = target?.closest('.card-drag-handle');
    if (!isDragHandle) {
      return;
    }

    this.isDragging.set(true);
    this.dragStartPointerX = event.clientX;
    this.dragStartPointerY = event.clientY;

    const groupEl = this.movableGroupRef?.nativeElement;
    this.dragStartGroupX = this.groupX() ?? groupEl?.offsetLeft ?? 0;
    this.dragStartGroupY = this.groupY() ?? groupEl?.offsetTop ?? 0;

    try {
      (event.currentTarget as HTMLElement)?.setPointerCapture(event.pointerId);
    } catch {}
    event.preventDefault();
  }

  onGroupPointerMove(event: PointerEvent): void {
    if (!this.isDragging()) return;
    const dx = event.clientX - this.dragStartPointerX;
    const dy = event.clientY - this.dragStartPointerY;
    const nextX = this.clampX(this.dragStartGroupX + dx);
    const nextY = this.clampY(this.dragStartGroupY + dy);
    this.groupX.set(Math.round(nextX));
    this.groupY.set(Math.round(nextY));
    this.hasUserMoved.set(true);
  }

  onGroupPointerUp(event: PointerEvent): void {
    if (this.isDragging()) {
      this.isDragging.set(false);
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  onGroupPointerCancel(event: PointerEvent): void {
    if (this.isDragging()) {
      this.isDragging.set(false);
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  @HostListener('document:mousemove', ['$event'])
  onDocMouseMove(event: MouseEvent): void {
    if (!this.isDragging()) return;
    const dx = event.clientX - this.dragStartPointerX;
    const dy = event.clientY - this.dragStartPointerY;
    this.groupX.set(Math.round(this.clampX(this.dragStartGroupX + dx)));
    this.groupY.set(Math.round(this.clampY(this.dragStartGroupY + dy)));
    this.hasUserMoved.set(true);
  }

  @HostListener('document:mouseup')
  onDocMouseUp(): void {
    if (this.isDragging()) {
      this.isDragging.set(false);
    }
  }

  onNudgePosition(delta: { dx: number; dy: number }): void {
    if (this.bridge.isPaused() || this.isCompleted()) return;
    const groupEl = this.movableGroupRef?.nativeElement;
    const curX = this.groupX() ?? groupEl?.offsetLeft ?? 0;
    const curY = this.groupY() ?? groupEl?.offsetTop ?? 0;
    this.groupX.set(Math.round(this.clampX(curX + delta.dx)));
    this.groupY.set(Math.round(this.clampY(curY + delta.dy)));
    this.hasUserMoved.set(true);
  }

  private handleViewportOrGroupResize(): void {
    if (this.isDragging()) return;
    if (!this.workLayoutRef?.nativeElement) return;
    if (!this.hasUserMoved()) {
      this.computeInitialPlacement();
    } else if (this.groupX() !== null && this.groupY() !== null) {
      const clampedX = Math.round(this.clampX(this.groupX()!));
      const clampedY = Math.round(this.clampY(this.groupY()!));
      this.groupX.set(clampedX);
      this.groupY.set(clampedY);
      this.cdr.markForCheck();
    }
  }

  @HostListener('window:keydown', ['$event'])
  onWindowKeyDown(event: KeyboardEvent): void {
    if (this.bridge.isPaused() || this.isCompleted()) return;
    if ((event.ctrlKey || event.metaKey) && !this.state.isEditing()) {
      const target = event.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.isContentEditable)
      ) {
        return;
      }
      const key = event.key.toLowerCase();
      if (key === 'z' && !event.shiftKey) {
        event.preventDefault();
        event.stopPropagation();
        this.onUndo();
      } else if (key === 'y' || (key === 'z' && event.shiftKey)) {
        event.preventDefault();
        event.stopPropagation();
        this.onRedo();
      }
    }
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    this.handleViewportOrGroupResize();
  }
}
