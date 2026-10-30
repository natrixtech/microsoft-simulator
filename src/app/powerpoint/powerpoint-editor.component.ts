import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  forwardRef,
  HostListener,
  Input,
  NgZone,
  OnDestroy,
  OnInit,
  ViewChild,
  computed,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PluginBridgeService } from '../shared/services/plugin-bridge.service';
import { ActivityPayload, LessonTask } from '../shared/models/activity.types';
import { ConfettiService } from '../shared/services/confetti.service';
import { TaskCardComponent } from '../shared/components/task-card/task-card.component';
import { PptTitleBarComponent } from './components/title-bar/ppt-title-bar.component';
import { PptRibbonComponent } from './components/ribbon/ribbon/ppt-ribbon.component';
import { PptSlidePanelComponent } from './components/slide-panel/slide-panel/ppt-slide-panel.component';
import { PptCanvasComponent } from './components/canvas/ppt-canvas.component';
import { PptNotesBarComponent } from './components/notes-bar/ppt-notes-bar.component';
import { PptStatusBarComponent } from './components/status-bar/ppt-status-bar.component';
import { PptSlideContextMenuComponent } from './components/menus/slide-context-menu/slide-context-menu.component';
import { PptInsertPictureDialogComponent } from './components/dialogs/insert-picture-dialog/insert-picture-dialog.component';
import { PptFormatBackgroundPaneComponent } from './components/panes/format-background-pane/ppt-format-background-pane.component';
import { PptDemoComponent } from './components/demo/ppt-demo.component';
import { PptTextFormatService } from './services/ppt-text-format.service';
import { PptStateService } from './services/ppt-state.service';
import { PptHistoryService } from './services/ppt-history.service';
import { PptActionLoggerService } from './services/ppt-action-logger.service';
import { PptGoalEvaluatorService } from './services/ppt-goal-evaluator.service';
import {
  Deck,
  PlaceholderType,
  PptGoal,
  PptSnapshot,
  RibbonTabId,
  ShapeKind,
  SlideElement,
  SlideLayout,
  SlideTransition,
  TableSlideElement,
  ThemeId,
} from './models/ppt.types';
import { normalizeDeck } from './utils/deck-normalize';
import { clone } from '../shared/utils/editor.utils';

@Component({
  selector: 'app-ppt-editor',
  standalone: true,
  imports: [
    CommonModule,
    TaskCardComponent,
    PptTitleBarComponent,
    PptRibbonComponent,
    PptSlidePanelComponent,
    PptCanvasComponent,
    PptNotesBarComponent,
    PptStatusBarComponent,
    PptSlideContextMenuComponent,
    PptInsertPictureDialogComponent,
    PptFormatBackgroundPaneComponent,
    forwardRef(() => PptDemoComponent),
  ],
  templateUrl: './powerpoint-editor.component.html',
  styleUrls: ['./powerpoint-editor.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptEditorComponent implements OnInit, AfterViewInit, OnDestroy {
  @Input() demoMode = false;
  @Input() demoDeck: Deck | null = null;
  @Input() demoActivity: ActivityPayload | null = null;

  @ViewChild('workLayout', { static: false }) workLayoutRef?: ElementRef<HTMLElement>;
  @ViewChild('canvasComponent', { static: false }) canvasComponent?: PptCanvasComponent;
  @ViewChild('movableGroup', { static: false }) movableGroupRef?: ElementRef<HTMLDivElement>;

  readonly currentActivity = signal<ActivityPayload | null>(null);
  readonly taskIndex = signal<number>(0);
  readonly passedTasks = signal<Set<number>>(new Set());
  readonly isPassed = signal<boolean>(false);
  readonly feedbackMsg = signal<string>('');
  readonly activeTab = signal<RibbonTabId>('home');
  readonly fitPercent = signal<number>(100);

  // Dialogs and Context Menus
  isPictureDialogOpen = false;
  pictureDialogPurpose: 'insert' | 'background' = 'insert';
  isContextMenuOpen = false;
  contextMenuPos = { x: 0, y: 0 };
  contextMenuSlideIndex = 0;
  isDemoOpen = false;
  demoTask: LessonTask | null = null;
  demoStartingDeck: Deck | null = null;

  private preMoveSnapshot: PptSnapshot | null = null;
  private historyPushedForEditSession = false;
  private nudgeTimer: any = null;
  private isNudgeRunActive = false;
  taskStartDeckSnapshot: Deck | null = null;

  // Floating task card dragging (top-right initial placement)
  readonly groupX = signal<number | null>(null);
  readonly groupY = signal<number | null>(null);
  readonly isDragging = signal<boolean>(false);
  readonly hasUserMoved = signal<boolean>(false);
  private dragStartPointerX = 0;
  private dragStartPointerY = 0;
  private dragStartGroupX = 0;
  private dragStartGroupY = 0;
  private resizeObserver?: ResizeObserver;

  private unsubActivity?: () => void;
  private unsubForceFinish?: () => void;
  private unsubPause?: () => void;
  private autoAdvanceTimer: any = null;
  private pendingAdvanceIdx: number | null = null;
  private startingTheme: string = 'simple-light';

  constructor(
    public bridge: PluginBridgeService,
    public state: PptStateService,
    public history: PptHistoryService,
    public actionLogger: PptActionLoggerService,
    public goalEvaluator: PptGoalEvaluatorService,
    public textFormat: PptTextFormatService,
    private confetti: ConfettiService,
    private cdr: ChangeDetectorRef,
    private ngZone: NgZone,
    public elRef: ElementRef<HTMLElement>
  ) {}

  get currentTask(): LessonTask | null {
    const act = this.currentActivity();
    if (!act || !act.tasks) return null;
    return act.tasks[this.taskIndex()] || null;
  }

  get totalTasks(): number {
    return this.currentActivity()?.tasks?.length || 1;
  }

  get isAssessment(): boolean {
    return this.currentActivity()?.mode === 'assessment';
  }

  get isDarkTheme(): boolean {
    return this.bridge.theme() === 'DARK';
  }

  get isFrozen(): boolean {
    return this.bridge.isPaused() || this.bridge.isCompleted();
  }

  get isPaused(): boolean {
    return this.bridge.isPaused();
  }

  get hasShowMeHow(): boolean {
    return !!this.currentTask?.showMeHow;
  }

  ngOnInit(): void {
    if (this.demoMode) {
      if (this.demoActivity) {
        this.currentActivity.set(this.demoActivity);
      }
      if (this.demoDeck) {
        const normalized = normalizeDeck(this.demoDeck);
        this.state.initDeck(normalized);
        this.startingTheme = normalized.theme;
      }
      return;
    }

    const initialAct = this.bridge.activity();
    if (initialAct && initialAct.app === 'powerpoint') {
      this.loadActivity(initialAct);
    }

    this.unsubActivity = this.bridge.onActivity((act: ActivityPayload) => {
      if (act.app === 'powerpoint') {
        this.loadActivity(act);
      }
    });

    this.unsubForceFinish = this.bridge.onForceFinish(() => {
      this.handleForceFinish();
    });

    this.unsubPause = this.bridge.onPauseChange((paused) => {
      if (!paused && this.pendingAdvanceIdx !== null && !this.bridge.isCompleted()) {
        const idx = this.pendingAdvanceIdx;
        this.pendingAdvanceIdx = null;
        this.setupTask(idx, false);
        this.cdr.markForCheck();
      }
    });
  }

  ngAfterViewInit(): void {
    this.textFormat.setRoot(this.elRef.nativeElement);

    if (this.demoMode) {
      return;
    }

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
    if (this.demoMode) return;
    this.unsubActivity?.();
    this.unsubForceFinish?.();
    this.unsubPause?.();
    if (this.autoAdvanceTimer) {
      clearTimeout(this.autoAdvanceTimer);
    }
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = undefined;
    }
  }

  private loadActivity(act: ActivityPayload): void {
    this.currentActivity.set(act);
    this.taskIndex.set(0);
    this.passedTasks.set(new Set());
    this.isPassed.set(false);
    this.feedbackMsg.set('');
    this.hasUserMoved.set(false);
    this.setupTask(0, true);
    requestAnimationFrame(() => {
      this.computeInitialPlacement();
    });
    this.cdr.markForCheck();
  }

  private setupTask(index: number, isInitial: boolean): void {
    this.taskIndex.set(index);
    this.isPassed.set(false);
    this.feedbackMsg.set('');
    this.history.clear();
    this.actionLogger.clear();

    const task = this.currentActivity()?.tasks?.[index];
    if (!task) return;

    if (task.startingDeck !== null && task.startingDeck !== undefined) {
      const normalized = normalizeDeck(task.startingDeck);
      this.state.initDeck(normalized);
      this.taskStartDeckSnapshot = clone(normalized);
    } else if (isInitial) {
      this.state.reset();
      this.taskStartDeckSnapshot = clone(this.state.deck());
    } else {
      this.taskStartDeckSnapshot = clone(this.state.deck());
    }
    this.startingTheme = this.state.deck().theme;
    this.cdr.markForCheck();
  }

  // --- Slide & Ribbon Interactions ---
  readonly isFormatBackgroundOpen = signal<boolean>(false);

  onOpenFormatBackground(): void {
    this.isFormatBackgroundOpen.set(true);
    this.actionLogger.log('openPane', { pane: 'FormatBackground' });
  }

  onCloseFormatBackground(): void {
    this.isFormatBackgroundOpen.set(false);
  }

  onSetSlideBackground(bg: string | undefined): void {
    this.pushHistory();
    this.state.setSlideBackground(bg);
    this.actionLogger.log('formatBackground', { background: bg });
  }

  onResetSlideBackground(): void {
    this.pushHistory();
    this.state.setSlideBackground(undefined);
    this.actionLogger.log('resetBackground');
  }

  onToggleBold(): void {
    this.textFormat.toggleBold();
  }

  onToggleItalic(): void {
    this.textFormat.toggleItalic();
  }

  onToggleUnderline(): void {
    this.textFormat.toggleUnderline();
  }

  onChangeFont(font: string): void {
    this.textFormat.changeFont(font);
  }

  onChangeFontSize(size: number): void {
    this.textFormat.changeFontSize(size);
  }

  onIncreaseFontSize(): void {
    this.textFormat.increaseFontSize();
  }

  onDecreaseFontSize(): void {
    this.textFormat.decreaseFontSize();
  }

  onSetFontColor(color: string): void {
    this.textFormat.setFontColor(color);
  }

  onSetBullets(): void {
    this.textFormat.setBullets();
  }

  onSetNumbering(): void {
    this.textFormat.setNumbering();
  }

  onSetAlign(align: 'left' | 'center' | 'right'): void {
    this.textFormat.setAlign(align);
  }

  onSetLineSpacing(spacing: number): void {
    this.textFormat.setLineSpacing(spacing);
  }


  onSelectSlide(idx: number): void {
    this.historyPushedForEditSession = false;
    this.state.selectSlide(idx);
  }

  onNewSlide(layout?: SlideLayout | void): void {
    this.pushHistory();
    const l = (typeof layout === 'string' ? layout : 'titleAndBody') as SlideLayout;
    this.state.addSlide(l);
  }

  onDeleteSlide(slideIndex?: number): void {
    if (this.state.totalSlides() <= 1) return;
    this.pushHistory();
    const idx = slideIndex !== undefined ? slideIndex : this.state.currentSlideIndex();
    this.state.deleteSlide(idx);
  }

  onOpenSlideContextMenu(event: { index: number; clientX: number; clientY: number }): void {
    this.contextMenuSlideIndex = event.index;
    this.contextMenuPos = { x: event.clientX, y: event.clientY };
    this.isContextMenuOpen = true;
    this.actionLogger.log('openMenu', { menu: 'Slide' });
    this.cdr.markForCheck();
  }

  onChangeLayout(layout: SlideLayout): void {
    this.pushHistory();
    this.state.setLayout(layout);
  }

  onSelectTheme(themeId: ThemeId): void {
    this.pushHistory();
    this.state.setTheme(themeId);
  }

  onSelectTransition(tr: SlideTransition): void {
    this.pushHistory();
    this.state.setSlideTransition(tr);
  }

  onPreviewTransition(): void {
    if (this.isFrozen) return;
    const tr = this.state.currentSlide()?.transition;
    if (!tr || tr.type === 'none') return;
    this.stopTextEditing();
    this.canvasComponent?.playTransition(tr);
  }

  onCopySlide(): void {
    this.state.copyCurrentSlide();
  }

  onPasteSlide(): void {
    if (!this.state.clipboardSlide()) return;
    this.pushHistory();
    this.state.pasteSlide();
  }

  onInsertShape(shape: string): void {
    if (shape === 'textBox') {
      this.onInsertTextBox();
      return;
    }
    this.canvasComponent?.setPendingInsert({
      kind: 'shape',
      shape: shape as ShapeKind,
    });
  }

  onInsertTextBox(): void {
    this.canvasComponent?.setPendingInsert({
      kind: 'textBox',
    });
  }

  onInsertTable(event: { rows: number; cols: number }): void {
    const rows = Math.max(1, Math.min(6, event.rows));
    const cols = Math.max(1, Math.min(6, event.cols));
    const w = Math.min(720, cols * 110);
    const h = Math.min(360, rows * 40);
    const x = Math.round((960 - w) / 2);
    const y = Math.round((540 - h) / 2);

    const cells: string[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: string[] = [];
      for (let c = 0; c < cols; c++) {
        row.push('');
      }
      cells.push(row);
    }

    const tableEl: TableSlideElement = {
      id: 'el-' + Date.now(),
      kind: 'table',
      x,
      y,
      w,
      h,
      rows,
      columns: cols,
      cells,
    };

    this.pushHistory();
    this.state.addElement(tableEl);
  }

  onOpenPictures(purpose: 'insert' | 'background' = 'insert'): void {
    this.pictureDialogPurpose = purpose;
    this.isPictureDialogOpen = true;
    this.cdr.markForCheck();
  }

  onClosePictureDialog(): void {
    this.isPictureDialogOpen = false;
    this.pictureDialogPurpose = 'insert';
    this.cdr.markForCheck();
  }

  onPictureChosen(file: string): void {
    if (this.pictureDialogPurpose === 'background') {
      this.onSetSlideBackground(file);
    } else {
      this.onInsertPicture(file);
    }
    this.pictureDialogPurpose = 'insert';
  }

  onInsertPicture(file: string): void {
    const w = 400; // max 50% slide width (480)
    const h = 250;
    const x = Math.round((960 - w) / 2);
    const y = Math.round((540 - h) / 2);

    const imageEl: SlideElement = {
      id: 'el-' + Date.now(),
      kind: 'image',
      file,
      alt: 'Picture',
      x,
      y,
      w,
      h,
    };

    this.pushHistory();
    this.state.addElement(imageEl);
  }

  onSelectElement(id: string | null): void {
    this.historyPushedForEditSession = false;
    this.state.setSelectedElement(id);
  }

  onAddElement(element: SlideElement): void {
    if (element.kind !== 'textBox') {
      this.pushHistory();
    }
    this.state.addElement(element);
    if (element.kind === 'textBox') {
      this.historyPushedForEditSession = false;
      setTimeout(() => {
        this.canvasComponent?.slideView?.focusTextBox(element.id);
      }, 0);
    }
  }

  onUpdateElementBounds(event: {
    elementId?: string;
    x: number;
    y: number;
    w: number;
    h: number;
    isFinal: boolean;
  }): void {
    const selId = event.elementId || this.state.selectedElementId();
    if (!selId) return;

    if (!this.preMoveSnapshot) {
      this.preMoveSnapshot = this.state.getSnapshot();
    }

    this.state.updateElement(selId, {
      x: event.x,
      y: event.y,
      w: event.w,
      h: event.h,
    });

    if (event.isFinal) {
      if (this.preMoveSnapshot) {
        this.history.push(this.preMoveSnapshot);
        this.preMoveSnapshot = null;
      }
    }
  }

  private ensureEditSessionHistory(): void {
    if (!this.historyPushedForEditSession) {
      this.pushHistory();
      this.historyPushedForEditSession = true;
    }
  }

  onUpdateTableCellText(event: {
    elementId: string;
    row: number;
    col: number;
    text: string;
  }): void {
    this.ensureEditSessionHistory();
    this.state.updateTableCell(event.elementId, event.row, event.col, event.text);
  }

  onUpdateTextBoxText(event: { elementId: string; text: string; html?: string }): void {
    this.ensureEditSessionHistory();
    this.state.updateElement(event.elementId, { text: event.text, html: event.html });
  }

  onUpdatePlaceholder(event: { field: PlaceholderType; text: string; html?: string }): void {
    this.ensureEditSessionHistory();
    this.state.updatePlaceholderText(this.state.currentSlideIndex(), event.field, event.text, event.html);
  }

  onStartEditPlaceholder(field: PlaceholderType): void {
    const cur = this.state.currentSlide();
    if (!cur) return;
    this.historyPushedForEditSession = false;
    this.state.setEditingPlaceholder({ slideId: cur.id, field });
  }

  onStopEditPlaceholder(): void {
    this.historyPushedForEditSession = false;
    this.state.setEditingPlaceholder(null);
  }

  @HostListener('document:mousedown', ['$event'])
  onDocumentMouseDown(event: MouseEvent): void {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;

    if (this.demoMode) {
      if (!this.elRef.nativeElement.contains(target)) return;
    } else {
      if (target?.closest('.ppt-demo-dialog, .demo-modal-backdrop')) return;
    }

    const isRibbonOrPopup = target?.closest(
      '.ppt-ribbon, app-ppt-ribbon, .ppt-folded-panel, .ppt-ribbon-overflow-panel, app-ppt-color-picker, .ppt-ribbon-popup'
    );
    const isEditable = target?.closest(
      '.ppt-editable-text, .ppt-element-textbox[contenteditable="true"]'
    );

    if (!isRibbonOrPopup && !isEditable) {
      if (this.state.editingPlaceholder()) {
        this.onStopEditPlaceholder();
      }
      const selId = this.state.selectedElementId();
      if (selId) {
        const el = this.state.selectedElement();
        if (el?.kind === 'textBox' && !target?.closest('.ppt-selection-frame')) {
          if (!el.text?.trim() && !el.html?.trim()) {
            this.state.deleteElement(selId);
          } else {
            this.state.setSelectedElement(null);
          }
          this.historyPushedForEditSession = false;
        }
      }
    }
  }

  onSelectAll(): void {
    if (this.canvasComponent?.pendingInsert) {
      this.canvasComponent.cancelPendingInsert();
    }
    this.state.selectAllElements();
    this.cdr.markForCheck();
  }

  onSelectObjects(): void {
    if (this.canvasComponent?.pendingInsert) {
      this.canvasComponent.cancelPendingInsert();
    }
    this.state.setSelectedElement(null);
    this.cdr.markForCheck();
  }

  onRefit(): void {
    this.canvasComponent?.recalculateFit();
  }

  onTabChange(tab: RibbonTabId): void {
    this.activeTab.set(tab);
  }

  // --- History (Undo/Redo) ---

  private pushHistory(): void {
    this.history.push(this.state.getSnapshot());
  }

  onUndo(): void {
    if (!this.history.canUndo()) return;
    const restored = this.history.undo(this.state.getSnapshot());
    if (restored) {
      this.state.restoreSnapshot(restored);
      this.actionLogger.log('undo', { slide: this.state.currentSlideIndex() + 1 });
    }
  }

  onRedo(): void {
    if (!this.history.canRedo()) return;
    const restored = this.history.redo(this.state.getSnapshot());
    if (restored) {
      this.state.restoreSnapshot(restored);
      this.actionLogger.log('redo', { slide: this.state.currentSlideIndex() + 1 });
    }
  }

  @HostListener('window:keydown', ['$event'])
  onGlobalKeyDown(event: KeyboardEvent): void {
    if (event.defaultPrevented || this.isFrozen) return;

    const target = event.target instanceof Element ? event.target : null;
    if (this.demoMode) {
      if (target && !this.elRef.nativeElement.contains(target)) return;
    } else {
      if (target?.closest('.ppt-demo-dialog, .demo-modal-backdrop')) return;
    }

    if (event.key === 'Escape') {
      if (this.isContextMenuOpen) {
        this.isContextMenuOpen = false;
        this.cdr.markForCheck();
        return;
      }
      if (this.isPictureDialogOpen) {
        this.onClosePictureDialog();
        return;
      }
      if (this.canvasComponent?.pendingInsert) {
        this.canvasComponent.cancelPendingInsert();
        return;
      }
    }

    const isCtrl = event.ctrlKey || event.metaKey;
    if (isCtrl && this.state.isTextActive()) {
      const k = event.key.toLowerCase();
      if (k === 'b') {
        event.preventDefault();
        event.stopPropagation();
        this.onToggleBold();
        return;
      }
      if (k === 'i') {
        event.preventDefault();
        event.stopPropagation();
        this.onToggleItalic();
        return;
      }
      if (k === 'u') {
        event.preventDefault();
        event.stopPropagation();
        this.onToggleUnderline();
        return;
      }
    }

    if (isCtrl && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      event.stopPropagation();
      if (event.shiftKey) {
        this.onRedo();
      } else {
        this.onUndo();
      }
      return;
    }

    if (isCtrl && event.key.toLowerCase() === 'y') {
      event.preventDefault();
      event.stopPropagation();
      this.onRedo();
      return;
    }

    if (isCtrl && event.key.toLowerCase() === 'a') {
      const activeEl = document.activeElement;
      const isInput =
        activeEl &&
        (activeEl.tagName === 'INPUT' ||
          activeEl.tagName === 'TEXTAREA' ||
          activeEl.tagName === 'TD' ||
          activeEl.closest('td[contenteditable="true"]') !== null ||
          activeEl.getAttribute('contenteditable') === 'true' ||
          (activeEl as HTMLElement).isContentEditable);

      if (!isInput && !this.state.isTextActive() && !this.state.editingPlaceholder()) {
        event.preventDefault();
        event.stopPropagation();
        this.onSelectAll();
        return;
      }
    }

    if (event.key === 'Delete' || event.key === 'Backspace') {
      const activeEl = document.activeElement;
      const isInput =
        activeEl &&
        (activeEl.tagName === 'INPUT' ||
          activeEl.tagName === 'TEXTAREA' ||
          activeEl.tagName === 'TD' ||
          activeEl.closest('td[contenteditable="true"]') !== null ||
          activeEl.getAttribute('contenteditable') === 'true' ||
          (activeEl as HTMLElement).isContentEditable);

      if (!isInput && (this.state.selectedElementIds().length > 0 || this.state.selectedElementId())) {
        event.preventDefault();
        event.stopPropagation();
        this.pushHistory();
        this.state.deleteSelectedElements();
        this.cdr.markForCheck();
      }
    }

    const isArrow = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key);
    if (isArrow) {
      const activeEl = document.activeElement;
      const isInput =
        activeEl &&
        (activeEl.tagName === 'INPUT' ||
          activeEl.tagName === 'TEXTAREA' ||
          activeEl.tagName === 'TD' ||
          activeEl.closest('td[contenteditable="true"]') !== null ||
          activeEl.getAttribute('contenteditable') === 'true' ||
          (activeEl as HTMLElement).isContentEditable);

      let ids = this.state.selectedElementIds();
      if (ids.length === 0 && this.state.selectedElementId()) {
        ids = [this.state.selectedElementId()!];
      }

      if (
        !isInput &&
        !this.state.isTextActive() &&
        !this.state.editingPlaceholder() &&
        !this.canvasComponent?.slideView?.focusedTextBoxId &&
        !this.canvasComponent?.slideView?.focusedCellKey &&
        ids.length > 0
      ) {
        event.preventDefault();
        event.stopPropagation();

        if (!this.isNudgeRunActive) {
          this.pushHistory();
          this.isNudgeRunActive = true;
        }
        if (this.nudgeTimer) {
          clearTimeout(this.nudgeTimer);
        }
        this.nudgeTimer = setTimeout(() => {
          this.isNudgeRunActive = false;
          this.nudgeTimer = null;
        }, 1000);

        const step = (event.ctrlKey || event.metaKey) ? 1 : 8;
        let dx = 0;
        let dy = 0;
        if (event.key === 'ArrowLeft') dx = -step;
        if (event.key === 'ArrowRight') dx = step;
        if (event.key === 'ArrowUp') dy = -step;
        if (event.key === 'ArrowDown') dy = step;

        const curSlide = this.state.currentSlide();
        if (curSlide) {
          const selectedEls = curSlide.elements.filter((e) => ids.includes(e.id));
          let clampedDx = dx;
          let clampedDy = dy;
          for (const el of selectedEls) {
            if (clampedDx < 0) {
              clampedDx = Math.max(clampedDx, -el.x);
            } else if (clampedDx > 0) {
              clampedDx = Math.min(clampedDx, 960 - (el.x + el.w));
            }
            if (clampedDy < 0) {
              clampedDy = Math.max(clampedDy, -el.y);
            } else if (clampedDy > 0) {
              clampedDy = Math.min(clampedDy, 540 - (el.y + el.h));
            }
          }

          if (clampedDx !== 0 || clampedDy !== 0) {
            for (const el of selectedEls) {
              this.state.updateElement(el.id, {
                x: el.x + clampedDx,
                y: el.y + clampedDy,
              });
            }
            this.cdr.markForCheck();
          }
        }
        return;
      }
    }
  }

  // --- Task Card & Assessment Protocol ---

  checkWork(): void {
    if (this.isFrozen || this.isPassed()) return;

    const task = this.currentTask;
    if (!task) return;

    const goals = (task.goals as PptGoal[]) || [];
    if (goals.length === 0) {
      this.handlePassTask();
      return;
    }

    const passes = this.goalEvaluator.evaluate(goals, this.state.deck(), {
      currentSlideIndex: this.state.currentSlideIndex(),
      startingTheme: this.startingTheme,
      actionLog: this.actionLogger.actions(),
    });

    if (passes) {
      this.handlePassTask();
    } else {
      const hint = task.hint || 'Not quite right. Try again.';
      this.feedbackMsg.set(hint);
      this.cdr.markForCheck();
    }
  }

  private handlePassTask(): void {
    this.isPassed.set(true);
    this.feedbackMsg.set('');
    this.confetti.celebrate();

    const currentIdx = this.taskIndex();
    const passed = new Set(this.passedTasks());
    passed.add(currentIdx);
    this.passedTasks.set(passed);

    if (this.isAssessment) {
      const nextIdx = currentIdx + 1;
      if (nextIdx < this.totalTasks) {
        this.autoAdvanceTimer = setTimeout(() => {
          this.ngZone.run(() => {
            if (this.bridge.isPaused()) {
              this.pendingAdvanceIdx = nextIdx;
              return;
            }
            this.setupTask(nextIdx, false);
          });
        }, 1200);
      } else {
        this.bridge.sendComplete(true, this.totalTasks, this.totalTasks);
      }
    } else {
      this.bridge.sendComplete(true, 1, 1);
    }
    this.cdr.markForCheck();
  }

  onStartAgain(): void {
    if (this.bridge.isCompleted()) return;
    this.setupTask(0, true);
  }

  stopTextEditing(): void {
    const active = document.activeElement as HTMLElement | null;
    if (
      active &&
      this.elRef.nativeElement.contains(active) &&
      (active.classList.contains('ppt-editable-text') ||
        active.classList.contains('ppt-element-textbox') ||
        active.isContentEditable)
    ) {
      active.blur();
    }
    if (this.state.editingPlaceholder()) {
      this.onStopEditPlaceholder();
    }
    this.cdr.markForCheck();
  }

  onShowDemo(): void {
    if (!this.currentTask) return;
    this.stopTextEditing();
    this.demoTask = this.currentTask;
    const starting = this.currentTask.startingDeck
      ? normalizeDeck(this.currentTask.startingDeck)
      : clone(this.taskStartDeckSnapshot || this.state.deck());
    this.demoStartingDeck = starting;
    this.isDemoOpen = true;
    this.cdr.markForCheck();
  }

  onCloseDemo(): void {
    this.isDemoOpen = false;
    this.demoTask = null;
    this.demoStartingDeck = null;
    this.cdr.markForCheck();
    try {
      this.elRef?.nativeElement?.focus?.();
    } catch {}
  }

  private handleForceFinish(): void {
    if (this.isDemoOpen) {
      this.onCloseDemo();
    }
    this.pendingAdvanceIdx = null;
    if (this.bridge.isCompleted()) return;
    const passed = this.passedTasks().size;
    const total = this.totalTasks;
    const allPassed = passed === total && total > 0;
    this.bridge.sendComplete(allPassed, passed, total);
  }

  // --- Task Card Dragging & Placement (Excel standard) ---

  computeInitialPlacement(): void {
    if (!this.workLayoutRef?.nativeElement) return;
    const layoutEl = this.workLayoutRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    if (layoutRect.width === 0 || layoutRect.height === 0) return;

    const groupEl = this.movableGroupRef?.nativeElement;
    const groupRect = groupEl?.getBoundingClientRect();
    const groupWidth = groupRect?.width || groupEl?.offsetWidth || 300;
    const groupHeight = groupRect?.height || groupEl?.offsetHeight || 480;

    // Top-right placement: top = 10 px, right = 12 px
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

    const minY = 8;
    const maxY = Math.max(minY, layoutRect.height - groupHeight - 8);
    return Math.min(Math.max(y, minY), maxY);
  }

  onGroupPointerDown(event: PointerEvent): void {
    if (this.isFrozen) return;
    const target = event.target instanceof Element ? event.target : null;
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
    if (this.isFrozen) return;
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
}

export { PptEditorComponent as PowerpointEditorComponent };
