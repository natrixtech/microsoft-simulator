import {
  AfterViewInit,
  ApplicationRef,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  OnDestroy,
  OnInit,
  Output,
  SimpleChanges,
  ViewChild,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { LessonTask, ActivityPayload } from '../../../shared/models/activity.types';
import { Deck, PptGoal } from '../../models/ppt.types';
import { PptDemoHostComponent } from './ppt-demo-host.component';
import { buildDemoSteps, DemoAction, DemoStep } from '../../utils/ppt-demo-steps';
import { clone } from '../../../shared/utils/editor.utils';
import { PptSafeHtmlPipe } from '../../pipes/ppt-safe-html.pipe';

export interface DemoStepItem {
  label: string;
  goalIndex: number;
  actions: DemoAction[];
  state: 'pending' | 'active' | 'done';
}

@Component({
  selector: 'app-ppt-demo',
  standalone: true,
  imports: [CommonModule, PptDemoHostComponent, PptSafeHtmlPipe],
  templateUrl: './ppt-demo.component.html',
  styleUrls: ['./ppt-demo.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptDemoComponent implements OnInit, AfterViewInit, OnChanges, OnDestroy {
  @ViewChild('demoReplicaArea', { static: false }) replicaAreaRef?: ElementRef<HTMLElement>;
  @ViewChild('demoHost', { static: false }) demoHostRef?: PptDemoHostComponent;

  @Input() task!: LessonTask;
  @Input() startingDeck!: Deck;
  @Input() images: Array<{ file: string; alt?: string }> = [];
  @Input() assetBaseUrl: string = './assets/';
  @Input() isPaused: boolean = false;
  @Input() isDarkTheme: boolean = false;

  @Output() readonly close = new EventEmitter<void>();

  readonly steps = signal<DemoStepItem[]>([]);
  readonly showToast = signal<boolean>(false);
  readonly cursorPixelX = signal<number>(100);
  readonly cursorPixelY = signal<number>(100);
  readonly isClicking = signal<boolean>(false);
  readonly isCursorVisible = signal<boolean>(false);

  readonly scaleTransform = signal<string>('scale(0.8)');
  readonly replicaHeight = signal<number>(860);
  readonly replicaReady = signal<boolean>(false);
  readonly replicaAlive = signal<boolean>(true);

  currentDemoDeck: Deck | null = null;
  demoActivity: ActivityPayload | null = null;

  private resizeObserver?: ResizeObserver;
  private isCancelled = false;
  private activeTimers: any[] = [];
  private rafId?: number;
  private lastActionTarget: string | null = null;

  constructor(
    private elRef: ElementRef<HTMLElement>,
    private cdr: ChangeDetectorRef,
    private appRef: ApplicationRef
  ) {}

  ngOnInit(): void {
    this.initReplicaState();
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isPaused'] && !changes['isPaused'].firstChange) {
      this.cdr.markForCheck();
    }
  }

  ngAfterViewInit(): void {
    if (typeof ResizeObserver !== 'undefined' && this.replicaAreaRef) {
      this.resizeObserver = new ResizeObserver(() => {
        this.updateScaling();
      });
      this.resizeObserver.observe(this.replicaAreaRef.nativeElement);
    }

    let frameCount = 0;
    const checkScale = () => {
      if (this.isCancelled) return;
      frameCount++;
      const scaled = this.updateScaling();
      if (scaled || frameCount >= 20) {
        this.replicaReady.set(true);
        this.cdr.detectChanges();
        this.play();
      } else {
        this.rafId = requestAnimationFrame(checkScale);
      }
    };

    checkScale();
  }

  ngOnDestroy(): void {
    this.cancelAll();
    this.resizeObserver?.disconnect();
  }

  @HostListener('window:keydown.escape', ['$event'])
  onEscape(e: KeyboardEvent): void {
    e.stopPropagation();
    this.onClose();
  }

  onClose(): void {
    this.cancelAll();
    this.close.emit();
  }

  replay(): void {
    this.cancelAll();                 // isCancelled = true, timers cleared
    this.showToast.set(false);
    this.isCursorVisible.set(false);
    this.replicaReady.set(false);
    this.replicaAlive.set(false);     // destroy old replica
    this.cdr.detectChanges();
    this.initReplicaState();          // fresh currentDemoDeck + steps (all 'pending')
    this.replicaAlive.set(true);      // create a new replica
    this.cdr.detectChanges();
    setTimeout(() => {
      this.updateScaling();
      this.replicaReady.set(true);
      this.cdr.detectChanges();
      this.isCancelled = false;       // only now allow the new run (old async steps stay stopped)
      this.play();
    }, 250);
  }

  get formattedFirstInstruction(): string {
    const raw = this.task?.instructions?.[0] || this.task?.explanation || '';
    if (!raw) return '';
    // Replace `value` with <span class="instruction-chip">value</span>
    return raw.replace(/`([^`]+)`/g, '<span class="instruction-chip">$1</span>');
  }

  getToastTitle(): string {
    const t = this.task?.title || 'complete this task';
    return t.toLowerCase().replace(/\.$/, '');
  }

  private initReplicaState(): void {
    this.currentDemoDeck = clone(this.startingDeck);
    this.demoActivity = {
      app: 'powerpoint',
      mode: 'lesson',
      title: this.task.title,
      assetBaseUrl: this.assetBaseUrl,
      images: this.images,
      theme: this.isDarkTheme ? 'DARK' : 'LIGHT',
      tasks: [this.task],
    };

    const rawSteps = buildDemoSteps(
      this.task.goals as PptGoal[],
      this.currentDemoDeck,
      this.task,
      this.currentDemoDeck.theme,
      this.images
    );

    this.steps.set(
      rawSteps.map((s) => ({
        ...s,
        state: 'pending',
      }))
    );
  }

  private forceReplicaRender(): void {
    try {
      (this.demoHostRef?.editorComponent as any)?.cdr?.detectChanges();
    } catch {
      // ignore
    }
    this.appRef.tick();
    this.cdr.markForCheck();
  }

  updateScaling(): boolean {
    if (!this.replicaAreaRef) return false;
    const el = this.replicaAreaRef.nativeElement;
    const w = el.clientWidth;
    const h = el.clientHeight;
    if (w <= 0 || h <= 0) return false;

    const scale = Math.min((w - 16) / 1380, 1);
    const logicalH = Math.max(600, Math.round((h - 16) / scale));
    this.replicaHeight.set(logicalH);
    this.scaleTransform.set(`scale(${scale})`);
    this.cdr.markForCheck();
    return true;
  }

  private cancelAll(): void {
    this.isCancelled = true;
    for (const t of this.activeTimers) {
      clearTimeout(t);
    }
    this.activeTimers = [];
    if (this.rafId) {
      cancelAnimationFrame(this.rafId);
      this.rafId = undefined;
    }
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

  private setStepState(idx: number, state: 'pending' | 'active' | 'done'): void {
    const next = this.steps().map((item, i) =>
      i === idx ? { ...item, state } : { ...item }
    );
    this.steps.set(next);
    this.cdr.markForCheck();
    this.cdr.detectChanges();
  }

  async play(): Promise<void> {
    this.lastActionTarget = null;
    await this.delay(500);
    if (this.isCancelled) return;

    this.isCursorVisible.set(true);

    const stepList = this.steps();
    for (let i = 0; i < stepList.length; i++) {
      if (this.isCancelled) return;
      this.setStepState(i, 'active');

      const step = stepList[i];
      const goal = this.task.goals[step.goalIndex] as PptGoal;

      // Before step: check if goal already passes
      if (this.demoHostRef?.goalEvaluator && this.demoHostRef?.state) {
        const passes = this.demoHostRef.goalEvaluator.evaluate(
          [goal],
          this.demoHostRef.state.deck(),
          {
            currentSlideIndex: this.demoHostRef.state.currentSlideIndex(),
            startingTheme: this.demoHostRef.state.deck().theme || 'simple-light',
            actionLog: this.demoHostRef.actionLogger.actions(),
          }
        );
        if (passes) {
          await this.delay(300);
          this.setStepState(i, 'done');
          continue;
        }
      }

      for (const action of step.actions) {
        if (this.isCancelled) return;
        await this.executeAction(action);
      }

      this.setStepState(i, 'done');
      await this.delay(400);
    }

    // QA Self-Check: evaluate all goals after last step
    if (this.demoHostRef?.goalEvaluator && this.demoHostRef?.state) {
      for (const goal of this.task.goals as PptGoal[]) {
        const passes = this.demoHostRef.goalEvaluator.evaluate(
          [goal],
          this.demoHostRef.state.deck(),
          {
            currentSlideIndex: this.demoHostRef.state.currentSlideIndex(),
            startingTheme: this.currentDemoDeck?.theme || 'simple-light',
            actionLog: this.demoHostRef.actionLogger.actions(),
          }
        );
        if (!passes) {
          console.error('[ShowMeHow] demo did not satisfy goal', goal);
        }
      }
    }

    this.showToast.set(true);
    this.cdr.markForCheck();
    this.cdr.detectChanges();
  }

  private async executeAction(action: DemoAction): Promise<void> {
    if (this.isCancelled) return;

    if (action.kind === 'ensureTab' && action.tab) {
      this.ensureRibbonTab(action.tab);
      await this.delay(350);
      return;
    }

    if (action.kind === 'wait') {
      await this.delay(action.durationMs || 500);
      return;
    }

    if (action.kind === 'drag' && action.target === 'slide-canvas') {
      await this.executeSlideDraw(action);
      this.lastActionTarget = action.target;
      return;
    }

    if (action.kind === 'drag' && action.dragToTarget) {
      await this.executeSlideOrderDrag(action);
      this.lastActionTarget = action.target || null;
      return;
    }

    if (action.kind === 'drag' && action.target) {
      await this.executeElementDrag(action);
      this.lastActionTarget = action.target;
      return;
    }

    if (action.target) {
      const targetEl = await this.findTarget(action.target);
      if (!targetEl) {
        console.warn('[ShowMeHow] missing target', action.target);
        return;
      }

      await this.moveCursorToElement(targetEl);
      if (this.isCancelled) return;

      if (action.kind === 'click') {
        await this.performClick(targetEl);
        if (this.isCancelled) return;
        if (action.target === 'slide-canvas') {
          this.demoHostRef?.editorComponent?.stopTextEditing();
        }
      } else if (action.kind === 'rightClick') {
        await this.performRightClick(targetEl);
        if (this.isCancelled) return;
      } else if (action.kind === 'dblClick') {
        await this.performDblClick(targetEl);
        if (this.isCancelled) return;
      } else if (action.kind === 'select') {
        await this.performSelect(targetEl, action.value);
        if (this.isCancelled) return;
      } else if (action.kind === 'selectText' && action.text) {
        await this.performSelectText(targetEl, action.text);
        if (this.isCancelled) return;
      }
    }

    if (this.isCancelled) return;

    if (action.kind === 'key') {
      await this.performKey(action.key || '');
    } else if (action.kind === 'type' && action.text) {
      await this.performType(action.text);
    }

    if (action.target) {
      this.lastActionTarget = action.target;
    }

    if (this.isCancelled) return;
    await this.delay(300);
  }

  private ensureRibbonTab(tabId: 'home' | 'insert' | 'design' | 'transitions'): void {
    const editor = this.demoHostRef?.editorComponent;
    if (editor && editor.activeTab() !== tabId) {
      editor.onTabChange(tabId);
      this.forceReplicaRender();
    }
  }

  private async findTarget(name: string): Promise<HTMLElement | null> {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    if (!hostEl) return null;

    // Switch ribbon tab automatically if target belongs to a tab
    if (name.startsWith('home-') || name === 'home-tab') {
      this.ensureRibbonTab('home');
    } else if (
      name.startsWith('insert-') ||
      name === 'insert-tab' ||
      name.startsWith('picture-item-') ||
      name === 'picture-dialog-insert' ||
      name.startsWith('table-cell-')
    ) {
      this.ensureRibbonTab('insert');
    } else if (
      name.startsWith('design-') ||
      name === 'design-tab' ||
      name.startsWith('format-bg-')
    ) {
      this.ensureRibbonTab('design');
    } else if (name.startsWith('transitions-') || name === 'transitions-tab') {
      this.ensureRibbonTab('transitions');
    }

    this.forceReplicaRender();

    const queryTarget = (): HTMLElement | null => {
      if (name === 'slide-element-last') {
        const elements = hostEl.querySelectorAll('[data-demo-target^="slide-element-"]');
        let maxNum = -1;
        let lastEl: HTMLElement | null = null;
        elements.forEach((el) => {
          const targetAttr = el.getAttribute('data-demo-target') || '';
          const match = targetAttr.match(/^slide-element-(\d+)$/);
          if (match) {
            const num = parseInt(match[1], 10);
            if (num > maxNum) {
              maxNum = num;
              lastEl = el as HTMLElement;
            }
          }
        });
        return lastEl;
      }
      return hostEl.querySelector(`[data-demo-target="${name}"]`) as HTMLElement | null;
    };

    let target = queryTarget();

    const tryOpeners = async () => {
      // 1. Menu openers - only click when not already open
      if (name.startsWith('insert-shape-')) {
        const isAlreadyOpen = !!hostEl.querySelector('.shapes-gallery-popup');
        const opener = hostEl.querySelector('[data-demo-target="insert-shapes"]') as HTMLElement | null;
        if (opener && !isAlreadyOpen) {
          opener.click();
          this.forceReplicaRender();
        }
      } else if (name.startsWith('table-cell-')) {
        const isAlreadyOpen = !!hostEl.querySelector('.table-picker-popup');
        const opener = hostEl.querySelector('[data-demo-target="insert-table"]') as HTMLElement | null;
        if (opener && !isAlreadyOpen) {
          opener.click();
          this.forceReplicaRender();
        }
      } else if (name === 'design-format-background') {
        const isAlreadyOpen = !!hostEl.querySelector('.bg-menu-popup');
        const opener = hostEl.querySelector('[data-demo-target="design-background-styles"]') as HTMLElement | null;
        if (opener && !isAlreadyOpen) {
          opener.click();
          this.forceReplicaRender();
        }
      } else if (name.startsWith('color-swatch-')) {
        const isAlreadyOpen = !!hostEl.querySelector('app-ppt-color-picker, .fixed-color-picker');
        if (!isAlreadyOpen) {
          const isFormatBg =
            this.lastActionTarget === 'format-bg-color-btn' ||
            (this.lastActionTarget?.startsWith('format-bg-') && hostEl.querySelector('.ppt-format-bg-pane'));
          const openerSelector = isFormatBg
            ? '[data-demo-target="format-bg-color-btn"]'
            : '[data-demo-target="home-font-color-arrow"]';
          const opener = hostEl.querySelector(openerSelector) as HTMLElement | null;
          if (opener) {
            opener.click();
            this.forceReplicaRender();
          }
        }
      } else if (name.startsWith('layout-choice-')) {
        const isAlreadyOpen = !!hostEl.querySelector('app-ppt-layout-gallery, .layout-gallery-popup');
        const opener = hostEl.querySelector('[data-demo-target="home-layout"]') as HTMLElement | null;
        if (opener && !isAlreadyOpen) {
          opener.click();
          this.forceReplicaRender();
        }
      }

      // 2. Folded group openers - only click when not already open
      const groupMap: Record<string, string> = {
        'insert-textbox': 'text',
        'insert-shapes': 'illustrations',
        'insert-table': 'tables',
        'insert-picture-btn': 'images',
        'insert-icons': 'illustrations',
        'insert-link': 'links',
        'design-background-styles': 'customize',
      };
      const gId = groupMap[name];
      if (gId) {
        const foldedBtn = hostEl.querySelector(`[data-demo-target="folded-group-${gId}"]`) as HTMLElement | null;
        if (foldedBtn && !foldedBtn.classList.contains('active')) {
          foldedBtn.click();
          this.forceReplicaRender();
        }
      } else {
        const foldedBtns = hostEl.querySelectorAll('[data-demo-target^="folded-group-"]');
        for (let b = 0; b < foldedBtns.length; b++) {
          const btn = foldedBtns[b] as HTMLElement;
          if (!btn.classList.contains('active')) {
            btn.click();
            this.forceReplicaRender();
            const found = queryTarget();
            if (found) return found;
          }
        }
      }
      return null;
    };

    if (!target) {
      const opened = await tryOpeners();
      if (opened) return opened;
      target = queryTarget();
    }

    // Poll every 100 ms for up to 3000 ms
    if (!target) {
      for (let i = 0; i < 30; i++) {
        await this.delay(100);
        if (this.isCancelled) return null;
        this.forceReplicaRender();
        target = queryTarget();
        if (target) break;

        if (i === 4 || i === 12 || i === 20) {
          const opened = await tryOpeners();
          if (opened) return opened;
          target = queryTarget();
          if (target) break;
        }
      }
    }

    return target;
  }

  private async moveCursorToElement(el: HTMLElement): Promise<void> {
    if (!this.replicaAreaRef) return;
    const targetRect = el.getBoundingClientRect();
    const areaRect = this.replicaAreaRef.nativeElement.getBoundingClientRect();

    const targetX = targetRect.left + targetRect.width / 2 - areaRect.left;
    const targetY = targetRect.top + targetRect.height / 2 - areaRect.top;

    this.cursorPixelX.set(Math.round(targetX));
    this.cursorPixelY.set(Math.round(targetY));
    this.cdr.markForCheck();

    await this.delay(550);
  }

  private async performClick(el: HTMLElement): Promise<void> {
    await this.delay(180);
    this.isClicking.set(true);
    await this.delay(180);
    this.isClicking.set(false);

    const rect = el.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;

    const eventInit: PointerEventInit = {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      isPrimary: true,
      pointerType: 'mouse',
      clientX,
      clientY,
    };

    el.dispatchEvent(new PointerEvent('pointerdown', { ...eventInit, button: 0, buttons: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 1 }));
    el.dispatchEvent(new PointerEvent('pointerup', { ...eventInit, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX, clientY }));

    const editableTarget =
      el.tagName === 'TD' ||
      el.isContentEditable ||
      el.getAttribute('contenteditable') === 'true' ||
      el.classList.contains('ppt-element-textbox') ||
      el.classList.contains('ppt-editable-text')
        ? el
        : (el.querySelector(
            'td, [contenteditable="true"], .ppt-element-textbox, .ppt-editable-text'
          ) as HTMLElement | null);

    if (editableTarget) {
      editableTarget.focus();
      const sel = window.getSelection();
      if (sel) {
        const range = document.createRange();
        range.selectNodeContents(editableTarget);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    }

    this.forceReplicaRender();
  }

  private async performRightClick(el: HTMLElement): Promise<void> {
    await this.delay(180);
    this.isClicking.set(true);
    await this.delay(180);
    this.isClicking.set(false);

    const rect = el.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;

    const eventInit: PointerEventInit = {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      isPrimary: true,
      pointerType: 'mouse',
      clientX,
      clientY,
    };

    el.dispatchEvent(new PointerEvent('pointerdown', { ...eventInit, button: 2, buttons: 2 }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX, clientY, button: 2, buttons: 2 }));
    el.dispatchEvent(new PointerEvent('pointerup', { ...eventInit, button: 2, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, clientX, clientY, button: 2, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX, clientY, button: 2 }));

    this.forceReplicaRender();
  }

  private async performDblClick(el: HTMLElement): Promise<void> {
    await this.delay(150);
    this.isClicking.set(true);
    await this.delay(120);
    this.isClicking.set(false);
    await this.delay(80);
    this.isClicking.set(true);
    await this.delay(120);
    this.isClicking.set(false);

    const rect = el.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;

    const eventInit: PointerEventInit = {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      isPrimary: true,
      pointerType: 'mouse',
      clientX,
      clientY,
    };

    // First click sequence
    el.dispatchEvent(new PointerEvent('pointerdown', { ...eventInit, button: 0, buttons: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 1 }));
    el.dispatchEvent(new PointerEvent('pointerup', { ...eventInit, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX, clientY }));

    // Second click sequence + dblclick
    el.dispatchEvent(new PointerEvent('pointerdown', { ...eventInit, button: 0, buttons: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 1 }));
    el.dispatchEvent(new PointerEvent('pointerup', { ...eventInit, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, clientX, clientY, button: 0, buttons: 0 }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX, clientY }));
    el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, clientX, clientY }));

    this.forceReplicaRender();
  }

  private async performSelect(el: HTMLElement, value: any): Promise<void> {
    const selectEl = el as HTMLSelectElement;
    if (selectEl && value !== undefined) {
      selectEl.value = String(value);
      selectEl.dispatchEvent(new Event('change', { bubbles: true }));
      this.forceReplicaRender();
    }
  }

  private async performSelectText(containerEl: HTMLElement, text: string): Promise<void> {
    const editable = await this.getLiveEditable();
    const targetElement = editable || containerEl;
    targetElement.focus();

    const sel = window.getSelection();
    if (!sel) return;

    const walker = document.createTreeWalker(targetElement, NodeFilter.SHOW_TEXT, null);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      const val = node.nodeValue || '';
      const idx = val.indexOf(text);
      if (idx !== -1) {
        const range = document.createRange();
        range.setStart(node, idx);
        range.setEnd(node, idx + text.length);
        sel.removeAllRanges();
        sel.addRange(range);
        break;
      }
    }

    if (this.demoHostRef?.textFormat) {
      this.demoHostRef.textFormat.saveSelection();
    }
    document.dispatchEvent(new Event('selectionchange'));
    await this.delay(150);
  }

  private findLiveEditable(): HTMLElement | null {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    if (!hostEl) return null;
    return hostEl.querySelector(
      '.ppt-editable-text[contenteditable="true"], .ppt-element-textbox[contenteditable="true"], .ppt-slide-rendered-table td[contenteditable="true"], .modal-dialog input:not(.ppt-search-input), .ppt-format-bg-pane input:not(.ppt-search-input)'
    ) as HTMLElement | null;
  }

  private async getLiveEditable(): Promise<HTMLElement | null> {
    let editable = this.findLiveEditable();
    if (!editable) {
      for (let i = 0; i < 6; i++) {
        await this.delay(60);
        if (this.isCancelled) return null;
        editable = this.findLiveEditable();
        if (editable) break;
      }
    }
    return editable;
  }

  private async performKey(key: string): Promise<void> {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    const activeEl = document.activeElement as HTMLElement | null;
    const isReplicaActive =
      activeEl &&
      hostEl?.contains(activeEl) &&
      (activeEl.tagName === 'TD' ||
        activeEl.classList.contains('ppt-editable-text') ||
        activeEl.classList.contains('ppt-element-textbox') ||
        activeEl.isContentEditable);
    const editable = isReplicaActive ? activeEl : await this.getLiveEditable();

    if (key === 'Ctrl+A') {
      if (editable) {
        editable.focus();
        const sel = window.getSelection();
        if (sel) {
          const range = document.createRange();
          range.selectNodeContents(editable);
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }
      return;
    }

    if (key === 'Enter') {
      if (editable) {
        editable.focus();
        try {
          document.execCommand('insertParagraph', false);
        } catch {
          document.execCommand('insertText', false, '\n');
        }
        editable.dispatchEvent(new Event('input', { bubbles: true }));
        this.cdr.markForCheck();
      }
    }
  }

  private async performType(text: string): Promise<void> {
    const placeCaretAtEnd = (el: HTMLElement) => {
      el.focus();
      if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
        const len = el.value.length;
        el.setSelectionRange(len, len);
      } else {
        const sel = window.getSelection();
        if (sel) {
          const range = document.createRange();
          range.selectNodeContents(el);
          range.collapse(false);
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }
    };

    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    const activeEl = document.activeElement as HTMLElement | null;
    const isReplicaActive =
      activeEl &&
      hostEl?.contains(activeEl) &&
      (activeEl.tagName === 'TD' ||
        activeEl.classList.contains('ppt-editable-text') ||
        activeEl.classList.contains('ppt-element-textbox') ||
        activeEl.isContentEditable);

    let editable: HTMLElement | null = isReplicaActive ? activeEl : null;

    if (!editable) {
      const startTime = Date.now();
      while (!this.isCancelled && Date.now() - startTime < 2000) {
        this.forceReplicaRender();
        editable = this.findLiveEditable();
        if (editable) {
          if (document.activeElement !== editable) {
            placeCaretAtEnd(editable);
          }
          if (document.activeElement === editable) {
            break;
          }
        }
        await this.delay(50);
      }
    }

    if (!editable) {
      console.warn('[ShowMeHow] no editable for typing');
      return;
    }

    if (document.activeElement !== editable) {
      placeCaretAtEnd(editable);
    }

    for (const ch of text) {
      if (this.isCancelled) return;
      if (ch === '\n') {
        try {
          document.execCommand('insertParagraph', false);
        } catch {
          document.execCommand('insertText', false, '\n');
        }
      } else {
        document.execCommand('insertText', false, ch);
      }
      editable.dispatchEvent(new Event('input', { bubbles: true }));
      this.cdr.markForCheck();
      await this.delay(45);
    }
  }

  private async executeSlideDraw(action: DemoAction): Promise<void> {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    const slideWrapper = hostEl?.querySelector(
      '.slide-transform-wrapper, [data-demo-target="slide-canvas"]'
    ) as HTMLElement | null;

    if (!slideWrapper || !this.replicaAreaRef) return;
    const sRect = slideWrapper.getBoundingClientRect();
    const areaRect = this.replicaAreaRef.nativeElement.getBoundingClientRect();

    const fromX = action.fromCoords?.x || 200;
    const fromY = action.fromCoords?.y || 200;
    const toX = action.toCoords?.x || 500;
    const toY = action.toCoords?.y || 400;

    const startClientX = sRect.left + (fromX / 960) * sRect.width;
    const startClientY = sRect.top + (fromY / 540) * sRect.height;
    const endClientX = sRect.left + (toX / 960) * sRect.width;
    const endClientY = sRect.top + (toY / 540) * sRect.height;

    this.cursorPixelX.set(Math.round(startClientX - areaRect.left));
    this.cursorPixelY.set(Math.round(startClientY - areaRect.top));
    await this.delay(450);

    slideWrapper.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        clientX: startClientX,
        clientY: startClientY,
        button: 0,
        buttons: 1,
        pointerId: 1,
      })
    );

    this.isClicking.set(true);

    const steps = 6;
    for (let s = 1; s <= steps; s++) {
      const curClientX = startClientX + ((endClientX - startClientX) * s) / steps;
      const curClientY = startClientY + ((endClientY - startClientY) * s) / steps;
      this.cursorPixelX.set(Math.round(curClientX - areaRect.left));
      this.cursorPixelY.set(Math.round(curClientY - areaRect.top));
      slideWrapper.dispatchEvent(
        new PointerEvent('pointermove', {
          bubbles: true,
          cancelable: true,
          clientX: curClientX,
          clientY: curClientY,
          button: 0,
          buttons: 1,
          pointerId: 1,
        })
      );
      await this.delay(50);
    }

    this.isClicking.set(false);

    slideWrapper.dispatchEvent(
      new PointerEvent('pointerup', {
        bubbles: true,
        cancelable: true,
        clientX: endClientX,
        clientY: endClientY,
        button: 0,
        buttons: 0,
        pointerId: 1,
      })
    );

    this.cdr.markForCheck();
    await this.delay(300);
  }

  private async executeSlideOrderDrag(action: DemoAction): Promise<void> {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    if (!hostEl || !this.replicaAreaRef || !action.target || !action.dragToTarget) return;

    const fromEl = hostEl.querySelector(
      `[data-demo-target="${action.target}"]`
    ) as HTMLElement | null;
    const toEl = hostEl.querySelector(
      `[data-demo-target="${action.dragToTarget}"]`
    ) as HTMLElement | null;

    if (!fromEl || !toEl) return;

    const fRect = fromEl.getBoundingClientRect();
    const tRect = toEl.getBoundingClientRect();
    const areaRect = this.replicaAreaRef.nativeElement.getBoundingClientRect();

    const startX = fRect.left + fRect.width / 2;
    const startY = fRect.top + fRect.height / 2;
    const endX = tRect.left + tRect.width / 2;
    const isDraggingUp = startY > (tRect.top + tRect.height / 2);
    // When dragging up, move slightly above midY so updateDragTarget selects index i; when down, slightly below
    const endY = (tRect.top + tRect.height / 2) + (isDraggingUp ? -8 : 8);

    this.cursorPixelX.set(Math.round(startX - areaRect.left));
    this.cursorPixelY.set(Math.round(startY - areaRect.top));
    await this.delay(450);

    fromEl.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        pointerId: 1,
        isPrimary: true,
        pointerType: 'mouse',
        clientX: startX,
        clientY: startY,
        button: 0,
        buttons: 1,
      })
    );

    this.isClicking.set(true);

    const steps = 10;
    for (let s = 1; s <= steps; s++) {
      const curX = startX + ((endX - startX) * s) / steps;
      const curY = startY + ((endY - startY) * s) / steps;
      this.cursorPixelX.set(Math.round(curX - areaRect.left));
      this.cursorPixelY.set(Math.round(curY - areaRect.top));
      window.dispatchEvent(
        new PointerEvent('pointermove', {
          bubbles: true,
          cancelable: true,
          pointerId: 1,
          isPrimary: true,
          pointerType: 'mouse',
          clientX: curX,
          clientY: curY,
          button: 0,
          buttons: 1,
        })
      );
      await this.delay(45);
    }

    this.isClicking.set(false);

    window.dispatchEvent(
      new PointerEvent('pointerup', {
        bubbles: true,
        cancelable: true,
        pointerId: 1,
        isPrimary: true,
        pointerType: 'mouse',
        clientX: endX,
        clientY: endY,
        button: 0,
        buttons: 0,
      })
    );

    this.forceReplicaRender();
    await this.delay(300);
  }

  private async executeElementDrag(action: DemoAction): Promise<void> {
    const hostEl = this.demoHostRef?.elRef?.nativeElement;
    const slideWrapper = hostEl?.querySelector(
      '.slide-transform-wrapper, [data-demo-target="slide-canvas"]'
    ) as HTMLElement | null;

    if (!hostEl || !slideWrapper || !this.replicaAreaRef || !action.target) return;
    const targetEl = await this.findTarget(action.target);

    const frame = hostEl.querySelector('.selection-frame') as HTMLElement | null;
    const downTarget = action.target === 'slide-element-last' && frame ? frame : (targetEl || slideWrapper);

    const sRect = slideWrapper.getBoundingClientRect();
    const areaRect = this.replicaAreaRef.nativeElement.getBoundingClientRect();
    const tRect = downTarget.getBoundingClientRect();

    const startX = action.fromCoords
      ? sRect.left + (action.fromCoords.x / 960) * sRect.width
      : tRect.left + tRect.width / 2;
    const startY = action.fromCoords
      ? sRect.top + (action.fromCoords.y / 540) * sRect.height
      : tRect.top + tRect.height / 2;

    const endX = action.toCoords
      ? sRect.left + (action.toCoords.x / 960) * sRect.width
      : startX;
    const endY = action.toCoords
      ? sRect.top + (action.toCoords.y / 540) * sRect.height
      : startY;

    this.cursorPixelX.set(Math.round(startX - areaRect.left));
    this.cursorPixelY.set(Math.round(startY - areaRect.top));
    await this.delay(350);

    const currentSlide = this.demoHostRef?.state?.currentSlide();
    const elements = currentSlide?.elements || [];
    const el =
      action.target === 'slide-element-last'
        ? elements[elements.length - 1]
        : (this.demoHostRef?.state?.selectedElement() || elements[elements.length - 1]);

    if (el && this.demoHostRef?.state?.selectedElementId() !== el.id) {
      this.demoHostRef?.state?.setSelectedElement(el.id);
    }

    const isResize = action.target.startsWith('handle-');
    const startElX = el?.x ?? 0;
    const startElY = el?.y ?? 0;
    const startElW = el?.w ?? 0;
    const startElH = el?.h ?? 0;

    const deltaSlideX = action.toCoords && action.fromCoords ? action.toCoords.x - action.fromCoords.x : 0;
    const deltaSlideY = action.toCoords && action.fromCoords ? action.toCoords.y - action.fromCoords.y : 0;

    this.isClicking.set(true);

    const steps = 10;
    for (let s = 1; s <= steps; s++) {
      if (this.isCancelled) return;
      const curX = startX + ((endX - startX) * s) / steps;
      const curY = startY + ((endY - startY) * s) / steps;
      this.cursorPixelX.set(Math.round(curX - areaRect.left));
      this.cursorPixelY.set(Math.round(curY - areaRect.top));

      if (el && this.demoHostRef?.editorComponent) {
        if (isResize) {
          const curW = Math.max(20, Math.round(startElW + (deltaSlideX * s) / steps));
          const curH = Math.max(20, Math.round(startElH + (deltaSlideY * s) / steps));
          this.demoHostRef.editorComponent.onUpdateElementBounds({
            elementId: el.id,
            x: startElX,
            y: startElY,
            w: curW,
            h: curH,
            isFinal: s === steps,
          });
        } else {
          const curElX = Math.round(startElX + (deltaSlideX * s) / steps);
          const curElY = Math.round(startElY + (deltaSlideY * s) / steps);
          this.demoHostRef.editorComponent.onUpdateElementBounds({
            elementId: el.id,
            x: curElX,
            y: curElY,
            w: startElW,
            h: startElH,
            isFinal: s === steps,
          });
        }
      }

      this.forceReplicaRender();
      await this.delay(40);
    }

    this.isClicking.set(false);
    this.forceReplicaRender();
    await this.delay(300);
  }
}
