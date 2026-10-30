import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  NgZone,
  OnDestroy,
  Output,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptSlideViewComponent } from '../slide-view/ppt-slide-view.component';
import { PptSelectionFrameComponent } from '../selection/ppt-selection-frame/ppt-selection-frame.component';
import {
  PlaceholderType,
  ShapeKind,
  Slide,
  SlideElement,
  SlideTransition,
  ThemeDef,
} from '../../models/ppt.types';
import {
  computeSlideFit,
  SlideViewportFit,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
} from '../../utils/slide-geometry';

export type PendingInsertMode =
  | { kind: 'shape'; shape: ShapeKind }
  | { kind: 'textBox' }
  | null;

@Component({
  selector: 'app-ppt-canvas',
  standalone: true,
  imports: [CommonModule, PptSlideViewComponent, PptSelectionFrameComponent],
  templateUrl: './ppt-canvas.component.html',
  styleUrls: ['./ppt-canvas.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptCanvasComponent implements AfterViewInit, OnDestroy {
  @ViewChild('canvasContainer', { static: true }) canvasContainer!: ElementRef<HTMLElement>;
  @ViewChild('slideWrapper', { static: false }) slideWrapperRef?: ElementRef<HTMLElement>;
  @ViewChild(PptSlideViewComponent) slideView?: PptSlideViewComponent;

  @Input() slide!: Slide;
  @Input() theme!: ThemeDef;
  @Input() selectedElementId: string | null = null;
  @Input() selectedElementIds: string[] = [];
  @Input() selectedElement: SlideElement | null = null;
  @Input() selectedElements: SlideElement[] = [];
  @Input() editingPlaceholderField: PlaceholderType | null = null;
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly selectElement = new EventEmitter<string | null>();
  @Output() readonly updateElementBounds = new EventEmitter<{
    elementId?: string;
    x: number;
    y: number;
    w: number;
    h: number;
    isFinal: boolean;
  }>();
  @Output() readonly addElement = new EventEmitter<SlideElement>();
  @Output() readonly updateTableCellText = new EventEmitter<{
    elementId: string;
    row: number;
    col: number;
    text: string;
  }>();
  @Output() readonly updateTextBoxText = new EventEmitter<{
    elementId: string;
    text: string;
    html?: string;
  }>();
  @Output() readonly updatePlaceholder = new EventEmitter<{
    field: PlaceholderType;
    text: string;
    html?: string;
  }>();
  @Output() readonly startEditPlaceholder = new EventEmitter<PlaceholderType>();
  @Output() readonly stopEditPlaceholder = new EventEmitter<void>();
  @Output() readonly fitPercentChange = new EventEmitter<number>();

  fit: SlideViewportFit = {
    scale: 1,
    width: SLIDE_WIDTH,
    height: SLIDE_HEIGHT,
    offsetX: 0,
    offsetY: 0,
    percent: 100,
  };

  pendingInsert: PendingInsertMode = null;
  isDrawingNewElement = false;
  drawStartX = 0;
  drawStartY = 0;
  drawPreviewRect: { x: number; y: number; w: number; h: number } | null = null;

  private resizeObserver: ResizeObserver | null = null;

  constructor(
    private ngZone: NgZone,
    private cdr: ChangeDetectorRef
  ) {}

  ngAfterViewInit(): void {
    if (typeof ResizeObserver !== 'undefined' && this.canvasContainer?.nativeElement) {
      this.resizeObserver = new ResizeObserver(() => {
        this.ngZone.run(() => {
          this.recalculateFit();
        });
      });
      this.resizeObserver.observe(this.canvasContainer.nativeElement);
    }
    this.recalculateFit();
  }

  previewClass: string | null = null;
  private previewTimer: any = null;

  playTransition(tr: SlideTransition): void {
    clearTimeout(this.previewTimer);
    let cls: string | null = null;
    if (tr.type === 'fade') cls = 'ppt-tr-fade';
    else if (tr.type === 'push') cls = tr.direction === 'fromLeft' ? 'ppt-tr-push-left' : 'ppt-tr-push-right';
    else if (tr.type === 'flip') cls = 'ppt-tr-flip';
    else if (tr.type === 'cube') cls = 'ppt-tr-cube';
    if (!cls) return;
    this.previewClass = null; // restart if clicked again
    this.cdr.detectChanges();
    requestAnimationFrame(() => {
      this.previewClass = cls;
      this.cdr.markForCheck();
      this.previewTimer = setTimeout(() => {
        this.previewClass = null;
        this.cdr.markForCheck();
      }, 1050);
    });
  }

  ngOnDestroy(): void {
    clearTimeout(this.previewTimer);
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }
  }

  recalculateFit(): void {
    const el = this.canvasContainer?.nativeElement;
    if (!el) return;

    const w = el.clientWidth; const h = el.clientHeight;
    if (w === 0 || h === 0) return;

    this.fit = computeSlideFit(w, h, 24);
    this.fitPercentChange.emit(this.fit.percent);
    this.cdr.markForCheck();
  }

  setPendingInsert(mode: PendingInsertMode): void {
    this.pendingInsert = mode;
    this.isDrawingNewElement = false;
    this.drawPreviewRect = null;
    this.cdr.markForCheck();
  }

  cancelPendingInsert(): void {
    this.pendingInsert = null;
    this.isDrawingNewElement = false;
    this.drawPreviewRect = null;
    this.cdr.markForCheck();
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.pendingInsert) {
      this.cancelPendingInsert();
    }
  }

  trackByElementId(_: number, el: SlideElement): string {
    return el.id;
  }

  onUpdateBounds(elementId: string, bounds: { x: number; y: number; w: number; h: number; isFinal: boolean }): void {
    this.updateElementBounds.emit({ elementId, ...bounds });
  }

  onCanvasClick(event: MouseEvent): void {
    if (event.target === this.canvasContainer.nativeElement) {
      this.stopEditPlaceholder.emit();
      this.selectElement.emit(null);
    }
  }

  onSlidePointerDown(event: PointerEvent): void {
    if (!this.pendingInsert) return;
    event.stopPropagation();
    event.preventDefault();

    const target = event.target as HTMLElement | null;
    const isEditable = target && (
      target.isContentEditable ||
      target.tagName === 'INPUT' ||
      target.tagName === 'TEXTAREA' ||
      target.tagName === 'TD' ||
      !!target.closest('[contenteditable="true"]') ||
      !!target.closest('td')
    );
    if (!isEditable) {
      this.canvasContainer?.nativeElement?.focus({ preventScroll: true });
    }

    const slideEl = this.slideWrapperRef?.nativeElement;
    if (!slideEl) return;
    const rect = slideEl.getBoundingClientRect();
    const s = rect.width / 960 || 1;

    this.drawStartX = Math.round((event.clientX - rect.left) / s);
    this.drawStartY = Math.round((event.clientY - rect.top) / s);
    this.isDrawingNewElement = true;
    this.drawPreviewRect = {
      x: this.drawStartX,
      y: this.drawStartY,
      w: 0,
      h: 0,
    };

    try {
      (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    } catch {}
  }

  onSlidePointerMove(event: PointerEvent): void {
    if (!this.isDrawingNewElement || !this.drawPreviewRect) return;
    const slideEl = this.slideWrapperRef?.nativeElement;
    if (!slideEl) return;
    const rect = slideEl.getBoundingClientRect();
    const s = rect.width / 960 || 1;

    const curX = Math.round((event.clientX - rect.left) / s);
    const curY = Math.round((event.clientY - rect.top) / s);

    let dx = curX - this.drawStartX;
    let dy = curY - this.drawStartY;
    let w = Math.abs(dx);
    let h = Math.abs(dy);

    if (event.shiftKey) {
      const side = Math.max(w, h);
      w = side;
      h = side;
      let x = dx >= 0 ? this.drawStartX : this.drawStartX - side;
      let y = dy >= 0 ? this.drawStartY : this.drawStartY - side;
      x = Math.max(0, Math.min(960 - w, x));
      y = Math.max(0, Math.min(540 - h, y));
      this.drawPreviewRect = { x, y, w, h };
    } else {
      const x = Math.max(0, Math.min(this.drawStartX, curX));
      const y = Math.max(0, Math.min(this.drawStartY, curY));
      w = Math.min(960 - x, w);
      h = Math.min(540 - y, h);
      this.drawPreviewRect = { x, y, w, h };
    }
  }

  onSlidePointerUp(event: PointerEvent): void {
    if (!this.isDrawingNewElement || !this.pendingInsert) return;
    try {
      (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
    } catch {}

    const slideEl = this.slideWrapperRef?.nativeElement;
    if (!slideEl) return;
    const rect = slideEl.getBoundingClientRect();
    const s = rect.width / 960 || 1;

    const curX = Math.round((event.clientX - rect.left) / s);
    const curY = Math.round((event.clientY - rect.top) / s);

    let dx = curX - this.drawStartX;
    let dy = curY - this.drawStartY;
    let w = Math.abs(dx);
    let h = Math.abs(dy);
    let x: number;
    let y: number;

    if (w < 10 && h < 10) {
      if (this.pendingInsert.kind === 'shape') {
        w = 72;
        h = 72;
      } else {
        w = 160;
        h = 120;
      }
      x = Math.max(0, Math.min(960 - w, this.drawStartX));
      y = Math.max(0, Math.min(540 - h, this.drawStartY));
    } else {
      if (event.shiftKey) {
        const side = Math.max(w, h);
        w = Math.max(20, side);
        h = Math.max(20, side);
        x = dx >= 0 ? this.drawStartX : this.drawStartX - side;
        y = dy >= 0 ? this.drawStartY : this.drawStartY - side;
      } else {
        w = Math.max(20, w);
        h = Math.max(20, h);
        x = Math.min(this.drawStartX, curX);
        y = Math.min(this.drawStartY, curY);
      }
      x = Math.max(0, Math.min(960 - w, x));
      y = Math.max(0, Math.min(540 - h, y));
    }

    const mode = this.pendingInsert;
    this.pendingInsert = null;
    this.isDrawingNewElement = false;
    this.drawPreviewRect = null;
    this.cdr.markForCheck();

    const id = 'el-' + Date.now();
    if (mode.kind === 'shape') {
      const newEl: SlideElement = {
        id,
        kind: 'shape',
        shape: mode.shape,
        x,
        y,
        w,
        h,
        fill: this.theme?.accentColors?.[3] || '#0078d4',
      };
      this.addElement.emit(newEl);
    } else if (mode.kind === 'textBox') {
      const newEl: SlideElement = {
        id,
        kind: 'textBox',
        text: '',
        html: '',
        x,
        y,
        w,
        h,
        fontSize: 20,
        fontFamily: this.theme?.bodyFont || 'Arial',
        color: this.theme?.bodyColor || '#000000',
        align: 'left',
      };
      this.addElement.emit(newEl);
      setTimeout(() => {
        this.slideView?.focusTextBox(id);
      }, 0);
    }
  }
}
