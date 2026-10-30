import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { SlideElement } from '../../../models/ppt.types';

export type ResizeHandleType = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

@Component({
  selector: 'app-ppt-selection-frame',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-selection-frame.component.html',
  styleUrls: ['./ppt-selection-frame.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptSelectionFrameComponent {
  @Input() element!: SlideElement;
  @Input() scale: number = 1;

  @Output() readonly updateBounds = new EventEmitter<{
    x: number;
    y: number;
    w: number;
    h: number;
    isFinal: boolean;
  }>();

  readonly handles: ResizeHandleType[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

  isDragging = false;
  isResizing = false;
  activeHandle: ResizeHandleType | null = null;

  private startPointerX = 0;
  private startPointerY = 0;
  private dragScale = 1;
  private startElX = 0;
  private startElY = 0;
  private startElW = 0;
  private startElH = 0;
  private hasChanged = false;

  get isImage(): boolean {
    return this.element?.kind === 'image';
  }

  get isTextEditable(): boolean {
    return this.element?.kind === 'textBox' || this.element?.kind === 'table';
  }

  private focusCanvasHost(): void {
    const canvasHost = document.querySelector('.ppt-canvas-area') as HTMLElement | null;
    canvasHost?.focus({ preventScroll: true });
  }

  onEdgePointerDown(event: PointerEvent): void {
    this.onFramePointerDown(event);
  }

  onFramePointerDown(event: PointerEvent): void {
    if ((event.target as HTMLElement).classList.contains('resize-handle')) {
      return;
    }
    // If it's textEditable and the click wasn't on an edge strip, don't drag
    if (this.isTextEditable && !(event.target as HTMLElement).classList.contains('edge-strip')) {
      return;
    }
    event.stopPropagation();
    event.preventDefault();

    if (this.isTextEditable) {
      const activeEl = document.activeElement as HTMLElement | null;
      if (activeEl && (activeEl.tagName === 'TD' || activeEl.closest('td') || activeEl.isContentEditable)) {
        activeEl.blur();
      }
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
      }
    }

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
      this.focusCanvasHost();
    }

    this.isDragging = true;
    this.hasChanged = false;
    const wrap = (event.currentTarget as HTMLElement | null)?.closest('.slide-transform-wrapper') as HTMLElement | null;
    this.dragScale = wrap ? (wrap.getBoundingClientRect().width / 960 || 1) : (this.scale || 1);
    this.startPointerX = event.clientX;
    this.startPointerY = event.clientY;
    this.startElX = this.element.x;
    this.startElY = this.element.y;
    this.startElW = this.element.w;
    this.startElH = this.element.h;

    try {
      (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    } catch {}
  }

  onHandlePointerDown(handle: ResizeHandleType, event: PointerEvent): void {
    event.stopPropagation();
    event.preventDefault();

    this.focusCanvasHost();

    if (this.isTextEditable) {
      const activeEl = document.activeElement as HTMLElement | null;
      if (activeEl && (activeEl.tagName === 'TD' || activeEl.closest('td') || activeEl.isContentEditable)) {
        activeEl.blur();
      }
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
      }
    }

    this.isResizing = true;
    this.activeHandle = handle;
    this.hasChanged = false;
    const wrap = (event.currentTarget as HTMLElement | null)?.closest('.slide-transform-wrapper') as HTMLElement | null;
    this.dragScale = wrap ? (wrap.getBoundingClientRect().width / 960 || 1) : (this.scale || 1);
    this.startPointerX = event.clientX;
    this.startPointerY = event.clientY;
    this.startElX = this.element.x;
    this.startElY = this.element.y;
    this.startElW = this.element.w;
    this.startElH = this.element.h;

    try {
      (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
    } catch {}
  }

  @HostListener('window:pointermove', ['$event'])
  onPointerMove(event: PointerEvent): void {
    if (!this.isDragging && !this.isResizing) return;
    const dx = (event.clientX - this.startPointerX) / this.dragScale;
    const dy = (event.clientY - this.startPointerY) / this.dragScale;

    if (this.isDragging) {
      const nextX = Math.round(Math.max(0, Math.min(960 - this.startElW, this.startElX + dx)));
      const nextY = Math.round(Math.max(0, Math.min(540 - this.startElH, this.startElY + dy)));
      if (nextX !== this.element.x || nextY !== this.element.y) {
        this.hasChanged = true;
        this.updateBounds.emit({
          x: nextX,
          y: nextY,
          w: this.element.w,
          h: this.element.h,
          isFinal: false,
        });
      }
    } else if (this.isResizing && this.activeHandle) {
      this.calculateResize(dx, dy, event.shiftKey);
    }
  }

  @HostListener('window:pointerup', ['$event'])
  onPointerUp(event: PointerEvent): void {
    if (this.isDragging || this.isResizing) {
      if (this.hasChanged) {
        this.updateBounds.emit({
          x: this.element.x,
          y: this.element.y,
          w: this.element.w,
          h: this.element.h,
          isFinal: true,
        });
      }
      this.isDragging = false;
      this.isResizing = false;
      this.activeHandle = null;
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  @HostListener('pointercancel', ['$event'])
  onPointerCancel(event: PointerEvent): void {
    this.onPointerUp(event);
  }

  private calculateResize(dx: number, dy: number, shiftKey?: boolean): void {
    let nextX = this.startElX;
    let nextY = this.startElY;
    let nextW = this.startElW;
    let nextH = this.startElH;
    const minSize = 20;
    const handle = this.activeHandle!;
    const isCorner = ['nw', 'ne', 'se', 'sw'].includes(handle);
    const keepAspect = isCorner && (this.isImage || !!shiftKey);
    const aspect = this.startElW / (this.startElH || 1);

    if (handle.includes('e')) {
      nextW = Math.max(minSize, this.startElW + dx);
    }
    if (handle.includes('s')) {
      nextH = Math.max(minSize, this.startElH + dy);
    }
    if (handle.includes('w')) {
      const allowedDx = Math.min(dx, this.startElW - minSize);
      nextX = this.startElX + allowedDx;
      nextW = this.startElW - allowedDx;
    }
    if (handle.includes('n')) {
      const allowedDy = Math.min(dy, this.startElH - minSize);
      nextY = this.startElY + allowedDy;
      nextH = this.startElH - allowedDy;
    }

    if (keepAspect) {
      nextH = Math.max(minSize, Math.round(nextW / aspect));
      if (handle === 'nw') {
        nextX = this.startElX + (this.startElW - nextW);
        nextY = this.startElY + (this.startElH - nextH);
      } else if (handle === 'ne') {
        nextY = this.startElY + (this.startElH - nextH);
      } else if (handle === 'sw') {
        nextX = this.startElX + (this.startElW - nextW);
      }
    }

    nextX = Math.round(nextX);
    nextY = Math.round(nextY);
    nextW = Math.round(nextW);
    nextH = Math.round(nextH);

    this.hasChanged = true;
    this.updateBounds.emit({
      x: nextX,
      y: nextY,
      w: nextW,
      h: nextH,
      isFinal: false,
    });
  }
}

