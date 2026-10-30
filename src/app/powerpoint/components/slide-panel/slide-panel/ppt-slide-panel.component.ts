import {
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptSlideThumbnailComponent } from '../slide-thumbnail/ppt-slide-thumbnail.component';
import { Slide, ThemeDef } from '../../../models/ppt.types';
import { PptStateService } from '../../../services/ppt-state.service';

@Component({
  selector: 'app-ppt-slide-panel',
  standalone: true,
  imports: [CommonModule, PptSlideThumbnailComponent],
  templateUrl: './ppt-slide-panel.component.html',
  styleUrls: ['./ppt-slide-panel.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptSlidePanelComponent {
  @Input() slides: Slide[] = [];
  @Input() currentSlideIndex: number = 0;
  @Input() theme!: ThemeDef;
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly selectSlide = new EventEmitter<number>();
  @Output() readonly deleteCurrentSlide = new EventEmitter<void>();
  @Output() readonly copyCurrentSlide = new EventEmitter<void>();
  @Output() readonly pasteSlide = new EventEmitter<void>();
  @Output() readonly openSlideContextMenu = new EventEmitter<{
    index: number;
    clientX: number;
    clientY: number;
  }>();

  isDragging = false;
  dragSourceIndex: number | null = null;
  dragTargetIndex: number | null = null;

  private startX = 0;
  private startY = 0;
  private hasMovedBeyondThreshold = false;
  private boundPointerMove = (e: PointerEvent) => this.onWindowPointerMove(e);
  private boundPointerUp = (e: PointerEvent) => this.onWindowPointerUp(e);

  constructor(
    public state: PptStateService,
    private elementRef: ElementRef,
    private cdr: ChangeDetectorRef
  ) {}

  onThumbPointerDown(index: number, event: PointerEvent): void {
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
      const strip = this.elementRef.nativeElement.querySelector('.ppt-slide-strip') as HTMLElement | null;
      if (strip) {
        strip.focus({ preventScroll: true });
      } else {
        const canvasHost = document.querySelector('.ppt-canvas-area') as HTMLElement | null;
        canvasHost?.focus({ preventScroll: true });
      }
    }

    this.dragSourceIndex = index;
    this.dragTargetIndex = null;
    this.startX = event.clientX;
    this.startY = event.clientY;
    this.hasMovedBeyondThreshold = false;
    this.isDragging = false;

    window.addEventListener('pointermove', this.boundPointerMove);
    window.addEventListener('pointerup', this.boundPointerUp);
    window.addEventListener('pointercancel', this.boundPointerUp);
  }

  private onWindowPointerMove(event: PointerEvent): void {
    if (this.dragSourceIndex === null) return;

    if (!this.hasMovedBeyondThreshold) {
      const dist = Math.hypot(event.clientX - this.startX, event.clientY - this.startY);
      if (dist >= 5) {
        this.hasMovedBeyondThreshold = true;
        this.isDragging = true;
      }
    }

    if (this.isDragging) {
      this.updateDragTarget(event.clientY);
      this.cdr.markForCheck();
    }
  }

  private onWindowPointerUp(event: PointerEvent): void {
    window.removeEventListener('pointermove', this.boundPointerMove);
    window.removeEventListener('pointerup', this.boundPointerUp);
    window.removeEventListener('pointercancel', this.boundPointerUp);

    const source = this.dragSourceIndex;
    const target = this.dragTargetIndex;
    const wasDragging = this.isDragging && this.hasMovedBeyondThreshold;

    this.isDragging = false;
    this.dragSourceIndex = null;
    this.dragTargetIndex = null;

    if (wasDragging && source !== null && target !== null) {
      let to: number | null = null;
      if (target < source) {
        to = target;
      } else if (target > source + 1) {
        to = target - 1;
      }

      if (to !== null && to !== source) {
        this.state.moveSlide(source, to);
        this.selectSlide.emit(to);
      }
    } else if (source !== null && !wasDragging) {
      this.state.selectSlide(source);
      this.selectSlide.emit(source);
    }
    this.cdr.markForCheck();
  }

  private updateDragTarget(clientY: number): void {
    const thumbElements = this.elementRef.nativeElement.querySelectorAll('.ppt-thumbnail-wrapper');
    if (!thumbElements || thumbElements.length === 0) return;

    const count = thumbElements.length;
    for (let i = 0; i < count; i++) {
      const rect = thumbElements[i].getBoundingClientRect();
      const midY = rect.top + rect.height / 2;
      if (clientY < midY) {
        this.dragTargetIndex = i;
        return;
      }
    }
    this.dragTargetIndex = count;
  }

  onKeyDown(event: KeyboardEvent): void {
    const isCtrl = event.ctrlKey || event.metaKey;

    if (isCtrl && (event.key === 'ArrowUp' || event.key === 'Up')) {
      event.preventDefault();
      if (this.currentSlideIndex > 0) {
        this.state.moveSlide(this.currentSlideIndex, this.currentSlideIndex - 1);
        this.selectSlide.emit(this.currentSlideIndex - 1);
      }
    } else if (isCtrl && (event.key === 'ArrowDown' || event.key === 'Down')) {
      event.preventDefault();
      if (this.currentSlideIndex < this.slides.length - 1) {
        this.state.moveSlide(this.currentSlideIndex, this.currentSlideIndex + 1);
        this.selectSlide.emit(this.currentSlideIndex + 1);
      }
    } else if (event.key === 'Delete') {
      event.preventDefault();
      this.deleteCurrentSlide.emit();
    } else if (isCtrl && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      this.copyCurrentSlide.emit();
    } else if (isCtrl && event.key.toLowerCase() === 'v') {
      event.preventDefault();
      this.pasteSlide.emit();
    }
  }

  @HostListener('window:keydown', ['$event'])
  onWindowKeyDown(event: KeyboardEvent): void {
    if (event.defaultPrevented) return;
    const active = document.activeElement as HTMLElement | null;
    const isInsideStrip = !!(active && (active.classList.contains('ppt-slide-strip') || active.closest('.ppt-slide-strip')));
    if (isInsideStrip) {
      if (
        event.key === 'Delete' ||
        ((event.ctrlKey || event.metaKey) &&
          (event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'Up' || event.key === 'Down'))
      ) {
        this.onKeyDown(event);
      }
    }
  }

  trackBySlideId(_: number, slide: Slide): string {
    return slide.id || String(_);
  }
}
