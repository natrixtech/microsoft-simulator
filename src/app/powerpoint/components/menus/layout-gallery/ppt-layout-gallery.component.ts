import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { LayoutDef, SlideLayout } from '../../../models/ppt.types';
import { LAYOUT_ORDER, PPT_LAYOUTS } from '../../../constants/ppt-layouts';
import { computePopupPosition, PopupPosition } from '../../../utils/popup-position';

@Component({
  selector: 'app-ppt-layout-gallery',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-layout-gallery.component.html',
  styleUrls: ['./ppt-layout-gallery.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptLayoutGalleryComponent implements OnChanges {
  @Input() isOpen: boolean = false;
  @Input() activeLayout: SlideLayout = 'title';
  @Input() headerTitle: string = 'Office Theme';
  @Input() anchorEl?: HTMLElement | null;

  @Output() readonly selectLayout = new EventEmitter<SlideLayout>();
  @Output() readonly close = new EventEmitter<void>();

  readonly layoutList: LayoutDef[] = LAYOUT_ORDER.map((id) => PPT_LAYOUTS[id]);
  popupPos: PopupPosition = { left: 0, top: 0 };

  constructor(private elementRef: ElementRef) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isOpen'] && this.isOpen) {
      this.recalculatePosition();
    }
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (this.isOpen && !this.elementRef.nativeElement.contains(event.target)) {
      this.close.emit();
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.isOpen) {
      this.close.emit();
    }
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (this.isOpen) {
      this.recalculatePosition();
    }
  }

  onPickLayout(layout: SlideLayout, event: MouseEvent): void {
    event.stopPropagation();
    this.selectLayout.emit(layout);
  }

  private recalculatePosition(): void {
    if (!this.anchorEl) return;
    this.popupPos = computePopupPosition(this.anchorEl, { width: 320, height: 260 });
  }
}
