import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostBinding,
  HostListener,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptIconComponent } from '../ppt-icon/ppt-icon.component';
import { computePopupPosition, PopupPosition } from '../../../utils/popup-position';

@Component({
  selector: 'app-ppt-ribbon-group',
  standalone: true,
  imports: [CommonModule, PptIconComponent],
  templateUrl: './ppt-ribbon-group.component.html',
  styleUrls: ['./ppt-ribbon-group.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptRibbonGroupComponent implements OnChanges {
  @Input() label: string = '';
  @Input() groupId: string = '';
  @Input() icon: string = '';
  @Input() isFolded: boolean = false;
  @Input() isOpen: boolean = false;
  @Input() isCompact: boolean = false;
  @Input() isOverflow: boolean = false;
  @HostBinding('class.in-overflow') get hostInOverflow(): boolean { return this.isOverflow; }

  @Output() readonly toggleFold = new EventEmitter<string>();
  @Output() readonly closeFold = new EventEmitter<void>();

  @ViewChild('foldedBtn', { static: false }) foldedBtnRef?: ElementRef<HTMLElement>;

  panelPos: PopupPosition = { left: 0, top: 0 };
  private lastAnchor: HTMLElement | null = null;

  constructor(private elementRef: ElementRef) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isOpen'] && this.isOpen) {
      this.recalculatePosition();
    }
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (this.isOpen && !this.elementRef.nativeElement.contains(event.target)) {
      this.closeFold.emit();
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.isOpen) {
      this.closeFold.emit();
    }
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (this.isOpen) {
      this.recalculatePosition();
    }
  }

  onFoldButtonClick(event: MouseEvent): void {
    event.stopPropagation();
    this.lastAnchor = event.currentTarget as HTMLElement;
    this.recalculatePosition();
    this.toggleFold.emit(this.groupId);
  }

  onPanelClick(event: MouseEvent): void {
    event.stopPropagation();
    const target = event.target as HTMLElement | null;
    if (!target) return;

    // Do NOT close if the user clicked inside an active popup/gallery/color-picker
    if (
      target.closest(
        '.ppt-shape-picker, .ppt-table-picker, .ppt-color-picker, .ppt-dropdown-menu, .ppt-gallery-menu, .ppt-popup-menu, .ppt-gallery'
      )
    ) {
      return;
    }

    const btn = target.closest('button, [role="button"]') as HTMLElement | null;
    if (!btn) return;

    if ((btn as HTMLButtonElement).disabled || btn.getAttribute('aria-disabled') === 'true') {
      return;
    }

    // Do not close if it is a submenu/gallery opener (e.g. insert-shapes, insert-table, or an item with an arrow)
    const isSubmenuOpener =
      btn.classList.contains('ppt-split-arrow') ||
      btn.classList.contains('ppt-gallery-arrow') ||
      btn.classList.contains('dropdown-arrow') ||
      btn.hasAttribute('aria-haspopup') ||
      btn.querySelector('.ppt-split-arrow, app-ppt-icon[name="chevron-down"]') !== null ||
      btn.getAttribute('data-action')?.includes('open-') ||
      btn.getAttribute('data-demo-target')?.includes('shapes') ||
      btn.getAttribute('data-demo-target')?.includes('table') ||
      btn.getAttribute('data-demo-target')?.includes('layout') ||
      btn.getAttribute('data-demo-target')?.includes('themes');

    if (isSubmenuOpener) {
      return;
    }

    setTimeout(() => {
      this.closeFold.emit();
    }, 0);
  }

  private recalculatePosition(): void {
    const anchor = this.lastAnchor || this.foldedBtnRef?.nativeElement;
    if (!anchor) return;
    this.panelPos = computePopupPosition(anchor, { width: 220, height: 160 });
  }
}
