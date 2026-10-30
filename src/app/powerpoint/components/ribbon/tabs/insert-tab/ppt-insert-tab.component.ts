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
import { PptIconComponent } from '../../ppt-icon/ppt-icon.component';
import { PptRibbonGroupComponent } from '../../ribbon-group/ppt-ribbon-group.component';
import { PptLayoutGalleryComponent } from '../../../menus/layout-gallery/ppt-layout-gallery.component';
import { SlideLayout } from '../../../../models/ppt.types';
import { PptRibbonFoldService } from '../../../../services/ppt-ribbon-fold.service';
import { PptActionLoggerService } from '../../../../services/ppt-action-logger.service';
import { computePopupPosition, PopupPosition } from '../../../../utils/popup-position';

@Component({
  selector: 'app-ppt-insert-tab',
  standalone: true,
  imports: [
    CommonModule,
    PptIconComponent,
    PptRibbonGroupComponent,
    PptLayoutGalleryComponent,
  ],
  templateUrl: './ppt-insert-tab.component.html',
  styleUrls: ['./ppt-insert-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptInsertTabComponent {
  @Input() currentLayout: SlideLayout = 'title';

  @Output() readonly newSlide = new EventEmitter<SlideLayout | void>();
  @Output() readonly insertTable = new EventEmitter<{ rows: number; cols: number }>();
  @Output() readonly openPictures = new EventEmitter<void>();
  @Output() readonly insertShape = new EventEmitter<string>();
  @Output() readonly insertTextBox = new EventEmitter<void>();

  isNewSlideGalleryOpen = false;
  isTablePickerOpen = false;
  isShapesGalleryOpen = false;
  hoverRows = 0;
  hoverCols = 0;
  tablePos: PopupPosition = { left: 0, top: 0 };
  shapesPos: PopupPosition = { left: 0, top: 0 };
  private lastTableAnchor: HTMLElement | null = null;
  private lastShapesAnchor: HTMLElement | null = null;

  constructor(
    public foldService: PptRibbonFoldService,
    private actionLogger: PptActionLoggerService,
    private elementRef: ElementRef
  ) {}

  
  get isCompact(): boolean {
    return this.foldService.isTabCompact('insert');
  }

  isOverflow(groupId: string): boolean {
    return this.foldService.isGroupInOverflow('insert', groupId);
  }

  isFolded(groupId: string): boolean {
    return this.foldService.isGroupFolded('insert', groupId);
  }

  isFoldOpen(groupId: string): boolean {
    return this.foldService.openFoldGroupId() === groupId;
  }

  onToggleFold(groupId: string): void {
    this.foldService.toggleFold(groupId);
  }

  onCloseFold(): void {
    this.foldService.closeFold();
  }

  onNewSlideBigClick(): void {
    this.newSlide.emit();
  }

  onToggleNewSlideGallery(event: MouseEvent): void {
    event.stopPropagation();
    this.isNewSlideGalleryOpen = !this.isNewSlideGalleryOpen;
    this.isTablePickerOpen = false;
    this.isShapesGalleryOpen = false;
    if (this.isNewSlideGalleryOpen) {
      this.actionLogger.log('openMenu', { menu: 'Slide' });
    }
  }

  onToggleTablePicker(event: MouseEvent): void {
    this.lastTableAnchor = event.currentTarget as HTMLElement;
    this.recalculatePositions();
    event.stopPropagation();
    this.isTablePickerOpen = !this.isTablePickerOpen;
    this.isNewSlideGalleryOpen = false;
    this.isShapesGalleryOpen = false;
  }

  onToggleShapesGallery(event: MouseEvent): void {
    this.lastShapesAnchor = event.currentTarget as HTMLElement;
    this.recalculatePositions();
    event.stopPropagation();
    this.isShapesGalleryOpen = !this.isShapesGalleryOpen;
    this.isNewSlideGalleryOpen = false;
    this.isTablePickerOpen = false;
  }

  onSelectShape(shape: string, event: MouseEvent): void {
    event.stopPropagation();
    this.isShapesGalleryOpen = false;
    this.insertShape.emit(shape);
  }

  onTableHover(r: number, c: number): void {
    this.hoverRows = r;
    this.hoverCols = c;
  }

  onSelectTableSize(r: number, c: number, event: MouseEvent): void {
    event.stopPropagation();
    this.isTablePickerOpen = false;
    this.insertTable.emit({ rows: r, cols: c });
  }

  @HostListener('document:click', ['$event'])
  onDocClick(event: MouseEvent): void {
    if (
      (this.isShapesGalleryOpen || this.isTablePickerOpen || this.isNewSlideGalleryOpen) &&
      !this.elementRef.nativeElement.contains(event.target)
    ) {
      this.isShapesGalleryOpen = false;
      this.isTablePickerOpen = false;
      this.isNewSlideGalleryOpen = false;
    }
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    this.recalculatePositions();
  }

  private recalculatePositions(): void {
    if (this.lastTableAnchor) {
      this.tablePos = computePopupPosition(this.lastTableAnchor, { width: 170, height: 190 });
    }
    if (this.lastShapesAnchor) {
      this.shapesPos = computePopupPosition(this.lastShapesAnchor, { width: 220, height: 260 });
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.isShapesGalleryOpen = false;
    this.isTablePickerOpen = false;
    this.isNewSlideGalleryOpen = false;
  }
}
