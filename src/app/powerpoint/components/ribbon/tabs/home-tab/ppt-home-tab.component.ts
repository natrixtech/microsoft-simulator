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
import { PptColorPickerComponent } from '../../../menus/color-picker/ppt-color-picker.component';
import { SlideLayout } from '../../../../models/ppt.types';
import { PPT_FONTS, PPT_FONT_SIZES } from '../../../../constants/ppt-fonts';
import { PptRibbonFoldService } from '../../../../services/ppt-ribbon-fold.service';
import { PptActionLoggerService } from '../../../../services/ppt-action-logger.service';
import { computePopupPosition, PopupPosition } from '../../../../utils/popup-position';

@Component({
  selector: 'app-ppt-home-tab',
  standalone: true,
  imports: [
    CommonModule,
    PptIconComponent,
    PptRibbonGroupComponent,
    PptLayoutGalleryComponent,
    PptColorPickerComponent,
  ],
  templateUrl: './ppt-home-tab.component.html',
  styleUrls: ['./ppt-home-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptHomeTabComponent {
  @Input() canUndo: boolean = false;
  @Input() canRedo: boolean = false;
  @Input() canPaste: boolean = false;
  @Input() currentLayout: SlideLayout = 'title';
  @Input() isTextActive: boolean = false;

  // Active format state
  @Input() isBold: boolean = false;
  @Input() isItalic: boolean = false;
  @Input() isUnderline: boolean = false;
  @Input() currentFont: string = 'Arial';
  @Input() currentFontSize: number = 24;
  @Input() currentAlign: 'left' | 'center' | 'right' = 'left';
  @Input() currentLineSpacing: number = 1.0;
  @Input() lastFontColor: string = '#c43e1c';
  @Input() hasBullets: boolean = false;
  @Input() hasNumbering: boolean = false;

  @Output() readonly undo = new EventEmitter<void>();
  @Output() readonly redo = new EventEmitter<void>();
  @Output() readonly paste = new EventEmitter<void>();
  @Output() readonly copy = new EventEmitter<void>();
  @Output() readonly newSlide = new EventEmitter<SlideLayout | void>();
  @Output() readonly changeLayout = new EventEmitter<SlideLayout>();

  // Formatting events
  @Output() readonly toggleBold = new EventEmitter<void>();
  @Output() readonly toggleItalic = new EventEmitter<void>();
  @Output() readonly toggleUnderline = new EventEmitter<void>();
  @Output() readonly changeFont = new EventEmitter<string>();
  @Output() readonly changeFontSize = new EventEmitter<number>();
  @Output() readonly increaseFontSize = new EventEmitter<void>();
  @Output() readonly decreaseFontSize = new EventEmitter<void>();
  @Output() readonly setFontColor = new EventEmitter<string>();
  @Output() readonly setBullets = new EventEmitter<void>();
  @Output() readonly setNumbering = new EventEmitter<void>();
  @Output() readonly setAlign = new EventEmitter<'left' | 'center' | 'right'>();
  @Output() readonly setLineSpacing = new EventEmitter<number>();
  @Output() readonly insertShape = new EventEmitter<string>();
  @Output() readonly selectAll = new EventEmitter<void>();
  @Output() readonly selectObjects = new EventEmitter<void>();

  isNewSlideGalleryOpen = false;
  isLayoutGalleryOpen = false;

  // Color picker state
  isColorPickerOpen = false;
  colorPickerPos: PopupPosition = { left: 0, top: 0 };

  // Line spacing dropdown state
  isLineSpacingOpen = false;
  lineSpacingPos: PopupPosition = { left: 0, top: 0 };
  readonly lineSpacingOptions: number[] = [1.0, 1.15, 1.5, 2.0, 2.5, 3.0];

  // Select dropdown state
  isSelectMenuOpen = false;
  selectMenuPos: PopupPosition = { left: 0, top: 0 };

  readonly fonts = PPT_FONTS;
  readonly fontSizes = PPT_FONT_SIZES;

  constructor(
    public foldService: PptRibbonFoldService,
    private actionLogger: PptActionLoggerService,
    private elementRef: ElementRef
  ) {}

  @HostListener('document:click', ['$event'])
  onDocClick(event: MouseEvent): void {
    const target = event.target as HTMLElement | null;
    if (!target?.closest('.ppt-select-dropdown') && !target?.closest('[data-demo-target="home-select"]')) {
      this.isSelectMenuOpen = false;
    }
    if (!target?.closest('.ppt-spacing-dropdown') && !target?.closest('[data-demo-target="home-line-spacing"]')) {
      this.isLineSpacingOpen = false;
    }
    if (!target?.closest('app-ppt-color-picker') && !target?.closest('[data-demo-target="home-font-color-arrow"]')) {
      this.isColorPickerOpen = false;
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.isColorPickerOpen = false;
    this.isLineSpacingOpen = false;
    this.isSelectMenuOpen = false;
  }

  
  get isCompact(): boolean {
    return this.foldService.isTabCompact('home');
  }

  isOverflow(groupId: string): boolean {
    return this.foldService.isGroupInOverflow('home', groupId);
  }

  isFolded(groupId: string): boolean {
    return this.foldService.isGroupFolded('home', groupId);
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
    this.isLayoutGalleryOpen = false;
    this.isColorPickerOpen = false;
    this.isLineSpacingOpen = false;
    this.isSelectMenuOpen = false;
    if (this.isNewSlideGalleryOpen) {
      this.actionLogger.log('openMenu', { menu: 'Slide' });
    }
  }

  onToggleLayoutGallery(event: MouseEvent): void {
    event.stopPropagation();
    this.isLayoutGalleryOpen = !this.isLayoutGalleryOpen;
    this.isNewSlideGalleryOpen = false;
    this.isColorPickerOpen = false;
    this.isLineSpacingOpen = false;
    this.isSelectMenuOpen = false;
    if (this.isLayoutGalleryOpen) {
      this.actionLogger.log('openMenu', { menu: 'Slide' });
    }
  }

  onPickNewSlideLayout(layout: SlideLayout): void {
    this.isNewSlideGalleryOpen = false;
    this.newSlide.emit(layout);
  }

  onPickChangeLayout(layout: SlideLayout): void {
    this.isLayoutGalleryOpen = false;
    this.changeLayout.emit(layout);
  }

  onToggleColorPicker(event: MouseEvent): void {
    event.stopPropagation();
    const btn = event.currentTarget as HTMLElement;
    this.colorPickerPos = computePopupPosition(btn, { width: 176, height: 210 });
    this.isColorPickerOpen = !this.isColorPickerOpen;
    this.isLineSpacingOpen = false;
    this.isSelectMenuOpen = false;
  }

  onPickColor(color: string): void {
    this.isColorPickerOpen = false;
    this.setFontColor.emit(color);
  }

  onToggleLineSpacing(event: MouseEvent): void {
    event.stopPropagation();
    const btn = event.currentTarget as HTMLElement;
    this.lineSpacingPos = computePopupPosition(btn, { width: 100, height: 160 });
    this.isLineSpacingOpen = !this.isLineSpacingOpen;
    this.isColorPickerOpen = false;
    this.isSelectMenuOpen = false;
  }

  onPickLineSpacing(spacing: number): void {
    this.isLineSpacingOpen = false;
    this.setLineSpacing.emit(spacing);
  }

  onToggleSelectMenu(event: MouseEvent): void {
    event.stopPropagation();
    const btn = event.currentTarget as HTMLElement;
    this.selectMenuPos = computePopupPosition(btn, { width: 170, height: 110 });
    this.isSelectMenuOpen = !this.isSelectMenuOpen;
    this.isColorPickerOpen = false;
    this.isLineSpacingOpen = false;
    this.isNewSlideGalleryOpen = false;
    this.isLayoutGalleryOpen = false;
  }

  onPickSelectAll(): void {
    this.isSelectMenuOpen = false;
    this.selectAll.emit();
  }

  onPickSelectObjects(): void {
    this.isSelectMenuOpen = false;
    this.selectObjects.emit();
  }
}
