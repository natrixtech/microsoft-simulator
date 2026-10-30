import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptIconComponent } from '../../ppt-icon/ppt-icon.component';
import { PptRibbonGroupComponent } from '../../ribbon-group/ppt-ribbon-group.component';
import { PptLayoutGalleryComponent } from '../../../menus/layout-gallery/ppt-layout-gallery.component';
import { SlideLayout, ThemeDef, ThemeId } from '../../../../models/ppt.types';
import { PPT_THEMES } from '../../../../constants/ppt-themes';
import { PptRibbonFoldService } from '../../../../services/ppt-ribbon-fold.service';
import { computePopupPosition, PopupPosition } from '../../../../utils/popup-position';
import { HostListener, ElementRef } from '@angular/core';

@Component({
  selector: 'app-ppt-design-tab',
  standalone: true,
  imports: [
    CommonModule,
    PptIconComponent,
    PptRibbonGroupComponent,
    PptLayoutGalleryComponent,
  ],
  templateUrl: './ppt-design-tab.component.html',
  styleUrls: ['./ppt-design-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptDesignTabComponent {
  @Input() currentTheme: ThemeId = 'simple-light';
  @Input() currentLayout: SlideLayout = 'title';

  @Output() readonly selectTheme = new EventEmitter<ThemeId>();
  @Output() readonly changeLayout = new EventEmitter<SlideLayout>();
  @Output() readonly formatBackground = new EventEmitter<void>();
  @Output() readonly resetBackground = new EventEmitter<void>();

  readonly themeList: ThemeDef[] = Object.values(PPT_THEMES);
  isThemesDropdownOpen = false;
  isBackgroundMenuOpen = false;
  isLayoutGalleryOpen = false;
  themesPos: PopupPosition = { left: 0, top: 0 };
  bgMenuPos: PopupPosition = { left: 0, top: 0 };
  private lastThemesAnchor: HTMLElement | null = null;
  private lastBgAnchor: HTMLElement | null = null;

  constructor(public foldService: PptRibbonFoldService) {}

  
  get isCompact(): boolean {
    return this.foldService.isTabCompact('design');
  }

  isOverflow(groupId: string): boolean {
    return this.foldService.isGroupInOverflow('design', groupId);
  }

  isFolded(groupId: string): boolean {
    return this.foldService.isGroupFolded('design', groupId);
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

  onPickTheme(themeId: ThemeId): void {
    this.isThemesDropdownOpen = false;
    this.selectTheme.emit(themeId);
  }

  onToggleThemesDropdown(event: MouseEvent): void {
    const btn = event.currentTarget as HTMLElement;
    this.lastThemesAnchor = (btn.closest('.themes-strip-container') as HTMLElement) || btn;
    this.recalculatePositions();
    event.stopPropagation();
    this.isThemesDropdownOpen = !this.isThemesDropdownOpen;
    this.isBackgroundMenuOpen = false;
    this.isLayoutGalleryOpen = false;
  }

  onToggleBackgroundMenu(event: MouseEvent): void {
    this.lastBgAnchor = event.currentTarget as HTMLElement;
    this.recalculatePositions();
    event.stopPropagation();
    this.isBackgroundMenuOpen = !this.isBackgroundMenuOpen;
    this.isThemesDropdownOpen = false;
    this.isLayoutGalleryOpen = false;
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    this.recalculatePositions();
  }

  @HostListener('document:click', ['$event'])
  onDocClick(event: MouseEvent): void {
    if (this.isThemesDropdownOpen || this.isBackgroundMenuOpen) {
      this.isThemesDropdownOpen = false;
      this.isBackgroundMenuOpen = false;
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.isThemesDropdownOpen = false;
    this.isBackgroundMenuOpen = false;
  }

  private recalculatePositions(): void {
    if (this.lastThemesAnchor) {
      const g = this.lastThemesAnchor.getBoundingClientRect();
      const ribbonBody = this.lastThemesAnchor.closest('.ppt-ribbon-body') as HTMLElement | null;
      const bottom = ribbonBody ? ribbonBody.getBoundingClientRect().bottom : g.bottom;
      this.themesPos = computePopupPosition(new DOMRect(g.left, bottom, g.width, 0), { width: 448, height: 200 });
    }
    if (this.lastBgAnchor) {
      this.bgMenuPos = computePopupPosition(this.lastBgAnchor, { width: 180, height: 90 });
    }
  }

  onToggleLayoutGallery(event: MouseEvent): void {
    event.stopPropagation();
    this.isLayoutGalleryOpen = !this.isLayoutGalleryOpen;
    this.isThemesDropdownOpen = false;
    this.isBackgroundMenuOpen = false;
  }
}

