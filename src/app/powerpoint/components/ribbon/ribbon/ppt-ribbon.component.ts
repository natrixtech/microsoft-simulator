import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnDestroy,
  Output,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { PptIconComponent } from '../ppt-icon/ppt-icon.component';
import { PptHomeTabComponent } from '../tabs/home-tab/ppt-home-tab.component';
import { PptInsertTabComponent } from '../tabs/insert-tab/ppt-insert-tab.component';
import { PptDesignTabComponent } from '../tabs/design-tab/ppt-design-tab.component';
import { PptTransitionsTabComponent } from '../tabs/transitions-tab/ppt-transitions-tab.component';
import {
  RibbonTabId,
  SlideLayout,
  SlideTransition,
  ThemeId,
} from '../../../models/ppt.types';
import {
  RIBBON_TABS,
  RibbonTabConfig,
  SEARCH_COMMANDS,
  SearchCommand,
} from '../../../constants/ppt-ribbon';
import { PptRibbonFoldService } from '../../../services/ppt-ribbon-fold.service';
import { PptActionLoggerService } from '../../../services/ppt-action-logger.service';
import { computePopupPosition, PopupPosition } from '../../../utils/popup-position';

@Component({
  selector: 'app-ppt-ribbon',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    PptIconComponent,
    PptHomeTabComponent,
    PptInsertTabComponent,
    PptDesignTabComponent,
    PptTransitionsTabComponent,
  ],
  templateUrl: './ppt-ribbon.component.html',
  styleUrls: ['./ppt-ribbon.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptRibbonComponent implements AfterViewInit, OnDestroy {
  @ViewChild('tabStrip', { static: false }) tabStripRef?: ElementRef<HTMLElement>;
  @ViewChild('tabStripRight', { static: false }) tabStripRightRef?: ElementRef<HTMLElement>;
  @ViewChild('tabOverflowBtn', { static: false }) tabOverflowBtnRef?: ElementRef<HTMLElement>;
  @ViewChild('ribbonOverflowBtn', { static: false }) ribbonOverflowBtnRef?: ElementRef<HTMLElement>;
  @ViewChild('ribbonBody', { static: false }) ribbonBodyRef?: ElementRef<HTMLElement>;
  @ViewChild('searchInput', { static: false }) searchInputRef?: ElementRef<HTMLElement>;

  @Input() activeTab: RibbonTabId = 'home';
  @Input() canUndo: boolean = false;
  @Input() canRedo: boolean = false;
  @Input() canPaste: boolean = false;
  @Input() currentLayout: SlideLayout = 'title';
  @Input() currentTheme: ThemeId = 'simple-light';
  @Input() currentTransition: SlideTransition = { type: 'none' };
  @Input() isTextActive: boolean = false;
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

  @Output() readonly tabChange = new EventEmitter<RibbonTabId>();
  @Output() readonly undo = new EventEmitter<void>();
  @Output() readonly redo = new EventEmitter<void>();
  @Output() readonly paste = new EventEmitter<void>();
  @Output() readonly copy = new EventEmitter<void>();
  @Output() readonly newSlide = new EventEmitter<SlideLayout | void>();
  @Output() readonly changeLayout = new EventEmitter<SlideLayout>();
  @Output() readonly selectTheme = new EventEmitter<ThemeId>();
  @Output() readonly selectTransition = new EventEmitter<SlideTransition>();
  @Output() readonly previewTransition = new EventEmitter<void>();
  @Output() readonly formatBackground = new EventEmitter<void>();
  @Output() readonly resetBackground = new EventEmitter<void>();

  // Text formatting
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
  @Output() readonly insertTable = new EventEmitter<{ rows: number; cols: number }>();
  @Output() readonly openPictures = new EventEmitter<void>();
  @Output() readonly insertShape = new EventEmitter<string>();
  @Output() readonly insertTextBox = new EventEmitter<void>();
  @Output() readonly selectAll = new EventEmitter<void>();
  @Output() readonly selectObjects = new EventEmitter<void>();

  readonly tabs: RibbonTabConfig[] = RIBBON_TABS;

  // Tab strip overflow state
  visibleTabs: RibbonTabConfig[] = [...RIBBON_TABS];
  hiddenTabs: RibbonTabConfig[] = [];
  isTabOverflowOpen = false;
  tabOverflowPos: PopupPosition = { left: 0, top: 0 };
  private tabStripObserver?: ResizeObserver;
  private tabWidths: number[] = [];

  // Stage 3 Ribbon overflow state
  isRibbonOverflowOpen = false;
  ribbonOverflowPos: PopupPosition = { left: 0, top: 0 };


  // Search feature
  searchQuery = '';
  isSearchDropdownOpen = false;
  searchPos: PopupPosition = { left: 0, top: 0 };
  filteredCommands: SearchCommand[] = [];

  constructor(
    public foldService: PptRibbonFoldService,
    private actionLogger: PptActionLoggerService,
    private elementRef: ElementRef,
    private cdr: ChangeDetectorRef
  ) {}

  ngAfterViewInit(): void {
    if (this.ribbonBodyRef?.nativeElement) {
      this.foldService.initObserver(this.ribbonBodyRef.nativeElement, () => this.activeTab);
    }
    if (this.tabStripRef?.nativeElement) {
      this.tabStripObserver = new ResizeObserver(() => {
        this.updateTabOverflow();
      });
      this.tabStripObserver.observe(this.tabStripRef.nativeElement);
      // Run initial check
      setTimeout(() => this.updateTabOverflow(), 0);
    }
  }

  ngOnDestroy(): void {
    this.foldService.destroyObserver();
    if (this.tabStripObserver) {
      this.tabStripObserver.disconnect();
    }
  }

  
  updateTabOverflow(): void {
    const stripEl = this.tabStripRef?.nativeElement;
    const rightEl = this.tabStripRightRef?.nativeElement;
    if (!stripEl || !rightEl) return;

    const available = stripEl.clientWidth - rightEl.offsetWidth - 24;
    if (available <= 0) return;

    // Standard button widths with 0 10px padding and 2px gap
    const fallbackWidths = [45, 58, 58, 52, 64, 88, 88, 86, 64, 64, 52, 50];
    const widths = this.tabWidths.length === this.tabs.length ? this.tabWidths : fallbackWidths;

    let totalWidth = 0;
    for (const w of widths) totalWidth += w;

    if (totalWidth <= available) {
      this.visibleTabs = [...this.tabs];
      this.hiddenTabs = [];
      this.isTabOverflowOpen = false;
      this.cdr.markForCheck();
      return;
    }

    const maxTabsWidth = available - 32;
    let currentWidth = 0;
    let split = 0;
    for (let i = 0; i < widths.length; i++) {
      if (currentWidth + widths[i] <= maxTabsWidth) {
        currentWidth += widths[i];
        split = i + 1;
      } else {
        break;
      }
    }
    split = Math.max(1, split);
    this.visibleTabs = this.tabs.slice(0, split);
    this.hiddenTabs = this.tabs.slice(split);
    this.cdr.markForCheck();
  }

  onToggleTabOverflow(event: MouseEvent): void {
    event.stopPropagation();
    this.isTabOverflowOpen = !this.isTabOverflowOpen;
    if (this.isTabOverflowOpen) {
      const btn = event.currentTarget as HTMLElement;
      this.tabOverflowPos = computePopupPosition(btn, { width: 140, height: 200 });
      this.isSearchDropdownOpen = false;
      this.isRibbonOverflowOpen = false;
    }
  }

  onSelectHiddenTab(tab: RibbonTabConfig): void {
    if (!tab.enabled) return;
    this.isTabOverflowOpen = false;
    this.onSelectTab(tab);
  }

  get currentOverflowGroups(): string[] {
    return this.foldService.overflowGroups()[this.activeTab] || [];
  }

  getGroupTitle(groupId: string): string {
    const titles: Record<string, string> = {
      designer: 'Designer',
      addins: 'Add-ins',
      voice: 'Voice',
      editing: 'Editing',
      camera: 'Camera',
      zoom: 'Zoom',
      links: 'Links',
      comments: 'Comments',
      symbols: 'Symbols',
      media: 'Media',
      slides: 'Slides',
      customize: 'Customize',
      edit: 'Edit',
      timing: 'Timing',
    };
    return titles[groupId] || groupId;
  }

  onToggleRibbonOverflow(event: MouseEvent): void {
    event.stopPropagation();
    this.isRibbonOverflowOpen = !this.isRibbonOverflowOpen;
    if (this.isRibbonOverflowOpen) {
      const btn = event.currentTarget as HTMLElement;
      this.ribbonOverflowPos = computePopupPosition(btn, { width: 220, height: 180 });
      this.isSearchDropdownOpen = false;
      this.isTabOverflowOpen = false;
    }
  }

  onRibbonOverflowClick(event: MouseEvent): void {
    event.stopPropagation();
    const target = event.target as HTMLElement | null;
    if (!target) return;
    const btn = target.closest('button') as HTMLButtonElement | null;
    if (btn && !btn.disabled) {
      this.isRibbonOverflowOpen = false;
    }
  }

  onSelectTab(tab: RibbonTabConfig): void {
    if (!tab.enabled) return;
    this.tabChange.emit(tab.id);
    if (tab.id === 'insert') {
      this.actionLogger.log('openMenu', { menu: 'Insert' });
    }
    this.foldService.setCurrentTab(tab.id);
  }

  onSearchFocus(event: FocusEvent): void {
    this.updateSearchFilter();
    this.recalculateSearchPos(event.target as HTMLElement);
    this.isSearchDropdownOpen = true;
  }

  onSearchInput(): void {
    this.updateSearchFilter();
    this.isSearchDropdownOpen = true;
  }

  private recalculateSearchPos(inputEl?: HTMLElement): void {
    const el = inputEl || this.searchInputRef?.nativeElement;
    if (!el) return;
    this.searchPos = computePopupPosition(el, { width: 220, height: 240 });
  }

  private updateSearchFilter(): void {
    const q = this.searchQuery.trim().toLowerCase();
    if (!q) {
      this.filteredCommands = SEARCH_COMMANDS.slice(0, 8);
    } else {
      this.filteredCommands = SEARCH_COMMANDS.filter(
        (c) => c.label.toLowerCase().includes(q) || c.group.toLowerCase().includes(q)
      );
    }
  }

  executeSearchCommand(cmd: SearchCommand): void {
    this.isSearchDropdownOpen = false;
    this.searchQuery = '';
    if (this.activeTab !== cmd.tab) {
      this.tabChange.emit(cmd.tab);
      this.foldService.setCurrentTab(cmd.tab);
    }

    switch (cmd.id) {
      case 'undo': this.undo.emit(); break;
      case 'redo': this.redo.emit(); break;
      case 'new-slide': this.newSlide.emit(); break;
      case 'bold': this.toggleBold.emit(); break;
      case 'italic': this.toggleItalic.emit(); break;
      case 'underline': this.toggleUnderline.emit(); break;
      case 'bullets': this.setBullets.emit(); break;
      case 'numbering': this.setNumbering.emit(); break;
      case 'align-left': this.setAlign.emit('left'); break;
      case 'align-center': this.setAlign.emit('center'); break;
      case 'align-right': this.setAlign.emit('right'); break;
      case 'insert-table': this.insertTable.emit({ rows: 3, cols: 3 }); break;
      case 'insert-picture': this.openPictures.emit(); break;
      case 'insert-shape': this.insertShape.emit('rectangle'); break;
      case 'insert-textbox': this.insertTextBox.emit(); break;
      case 'theme-gallery': this.tabChange.emit('design'); this.foldService.setCurrentTab('design'); break;
      case 'background-styles': this.formatBackground.emit(); break;
      case 'transition-fade': this.selectTransition.emit({ type: 'fade' }); break;
      case 'transition-push': this.selectTransition.emit({ type: 'push', direction: 'fromRight' }); break;
    }
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (this.isSearchDropdownOpen && !this.elementRef.nativeElement.contains(target)) {
      this.isSearchDropdownOpen = false;
    }
    if (this.isTabOverflowOpen && !target.closest('.ppt-tab-overflow-btn') && !target.closest('.ppt-tab-overflow-menu')) {
      this.isTabOverflowOpen = false;
    }
    if (this.isRibbonOverflowOpen && !target.closest('.ppt-ribbon-more-btn') && !target.closest('.ppt-ribbon-overflow-panel')) {
      this.isRibbonOverflowOpen = false;
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.isSearchDropdownOpen = false;
    this.isTabOverflowOpen = false;
    this.isRibbonOverflowOpen = false;
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (this.isSearchDropdownOpen) {
      this.recalculateSearchPos();
    }
    if (this.isTabOverflowOpen) {
      this.isTabOverflowOpen = false;
    }
    if (this.isRibbonOverflowOpen) {
      this.isRibbonOverflowOpen = false;
    }
    this.updateTabOverflow();
  }
}
