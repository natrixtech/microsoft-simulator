import {
  AfterViewInit,
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  HostListener,
  OnDestroy,
  OnInit,
  ViewChild,
  computed,
  effect,
  signal,
  untracked,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  ActivityPayload,
  COLORS,
  ConfettiService,
  FONTS,
  LessonTask,
  PluginBridgeService,
  ShowMeHowComponent,
  TaskCardComponent,
  ThemeService,
  clone,
  sanitizeHtml,
} from 'shared';
import { PageSetup, WordComment, WordDocumentService } from './services/word-document.service';
import { WORD_LESSONS } from './word-lessons.data';

export interface TaskStartSnapshot {
  documentHtml: string;
  pageSetup: PageSetup;
  comments: WordComment[];
  header: string;
  footer: string;
  numbering: boolean;
  variables: Record<string, string>;
  images: any;
  assetBaseUrl: string;
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, TaskCardComponent, ShowMeHowComponent],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class AppComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('editorElement') editorRef?: ElementRef<HTMLElement>;
  @ViewChild('page2EditorElement') page2EditorRef?: ElementRef<HTMLElement>;
  @ViewChild('headerEl') headerRef?: ElementRef<HTMLElement>;
  @ViewChild('footerEl') footerRef?: ElementRef<HTMLElement>;
  @ViewChild('workLayout') workLayoutRef?: ElementRef<HTMLElement>;
  @ViewChild('wordPaper') paperRef?: ElementRef<HTMLElement>;
  @ViewChild('movableGroup') movableGroupRef?: ElementRef<HTMLElement>;

  readonly hasPageBreak = signal<boolean>(false);
  readonly activeRegion = signal<'body' | 'header' | 'footer'>('body');
  readonly activePageNum = signal<number>(1);
  readonly totalDocPages = computed<number>(() => this.hasPageBreak() ? 2 : 1);

  // Draggable Image Resizing State
  readonly imageBoxLeft = signal<number>(0);
  readonly imageBoxTop = signal<number>(0);
  readonly imageBoxWidth = signal<number>(0);
  readonly imageBoxHeight = signal<number>(0);
  private isResizing = false;
  private resizeHandle: string = '';
  private resizeStartX = 0;
  private resizeStartY = 0;
  private resizeStartWidth = 0;
  private resizeStartHeight = 0;
  private resizeAspectRatio = 1;

  readonly currentActivity = signal<ActivityPayload | null>(null);
  readonly taskIndex = signal<number>(0);
  readonly passedTasks = signal<Set<number>>(new Set());
  readonly isPassed = signal<boolean>(false);
  readonly isCompleted = signal<boolean>(false);
  readonly feedbackMsg = signal<string>('');
  readonly isDemoOpen = signal<boolean>(false);
  readonly currentTaskSnapshot = signal<TaskStartSnapshot | null>(null);
  private readonly taskStartSnapshots = new Map<number, TaskStartSnapshot>();
  readonly showFinalTealCard = signal<boolean>(false);
  readonly wordCount = signal<number>(0);

  // Movable Group Positioning State (initially beside right of page)
  readonly groupX = signal<number | null>(null);
  readonly groupY = signal<number | null>(null);
  readonly hasUserMoved = signal<boolean>(false);
  private isDragging = false;
  private dragStartPointerX = 0;
  private dragStartPointerY = 0;
  private dragStartGroupX = 0;
  private dragStartGroupY = 0;

  readonly showPageSetup = signal<boolean>(false);
  readonly showFindReplace = signal<boolean>(false);
  readonly showCommentDialog = signal<boolean>(false);
  readonly commentSelectedText = signal<string>('');
  readonly newCommentText = signal<string>('Please review this section.');
  readonly activeRibbonTab = signal<'Home' | 'Insert' | 'Layout'>('Home');
  readonly showOrientationPicker = signal<boolean>(false);
  readonly openMenuName = signal<string | null>(null);

  readonly selectedStyle = signal<string>('normal');
  readonly selectedFont = signal<string>('Calibri');
  readonly selectedFontSize = signal<number>(11);

  readonly standaloneLessons = WORD_LESSONS;
  readonly standaloneLessonIndex = signal<number>(0);

  private advanceTimer: any = null;

  readonly isAssessment = computed<boolean>(() => {
    return this.currentActivity()?.mode === 'assessment';
  });

  readonly totalTasks = computed<number>(() => {
    return this.currentActivity()?.tasks?.length || 1;
  });

  readonly currentTask = computed<LessonTask | null>(() => {
    const act = this.currentActivity();
    if (!act || !act.tasks) return null;
    return act.tasks[this.taskIndex()] || null;
  });

  readonly isDarkTheme = computed<boolean>(() => {
    const t = (this.theme && typeof this.theme.currentTheme === 'function' ? this.theme.currentTheme() : null) || this.currentActivity()?.theme || 'LIGHT';
    return t === 'DARK';
  });

  findTextQuery: string = '';
  replaceTextQuery: string = '';
  readonly matchCase = signal<boolean>(false);
  readonly matchCount = signal<number>(0);

  readonly showImageDialog = signal<boolean>(false);
  readonly selectedImageFile = signal<string>('');
  readonly selectedImageEl = signal<HTMLImageElement | null>(null);
  readonly selectedTableEl = signal<HTMLTableElement | null>(null);
  readonly selectedTocEl = signal<HTMLElement | null>(null);
  readonly selectedPageNum = signal<boolean>(false);

  readonly showTablePicker = signal<boolean>(false);
  readonly tableHoverRows = signal<number>(2);
  readonly tableHoverCols = signal<number>(2);
  readonly gridRows = [1, 2, 3, 4, 5, 6];
  readonly gridCols = [1, 2, 3, 4, 5, 6, 7, 8];

  readonly showFontColorPicker = signal<boolean>(false);
  readonly showHighlightColorPicker = signal<boolean>(false);
  readonly activeFontColor = signal<string>('#000000');
  readonly activeHighlightColor = signal<string>('#ffff00');

  readonly schemeBaseColors = [
    { name: 'White, Background 1', hex: '#ffffff', key: 'white' },
    { name: 'Black, Text 1', hex: '#000000', key: 'black' },
    { name: 'Gray 25%', hex: '#e7e6e6', key: 'lightGray' },
    { name: 'Dark Blue, Text 2', hex: '#44546a', key: 'darkBlue' },
    { name: 'Blue, Accent 1', hex: '#4472c4', key: 'blue' },
    { name: 'Orange, Accent 2', hex: '#ed7d31', key: 'orange' },
    { name: 'Gray, Accent 3', hex: '#a5a5a5', key: 'gray' },
    { name: 'Gold, Accent 4', hex: '#ffc000', key: 'yellow' },
    { name: 'Blue, Accent 5', hex: '#5b9bd5', key: 'cyan' },
    { name: 'Green, Accent 6', hex: '#70ad47', key: 'green' },
  ];

  readonly schemeColumns = [
    {
      shades: [
        { name: 'White, Darker 5%', hex: '#f2f2f2', key: 'white' },
        { name: 'White, Darker 15%', hex: '#d9d9d9', key: 'lightGray' },
        { name: 'White, Darker 25%', hex: '#bfbfbf', key: 'gray' },
        { name: 'White, Darker 35%', hex: '#a6a6a6', key: 'gray' },
        { name: 'White, Darker 50%', hex: '#7f7f7f', key: 'darkGray' },
      ],
    },
    {
      shades: [
        { name: 'Black, Lighter 50%', hex: '#7f7f7f', key: 'darkGray' },
        { name: 'Black, Lighter 35%', hex: '#595959', key: 'darkGray' },
        { name: 'Black, Lighter 25%', hex: '#3f3f3f', key: 'darkGray' },
        { name: 'Black, Lighter 15%', hex: '#262626', key: 'black' },
        { name: 'Black, Lighter 5%', hex: '#0c0c0c', key: 'black' },
      ],
    },
    {
      shades: [
        { name: 'Gray, Lighter 80%', hex: '#f2f2f2', key: 'lightGray' },
        { name: 'Gray, Lighter 60%', hex: '#d9d9d9', key: 'lightGray' },
        { name: 'Gray, Lighter 40%', hex: '#bfbfbf', key: 'gray' },
        { name: 'Gray, Darker 25%', hex: '#7f7f7f', key: 'gray' },
        { name: 'Gray, Darker 50%', hex: '#595959', key: 'darkGray' },
      ],
    },
    {
      shades: [
        { name: 'Dark Blue, Lighter 80%', hex: '#d6dce4', key: 'lightBlue' },
        { name: 'Dark Blue, Lighter 60%', hex: '#acb9ca', key: 'lightBlue' },
        { name: 'Dark Blue, Lighter 40%', hex: '#8497b0', key: 'blue' },
        { name: 'Dark Blue, Darker 25%', hex: '#2f3b4c', key: 'darkBlue' },
        { name: 'Dark Blue, Darker 50%', hex: '#1f2733', key: 'darkBlue' },
      ],
    },
    {
      shades: [
        { name: 'Blue, Lighter 80%', hex: '#d9e1f2', key: 'lightBlue' },
        { name: 'Blue, Lighter 60%', hex: '#b4c6e7', key: 'lightBlue' },
        { name: 'Blue, Lighter 40%', hex: '#8ea9db', key: 'blue' },
        { name: 'Blue, Darker 25%', hex: '#2f5597', key: 'blue' },
        { name: 'Blue, Darker 50%', hex: '#1f3864', key: 'darkBlue' },
      ],
    },
    {
      shades: [
        { name: 'Orange, Lighter 80%', hex: '#fce4d6', key: 'lightOrange' },
        { name: 'Orange, Lighter 60%', hex: '#f8cbad', key: 'lightOrange' },
        { name: 'Orange, Lighter 40%', hex: '#f4b084', key: 'orange' },
        { name: 'Orange, Darker 25%', hex: '#c65911', key: 'orange' },
        { name: 'Orange, Darker 50%', hex: '#833c0c', key: 'orange' },
      ],
    },
    {
      shades: [
        { name: 'Gray, Lighter 80%', hex: '#ededed', key: 'lightGray' },
        { name: 'Gray, Lighter 60%', hex: '#dbdbdb', key: 'lightGray' },
        { name: 'Gray, Lighter 40%', hex: '#c4c4c4', key: 'gray' },
        { name: 'Gray, Darker 25%', hex: '#7b7b7b', key: 'gray' },
        { name: 'Gray, Darker 50%', hex: '#525252', key: 'darkGray' },
      ],
    },
    {
      shades: [
        { name: 'Gold, Lighter 80%', hex: '#fff2cc', key: 'lightYellow' },
        { name: 'Gold, Lighter 60%', hex: '#ffe699', key: 'lightYellow' },
        { name: 'Gold, Lighter 40%', hex: '#ffd966', key: 'yellow' },
        { name: 'Gold, Darker 25%', hex: '#bf8f00', key: 'yellow' },
        { name: 'Gold, Darker 50%', hex: '#7f6000', key: 'yellow' },
      ],
    },
    {
      shades: [
        { name: 'Accent Blue, Lighter 80%', hex: '#deebf7', key: 'lightCyan' },
        { name: 'Accent Blue, Lighter 60%', hex: '#bdd7ee', key: 'lightBlue' },
        { name: 'Accent Blue, Lighter 40%', hex: '#9dc3e6', key: 'cyan' },
        { name: 'Accent Blue, Darker 25%', hex: '#2e75b6', key: 'blue' },
        { name: 'Accent Blue, Darker 50%', hex: '#1e4e79', key: 'darkBlue' },
      ],
    },
    {
      shades: [
        { name: 'Green, Lighter 80%', hex: '#e2efda', key: 'lightGreen' },
        { name: 'Green, Lighter 60%', hex: '#c6e0b4', key: 'lightGreen' },
        { name: 'Green, Lighter 40%', hex: '#a9d08e', key: 'green' },
        { name: 'Green, Darker 25%', hex: '#548235', key: 'green' },
        { name: 'Green, Darker 50%', hex: '#375623', key: 'green' },
      ],
    },
  ];

  readonly standardColors = [
    { name: 'Dark Red', hex: '#c00000', key: 'darkRed' },
    { name: 'Red', hex: '#cc0000', key: 'red' },
    { name: 'Orange', hex: '#e69138', key: 'orange' },
    { name: 'Yellow', hex: '#f1c232', key: 'yellow' },
    { name: 'Light Green', hex: '#d9ead3', key: 'lightGreen' },
    { name: 'Green', hex: '#6aa84f', key: 'green' },
    { name: 'Light Blue', hex: '#cfe2f3', key: 'lightBlue' },
    { name: 'Blue', hex: '#3d85c6', key: 'blue' },
    { name: 'Dark Blue', hex: '#185abd', key: 'darkBlue' },
    { name: 'Purple', hex: '#674ea7', key: 'purple' },
  ];

  readonly highlightColors = [
    { name: 'Yellow', hex: '#ffff00', key: 'yellow' },
    { name: 'Bright Green', hex: '#00ff00', key: 'green' },
    { name: 'Turquoise', hex: '#00ffff', key: 'cyan' },
    { name: 'Pink', hex: '#ff00ff', key: 'magenta' },
    { name: 'Blue', hex: '#0000ff', key: 'blue' },
    { name: 'Red', hex: '#ff0000', key: 'red' },
    { name: 'Dark Blue', hex: '#000080', key: 'darkBlue' },
    { name: 'Teal', hex: '#008080', key: 'cyan' },
    { name: 'Green', hex: '#008000', key: 'green' },
    { name: 'Violet', hex: '#800080', key: 'purple' },
    { name: 'Dark Red', hex: '#800000', key: 'darkRed' },
    { name: 'Dark Yellow', hex: '#808000', key: 'yellow' },
    { name: '50% Gray', hex: '#808080', key: 'gray' },
    { name: '25% Gray', hex: '#c0c0c0', key: 'gray' },
    { name: 'Black', hex: '#000000', key: 'black' },
  ];

  tempSetup: PageSetup = {
    orientation: 'portrait',
    paperSize: 'Letter',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
  };

  readonly fontList = FONTS;
  readonly sizeList = [8, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  readonly colorKeys = Object.keys(COLORS);

  readonly paperPadding = computed<string>(() => {
    const m = this.wordDoc.setup.margins;
    return `${m.top * 48}px ${m.right * 48}px ${m.bottom * 48}px ${m.left * 48}px`;
  });

  constructor(
    public bridge: PluginBridgeService,
    public theme: ThemeService,
    private confetti: ConfettiService,
    public wordDoc: WordDocumentService
  ) {
    // A host may send a new activity at any time (e.g. after activityCompleted
    // it sends the next lesson as a separate PLUGIN_DATA_IN). Reload + fully
    // reset the plugin whenever the activity signal changes, so a finished
    // plugin always accepts new data. (ngAfterViewInit covers an activity that
    // arrived before the editor existed; loadWordActivity is idempotent.)
    effect(() => {
      const act = this.bridge.activity();
      if (act) {
        untracked(() => this.loadWordActivity(act));
      }
    });

    effect(() => {
      const trigger = this.bridge.forceFinishTrigger();
      if (trigger > 0) {
        untracked(() => {
          if (this.advanceTimer) {
            clearTimeout(this.advanceTimer);
            this.advanceTimer = null;
          }
          this.isDemoOpen.set(false);
          const total = this.totalTasks();
          const passedCount = this.passedTasks().size;
          const allDone = passedCount >= total;
          this.bridge.sendComplete(allDone, passedCount, total);
          this.isCompleted.set(true);
        });
      }
    }, { allowSignalWrites: true });
  }

  ngOnInit(): void {
    this.bridge.init('word');
    const initialAct = this.bridge.activity();
    if (initialAct) {
      this.loadWordActivity(initialAct);
    } else {
      this.loadFallbackActivity();
    }
  }

  ngAfterViewInit(): void {
    if (this.editorRef) {
      const act = this.currentActivity();
      const task = this.currentTask();
      if (act && task) {
        this.wordDoc.init(task, act, this.editorRef.nativeElement);
        this.captureTaskStartSnapshot(this.taskIndex(), act, task);
      }
    }

    setTimeout(() => {
      this.computeInitialPlacement();
      this.updateHeaderFooterDom();
    }, 50);
  }

  ngOnDestroy(): void {
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
    }
  }

  toggleMenu(name: string, event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.openMenuName.update(current => (current === name ? null : name));
  }

  closeMenus(): void {
    this.openMenuName.set(null);
    this.showTablePicker.set(false);
    this.showOrientationPicker.set(false);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event?: MouseEvent): void {
    const target = event?.target as HTMLElement | null;
    if (!target?.closest('.ribbon-tab-menu') && !target?.closest('.ribbon-popover-wrap')) {
      this.closeMenus();
    }
    if (!target?.closest('.color-picker-wrap')) {
      this.showFontColorPicker.set(false);
      this.showHighlightColorPicker.set(false);
    }
    if (!target?.closest('.orientation-dropdown-wrap')) {
      this.showOrientationPicker.set(false);
    }
  }

  @HostListener('document:keydown.escape')
  onEscapeKey(): void {
    this.closeMenus();
    this.showFontColorPicker.set(false);
    this.showHighlightColorPicker.set(false);
    this.showCommentDialog.set(false);
    this.showOrientationPicker.set(false);
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.activeRegion.set('body');
  }

  @HostListener('document:keydown', ['$event'])
  onDocumentKeydown(e: KeyboardEvent): void {
    if (e.key === 'Backspace' || e.key === 'Delete') {
      if (this.selectedImageEl()) {
        e.preventDefault();
        this.deleteSelectedImage();
        return;
      }
      if (this.selectedPageNum()) {
        e.preventDefault();
        this.deletePageNumbers();
        return;
      }
      if (this.selectedTocEl()) {
        e.preventDefault();
        this.deleteSelectedToc();
        return;
      }
      if (this.selectedTableEl()) {
        const sel = window.getSelection();
        const anchor = sel?.anchorNode;
        const el = anchor?.nodeType === Node.ELEMENT_NODE ? anchor as Element : anchor?.parentElement;
        const inCell = el?.closest('td, th');
        if (!inCell) {
          e.preventDefault();
          this.deleteSelectedTable();
          return;
        }
      }
    }
  }

  private captureTaskStartSnapshot(index: number, act: ActivityPayload, task: LessonTask): void {
    let startingHtml = '';
    if (task.startingDocument !== undefined && task.startingDocument !== null) {
      startingHtml = sanitizeHtml(task.startingDocument, act.assetBaseUrl);
    } else if (this.editorRef) {
      startingHtml = this.editorRef.nativeElement.innerHTML;
    } else {
      startingHtml = this.wordDoc.html || '';
    }

    const snapshot: TaskStartSnapshot = {
      documentHtml: startingHtml,
      pageSetup: clone(this.wordDoc.setup),
      comments: clone(this.wordDoc.comments),
      header: this.wordDoc.header,
      footer: this.wordDoc.footer,
      numbering: this.wordDoc.numbering,
      variables: clone(act.variables || {}),
      images: act.images,
      assetBaseUrl: act.assetBaseUrl || '',
    };
    this.taskStartSnapshots.set(index, snapshot);
    this.currentTaskSnapshot.set(snapshot);
  }

  loadWordActivity(act: ActivityPayload): void {
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }

    let tasks: LessonTask[] = (act.tasks || []).map((t) => ({
      ...t,
      showMeHow: t.showMeHow === true,
      startingDocument: t.startingDocument !== undefined ? t.startingDocument : act.startingDocument,
    }));
    if (!tasks.length) {
      tasks = [
        {
          title: act.title || 'Apply formatting',
          explanation: 'Practise the tool, then check your work.',
          instructions: [act.instruction || 'Complete the lesson task.'],
          goals: act.goals || [],
          showMeHow: act.showMeHow === true,
          startingDocument: act.startingDocument,
        },
      ];
    }

    const normalizedAct: ActivityPayload = {
      ...act,
      tasks,
    };

    this.currentActivity.set(normalizedAct);
    this.taskIndex.set(0);
    this.passedTasks.set(new Set());
    this.isPassed.set(false);
    this.isCompleted.set(false);
    this.feedbackMsg.set('');
    this.isDemoOpen.set(false);
    this.showFinalTealCard.set(false);
    this.showPageSetup.set(false);
    this.showFindReplace.set(false);
    this.showCommentDialog.set(false);
    this.commentSelectedText.set('');
    this.newCommentText.set('Please review this section.');
    this.showOrientationPicker.set(false);
    this.activeRibbonTab.set('Home');
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.findTextQuery = '';
    this.replaceTextQuery = '';
    this.taskStartSnapshots.clear();

    const firstTask = tasks[0];
    const initialDoc = firstTask.startingDocument !== undefined ? firstTask.startingDocument : (act.startingDocument || '');
    if (initialDoc !== undefined && initialDoc !== null) {
      this.wordDoc.html = initialDoc;
    }
    this.updateWordCount();

    if (this.editorRef) {
      this.setupDocumentForTask(firstTask, normalizedAct);
    }
    this.captureTaskStartSnapshot(0, normalizedAct, firstTask);
  }

  loadTaskAtIndex(index: number): void {
    if (!this.editorRef) return;
    const act = this.currentActivity();
    if (!act || !act.tasks) return;
    const task = act.tasks[index];
    if (!task) return;

    this.taskIndex.set(index);
    this.isPassed.set(false);
    this.feedbackMsg.set('');
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);

    this.setupDocumentForTask(task, act);
    this.captureTaskStartSnapshot(index, act, task);
  }

  private setupDocumentForTask(task: LessonTask, act: ActivityPayload): void {
    if (!this.editorRef) return;
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.activeRegion.set('body');
    this.activePageNum.set(1);

    const rawHtml = task.startingDocument !== undefined && task.startingDocument !== null
      ? task.startingDocument
      : (act.startingDocument || '');

    // Check if document contains a page break element
    if (rawHtml && (rawHtml.includes('data-element="pageBreak"') || rawHtml.includes('pageBreak'))) {
      const parts = rawHtml.split(/<hr[^>]*data-element=["']pageBreak["'][^>]*>/i);
      if (parts.length > 1) {
        this.wordDoc.init({ ...task, startingDocument: parts[0] + '<hr data-element="pageBreak" style="display:none">' }, act, this.editorRef.nativeElement);
        this.hasPageBreak.set(true);
        this.wordDoc.hasPageBreak = true;
        setTimeout(() => {
          if (this.page2EditorRef?.nativeElement) {
            this.page2EditorRef.nativeElement.innerHTML = parts.slice(1).join('') || '<p><br></p>';
            this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
          }
          this.updateHeaderFooterDom();
          this.syncDocumentContent();
        });
        return;
      }
    }

    this.hasPageBreak.set(false);
    this.wordDoc.hasPageBreak = false;
    this.wordDoc.page2Element = null;
    if (task.startingDocument !== null && task.startingDocument !== undefined) {
      this.wordDoc.init(task, act, this.editorRef.nativeElement);
    } else {
      this.wordDoc.actions = [];
    }
    this.updateHeaderFooterDom();
    this.updateWordCount();
  }

  loadFallbackActivity(index: number = 0): void {
    const clamped = Math.max(0, Math.min(index, this.standaloneLessons.length - 1));
    this.standaloneLessonIndex.set(clamped);
    this.bridge.resetCompleted();
    const sample = this.standaloneLessons[clamped];
    if (sample) {
      this.loadWordActivity(sample);
    }
  }

  selectStandaloneLesson(idx: number): void {
    this.loadFallbackActivity(idx);
  }

  onEditorInput(): void {
    const editor = this.editorRef?.nativeElement;
    if (editor) {
      // 1. Clean up empty variable tokens
      editor.querySelectorAll<HTMLElement>('span[data-element="variable"]').forEach((v) => {
        if (!v.textContent?.trim()) {
          v.remove();
        }
      });

      // 2. Clean up orphaned footnote section if all footnote references are gone
      const fnRefs = editor.querySelectorAll('sup.footnote-ref');
      const fnSection = editor.querySelector('.footnote-section');
      if (fnSection && fnRefs.length === 0) {
        fnSection.remove();
      }

      // 3. Clean up completely empty list shells
      editor.querySelectorAll<HTMLElement>('ul, ol').forEach((list) => {
        if (!list.textContent?.trim() && !list.contains(document.activeElement)) {
          list.remove();
        }
      });

      // 4. Prune empty page break if removed from page 1 DOM
      if (this.hasPageBreak() && !editor.querySelector('[data-element="pageBreak"]')) {
        this.mergePage2IntoPage1();
      }
    }
    this.syncDocumentContent();
  }

  onPage2Input(): void {
    this.syncDocumentContent();
  }

  onEditorKeydown(e: KeyboardEvent): void {
    if (e.key === 'Delete' && this.hasPageBreak()) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
        const range = sel.getRangeAt(0);
        const testRange = document.createRange();
        testRange.setStart(range.endContainer, range.endOffset);
        testRange.setEndAfter(this.editorRef.nativeElement.lastChild || this.editorRef.nativeElement);
        if (testRange.toString().trim() === '') {
          e.preventDefault();
          this.mergePage2IntoPage1();
          return;
        }
      }
    }

    if (e.key === 'Backspace') {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
        const anchor = sel.anchorNode;
        const el = anchor?.nodeType === Node.ELEMENT_NODE ? anchor as Element : anchor?.parentElement;

        // If inside an empty table
        const cell = el?.closest('td, th');
        const table = cell?.closest('table');
        if (cell && table) {
          const isEmpty = Array.from(table.querySelectorAll('td, th')).every(c => !c.textContent?.trim());
          if (isEmpty) {
            e.preventDefault();
            const p = document.createElement('p');
            p.innerHTML = '<br>';
            table.parentNode?.insertBefore(p, table);
            table.remove();
            this.selectedTableEl.set(null);
            this.onEditorInput();
            const r = document.createRange();
            r.setStart(p, 0);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
            return;
          }
        }

        // If in an empty li
        const li = el?.closest('li');
        if (li && !li.textContent?.trim()) {
          const list = li.closest('ul, ol');
          if (list) {
            e.preventDefault();
            li.remove();
            if (list.querySelectorAll('li').length === 0) {
              const p = document.createElement('p');
              p.innerHTML = '<br>';
              list.parentNode?.insertBefore(p, list);
              list.remove();
              const r = document.createRange();
              r.setStart(p, 0);
              r.collapse(true);
              sel.removeAllRanges();
              sel.addRange(r);
            }
            this.onEditorInput();
            return;
          }
        }
      }
    }
  }

  mergePage2IntoPage1(): void {
    const editor = this.editorRef?.nativeElement;
    const p2 = this.page2EditorRef?.nativeElement;
    if (editor) {
      const marker = editor.querySelector('[data-element="pageBreak"]');
      if (marker) marker.remove();
      if (p2) {
        const text = p2.innerText.trim();
        if (text) {
          editor.innerHTML += p2.innerHTML;
        }
      }
      editor.focus();
    }
    this.hasPageBreak.set(false);
    this.wordDoc.hasPageBreak = false;
    this.wordDoc.page2Element = null;
    this.activePageNum.set(1);
    this.syncDocumentContent();
  }

  onPage2Keydown(e: KeyboardEvent): void {
    if (e.key === 'Backspace' && this.page2EditorRef) {
      const p2 = this.page2EditorRef.nativeElement;
      const text = p2.innerText.trim();
      const sel = window.getSelection();
      const isAtStart = sel && sel.rangeCount > 0 && sel.anchorOffset === 0;

      if (!text || isAtStart) {
        e.preventDefault();
        this.mergePage2IntoPage1();
      }
    }
  }

  syncDocumentContent(): void {
    const p1Html = this.editorRef?.nativeElement.innerHTML || '';
    const p2Html = this.hasPageBreak() && this.page2EditorRef?.nativeElement ? this.page2EditorRef.nativeElement.innerHTML : '';
    this.wordDoc.html = p1Html + (this.hasPageBreak() ? '<hr data-element="pageBreak" style="display:none">' + p2Html : '');
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
    } else {
      this.wordDoc.page2Element = null;
    }
    this.updateWordCount();
  }

  updateHeaderFooterDom(): void {
    if (this.headerRef?.nativeElement && document.activeElement !== this.headerRef.nativeElement) {
      this.headerRef.nativeElement.textContent = this.wordDoc.header || '';
    }
    if (this.footerRef?.nativeElement && document.activeElement !== this.footerRef.nativeElement) {
      this.footerRef.nativeElement.textContent = this.wordDoc.footer || '';
    }
  }

  onHeaderInput(e: Event): void {
    const text = (e.target as HTMLElement).textContent || '';
    this.wordDoc.header = text.trim() === '' ? '' : text;
  }

  onFooterInput(e: Event): void {
    const text = (e.target as HTMLElement).textContent || '';
    this.wordDoc.footer = text.trim() === '' ? '' : text;
  }

  removeHeader(): void {
    this.wordDoc.header = '';
    if (this.headerRef?.nativeElement) {
      this.headerRef.nativeElement.textContent = '';
    }
    this.wordDoc.logAction('header', '');
    this.closeHeaderFooterEdit();
  }

  removeFooter(): void {
    this.wordDoc.footer = '';
    if (this.footerRef?.nativeElement) {
      this.footerRef.nativeElement.textContent = '';
    }
    this.wordDoc.logAction('footer', '');
    this.closeHeaderFooterEdit();
  }

  onHeaderBlur(): void {
    // Keep active region or update if needed
  }

  onFooterBlur(): void {
    // Keep active region or update if needed
  }

  openHeaderEdit(e?: MouseEvent): void {
    if (e) e.stopPropagation();
    this.activeRegion.set('header');
    setTimeout(() => {
      this.headerRef?.nativeElement?.focus();
    });
  }

  openFooterEdit(e?: MouseEvent): void {
    if (e) e.stopPropagation();
    this.activeRegion.set('footer');
    setTimeout(() => {
      this.footerRef?.nativeElement?.focus();
    });
  }

  closeHeaderFooterEdit(): void {
    this.activeRegion.set('body');
  }

  private updateWordCount(): void {
    let text = this.editorRef?.nativeElement.innerText || '';
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      text += ' ' + this.page2EditorRef.nativeElement.innerText;
    }
    if (!text && this.wordDoc.html) {
      const temp = document.createElement('div');
      temp.innerHTML = this.wordDoc.html;
      text = temp.innerText || temp.textContent || '';
    }
    const words = text.trim().split(/\s+/).filter(Boolean).length;
    this.wordCount.set(words);
  }

  formatCmd(cmd: string, val?: string): void {
    this.wordDoc.rich?.command(cmd, val);
    if (cmd === 'undo' || cmd === 'redo') {
      this.wordDoc.logAction(cmd);
    }
    this.onEditorInput();
  }

  cutText(): void {
    this.wordDoc.rich?.copy(true);
    this.onEditorInput();
  }

  copyText(): void {
    this.wordDoc.rich?.copy(false);
  }

  pasteText(): void {
    this.wordDoc.rich?.paste();
    this.onEditorInput();
  }

  applyFont(font: string): void {
    this.wordDoc.rich?.apply({ fontFamily: font });
    this.onEditorInput();
  }

  applyFontSize(size: string): void {
    this.wordDoc.rich?.apply({ fontSize: size + 'pt' });
    this.onEditorInput();
  }

  applyColor(colorNameOrHex: string, colorKey?: string): void {
    const c = COLORS[colorNameOrHex] || colorNameOrHex;
    this.wordDoc.rich?.applyColor(c);
    this.activeFontColor.set(c);
    this.wordDoc.logAction('fontColor', colorKey || colorNameOrHex);
    this.onEditorInput();
  }

  applyHighlight(colorNameOrHex: string, colorKey?: string): void {
    const c = COLORS[colorNameOrHex] || colorNameOrHex;
    this.wordDoc.rich?.applyHighlight(c);
    this.activeHighlightColor.set(c && c !== 'transparent' ? c : '#ffff00');
    this.wordDoc.logAction('highlight', colorKey || colorNameOrHex);
    this.onEditorInput();
  }

  toggleFontColorPicker(event?: MouseEvent): void {
    if (event) {
      event.stopPropagation();
    }
    this.showHighlightColorPicker.set(false);
    this.showFontColorPicker.update((v) => !v);
    this.wordDoc.rich?.capture();
  }

  toggleHighlightPicker(event?: MouseEvent): void {
    if (event) {
      event.stopPropagation();
    }
    this.showFontColorPicker.set(false);
    this.showHighlightColorPicker.update((v) => !v);
    this.wordDoc.rich?.capture();
  }

  selectFontColor(hex: string, colorKey?: string): void {
    this.applyColor(hex, colorKey);
    this.showFontColorPicker.set(false);
  }

  selectHighlightColor(hex: string, colorKey?: string): void {
    this.applyHighlight(hex, colorKey);
    this.showHighlightColorPicker.set(false);
  }

  applyAutomaticFontColor(): void {
    this.activeFontColor.set('#000000');
    this.showFontColorPicker.set(false);
    this.wordDoc.rich?.applyColor('#000000');
    this.wordDoc.logAction('fontColor', 'automatic');
    this.onEditorInput();
  }

  applyNoHighlight(): void {
    this.activeHighlightColor.set('#ffff00');
    this.showHighlightColorPicker.set(false);
    this.wordDoc.rich?.applyHighlight('transparent');
    this.wordDoc.logAction('highlight', 'none');
    this.onEditorInput();
  }

  applyLineSpacing(val: string): void {
    const b = this.wordDoc.rich?.block();
    if (b) {
      b.style.lineHeight = val;
      b.dataset['lineSpacing'] = val;
      this.wordDoc.logAction('lineSpacing', val);
      this.onEditorInput();
    }
  }

  applyParagraphStyle(style: string): void {
    this.selectedStyle.set(style);
    const tag = style === 'normal' ? 'p' : style === 'h1' || style === 'h2' || style === 'h3' ? style : 'p';
    this.wordDoc.rich?.command('formatBlock', tag);
    const b = this.wordDoc.rich?.block();
    if (b) {
      b.dataset['paragraphStyle'] = style;
      if (style === 'title') {
        b.classList.add('doc-title');
      } else if (style === 'subtitle') {
        b.classList.add('doc-subtitle');
      }
      this.onEditorInput();
    }
  }

  openImageDialog(): void {
    const imgs = this.currentActivity()?.images;
    this.selectedImageFile.set(imgs?.[0]?.file || 'photo.png');
    this.showImageDialog.set(true);
    this.closeMenus();
  }

  insertChosenImage(): void {
    const file = this.selectedImageFile() || 'photo.png';
    const base = this.currentActivity()?.assetBaseUrl || './assets/';
    const img = document.createElement('img');
    img.src = base ? (base.endsWith('/') ? base + file : base + '/' + file) : file;
    img.dataset['file'] = file;
    img.dataset['wrap'] = 'inline';
    img.dataset['align'] = 'left';
    img.style.maxWidth = '300px';
    img.style.cursor = 'pointer';
    this.editorRef?.nativeElement.append(img);
    this.wordDoc.logAction('insertImage', file);
    this.showImageDialog.set(false);
    this.onEditorInput();
  }

  insertSampleImage(): void {
    this.openImageDialog();
  }

  onEditorClick(event: MouseEvent): void {
    this.activeRegion.set('body');
    const target = event.target as HTMLElement;
    if (target && target.tagName === 'IMG') {
      const img = target as HTMLImageElement;
      this.selectedImageEl.set(img);
      this.wordDoc.logAction('selectImage', img.dataset['file'] || img.src);
      setTimeout(() => this.updateImageBoxRect());
    } else {
      this.selectedImageEl.set(null);
    }

    const table = target?.closest('table') as HTMLTableElement | null;
    this.selectedTableEl.set(table);

    const toc = target?.closest('.table-of-contents') as HTMLElement | null;
    this.selectedTocEl.set(toc);

    if (!target?.closest('.page-num')) {
      this.selectedPageNum.set(false);
    }
  }

  deleteSelectedImage(): void {
    const img = this.selectedImageEl();
    if (img) {
      img.remove();
      this.selectedImageEl.set(null);
      this.wordDoc.logAction('deleteImage');
      this.onEditorInput();
    }
  }

  deleteSelectedTable(): void {
    const table = this.selectedTableEl();
    if (table) {
      table.remove();
      this.selectedTableEl.set(null);
      this.wordDoc.logAction('deleteTable');
      this.onEditorInput();
    }
  }

  deleteSelectedToc(): void {
    const toc = this.selectedTocEl();
    if (toc) {
      toc.remove();
      this.selectedTocEl.set(null);
      this.wordDoc.logAction('deleteTableOfContents');
      this.onEditorInput();
    }
  }

  selectPageNumber(e: MouseEvent): void {
    e.stopPropagation();
    this.selectedPageNum.set(true);
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
  }

  deletePageNumbers(e?: Event): void {
    if (e) e.preventDefault();
    this.wordDoc.numbering = false;
    this.selectedPageNum.set(false);
    this.wordDoc.logAction('pageNumbers');
    this.onEditorInput();
  }

  onPage1Click(event: MouseEvent): void {
    this.activePageNum.set(1);
    this.onEditorClick(event);
  }

  onPage2Click(event: MouseEvent): void {
    this.activePageNum.set(2);
    this.onEditorClick(event);
  }

  onPaperClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (!target.closest('.page-header') && !target.closest('.page-footer')) {
      this.activeRegion.set('body');
    }
  }

  alignSelectedImage(align: 'left' | 'center' | 'right'): void {
    const img = this.selectedImageEl();
    if (img) {
      img.dataset['align'] = align;
      if (align === 'center') {
        img.style.display = 'block';
        img.style.margin = '8px auto';
      } else if (align === 'right') {
        img.style.display = 'block';
        img.style.margin = '8px 0 8px auto';
      } else {
        img.style.display = 'inline-block';
        img.style.margin = '0';
      }
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  wrapSelectedImage(wrap: 'inline' | 'wrap' | 'break'): void {
    const img = this.selectedImageEl();
    if (img) {
      img.dataset['wrap'] = wrap;
      if (wrap === 'wrap') {
        img.style.float = 'left';
        img.style.margin = '0 12px 12px 0';
      } else if (wrap === 'break') {
        img.style.display = 'block';
        img.style.clear = 'both';
      } else {
        img.style.float = 'none';
        img.style.display = 'inline-block';
      }
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  resizeSelectedImage(width: number): void {
    const img = this.selectedImageEl();
    if (img) {
      img.width = width;
      img.style.width = width + 'px';
      img.style.height = 'auto';
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  updateImageBoxRect(): void {
    const img = this.selectedImageEl();
    if (!img) return;
    const canvas = document.querySelector('.word-canvas') as HTMLElement | null;
    if (!canvas) return;

    const cRect = canvas.getBoundingClientRect();
    const iRect = img.getBoundingClientRect();

    this.imageBoxLeft.set(Math.round(iRect.left - cRect.left + canvas.scrollLeft));
    this.imageBoxTop.set(Math.round(iRect.top - cRect.top + canvas.scrollTop));
    this.imageBoxWidth.set(Math.round(iRect.width));
    this.imageBoxHeight.set(Math.round(iRect.height));
  }

  onResizeStart(e: MouseEvent, handle: string): void {
    e.preventDefault();
    e.stopPropagation();
    const img = this.selectedImageEl();
    if (!img) return;

    this.isResizing = true;
    this.resizeHandle = handle;
    this.resizeStartX = e.clientX;
    this.resizeStartY = e.clientY;
    this.resizeStartWidth = img.clientWidth || img.width || 300;
    this.resizeStartHeight = img.clientHeight || img.height || 200;
    this.resizeAspectRatio = this.resizeStartWidth / (this.resizeStartHeight || 1);

    const onMouseMove = (moveEvt: MouseEvent) => {
      if (!this.isResizing || !this.selectedImageEl()) return;
      const dx = moveEvt.clientX - this.resizeStartX;
      const dy = moveEvt.clientY - this.resizeStartY;
      let newWidth = this.resizeStartWidth;

      if (handle === 'se') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'sw') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 'ne') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'nw') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 'e') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'w') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 's' || handle === 'n') {
        const factor = handle === 's' ? 1 : -1;
        const newHeight = Math.max(40, this.resizeStartHeight + dy * factor);
        newWidth = Math.round(newHeight * this.resizeAspectRatio);
      }

      const targetImg = this.selectedImageEl()!;
      targetImg.width = Math.round(newWidth);
      targetImg.style.width = `${Math.round(newWidth)}px`;
      targetImg.style.height = 'auto';
      this.updateImageBoxRect();
    };

    const onMouseUp = () => {
      this.isResizing = false;
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      this.onEditorInput();
      this.updateImageBoxRect();
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  insertGridTable(rows: number, cols: number): void {
    const table = document.createElement('table');
    table.className = 'doc-table';
    for (let i = 0; i < rows; i++) {
      const r = table.insertRow();
      for (let j = 0; j < cols; j++) {
        const c = r.insertCell();
        c.innerHTML = '<br>';
      }
    }
    this.editorRef?.nativeElement.append(table);
    this.wordDoc.logAction('insertTable');
    this.showTablePicker.set(false);
    this.closeMenus();
    this.onEditorInput();
  }

  insertSampleTable(): void {
    this.insertGridTable(2, 2);
  }

  insertPageBreak(): void {
    if (!this.editorRef) return;
    const editor = this.editorRef.nativeElement;

    let moveHtml = '';
    const sel = window.getSelection();

    if (sel && sel.rangeCount > 0 && editor.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      const afterRange = document.createRange();
      afterRange.setStart(range.endContainer, range.endOffset);
      afterRange.setEndAfter(editor.lastChild || editor);
      const frag = afterRange.extractContents();
      const div = document.createElement('div');
      div.appendChild(frag);
      moveHtml = div.innerHTML.trim();
    }

    const hr = document.createElement('hr');
    hr.dataset['element'] = 'pageBreak';
    hr.className = 'page-break-marker';
    hr.style.display = 'none';
    editor.appendChild(hr);

    this.hasPageBreak.set(true);
    this.wordDoc.hasPageBreak = true;
    this.wordDoc.logAction('pageBreak');
    this.closeMenus();

    setTimeout(() => {
      if (this.page2EditorRef?.nativeElement) {
        this.page2EditorRef.nativeElement.innerHTML = moveHtml || '<p><br></p>';
        this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
        this.page2EditorRef.nativeElement.focus();
      }
      this.syncDocumentContent();
      this.activePageNum.set(2);
    });
  }

  toggleHeader(): void {
    const text = this.wordDoc.header || 'Document Header';
    this.wordDoc.header = text;
    this.activeRegion.set('header');
    this.wordDoc.logAction('header', text);
    this.closeMenus();
    setTimeout(() => {
      this.updateHeaderFooterDom();
      this.headerRef?.nativeElement?.focus();
    });
  }

  toggleFooter(): void {
    const text = this.wordDoc.footer || 'Document Footer';
    this.wordDoc.footer = text;
    this.activeRegion.set('footer');
    this.wordDoc.logAction('footer', text);
    this.closeMenus();
    setTimeout(() => {
      this.updateHeaderFooterDom();
      this.footerRef?.nativeElement?.focus();
    });
  }

  togglePageNumbers(): void {
    this.wordDoc.numbering = !this.wordDoc.numbering;
    this.wordDoc.logAction('pageNumbers');
    this.closeMenus();
  }

  insertTableOfContents(): void {
    if (!this.editorRef) return;
    const headings = Array.from(this.editorRef.nativeElement.querySelectorAll('h1, h2, h3, [data-paragraph-style="title"], [data-paragraph-style="h1"], [data-paragraph-style="h2"], [data-paragraph-style="h3"]'));
    const toc = document.createElement('div');
    toc.className = 'table-of-contents doc-toc';
    toc.dataset['element'] = 'tableOfContents';
    toc.tabIndex = 0;

    let listHtml = '';
    if (headings.length > 0) {
      headings.forEach((h: any, i: number) => {
        const id = h.id || `heading-anchor-${i}`;
        h.id = id;
        const level = h.tagName === 'H1' ? 1 : h.tagName === 'H2' ? 2 : 3;
        listHtml += `<li class="toc-item toc-level-${level}"><a href="#${id}">${h.textContent || 'Section ' + (i + 1)}</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li>`;
      });
    } else {
      listHtml = `<li class="toc-item"><a href="#doc-top">Document Overview</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li><li class="toc-item"><a href="#doc-details">Details</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li>`;
    }

    toc.innerHTML = `<div class="toc-header-bar"><strong>Table of Contents</strong><button type="button" class="toc-remove-btn" title="Remove Table of Contents" aria-label="Remove Table of Contents">✕</button></div><ul class="toc-list">${listHtml}</ul>`;
    const btn = toc.querySelector('.toc-remove-btn');
    if (btn) {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        toc.remove();
        this.selectedTocEl.set(null);
        this.wordDoc.logAction('deleteTableOfContents');
        this.onEditorInput();
      });
    }

    this.editorRef.nativeElement.prepend(toc);
    this.wordDoc.logAction('tableOfContents');
    this.wordDoc.logAction('navigateToc');
    this.closeMenus();
    this.onEditorInput();
  }

  insertFootnote(): void {
    if (!this.editorRef) return;
    const currentNotes = this.editorRef.nativeElement.querySelectorAll('sup.footnote-ref').length;
    const noteNum = currentNotes + 1;
    const fnRef = document.createElement('sup');
    fnRef.className = 'footnote-ref';
    fnRef.dataset['element'] = 'footnote';
    fnRef.dataset['noteId'] = String(noteNum);
    fnRef.tabIndex = 0;
    fnRef.textContent = `[${noteNum}]`;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && this.editorRef.nativeElement.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(fnRef);
      range.setStartAfter(fnRef);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      this.editorRef.nativeElement.append(fnRef);
    }

    let fnSection = this.editorRef.nativeElement.querySelector<HTMLElement>('.footnote-section');
    if (!fnSection) {
      fnSection = document.createElement('div');
      fnSection.className = 'footnote-section';
      fnSection.innerHTML = '<hr class="footnote-divider"/>';
      this.editorRef.nativeElement.append(fnSection);
    }
    const fnItem = document.createElement('p');
    fnItem.className = 'footnote';
    fnItem.dataset['element'] = 'footnote';
    fnItem.dataset['noteId'] = String(noteNum);
    fnItem.innerHTML = `<sup>${noteNum}</sup> Footnote ${noteNum} details <button type="button" class="footnote-remove-btn" title="Remove footnote" style="background:none;border:none;cursor:pointer;color:#888;font-size:11px;margin-left:6px;">✕</button>`;
    fnSection.append(fnItem);

    const btn = fnItem.querySelector('.footnote-remove-btn');
    if (btn) {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        fnItem.remove();
        fnRef.remove();
        if (fnSection && fnSection.querySelectorAll('.footnote').length === 0) {
          fnSection.remove();
        }
        this.onEditorInput();
      });
    }

    this.wordDoc.logAction('footnote');
    this.closeMenus();
    this.onEditorInput();
  }

  insertComment(commentText?: string, targetText?: string): void {
    const sel = window.getSelection();
    const text = targetText !== undefined ? targetText : (sel ? sel.toString().trim() : '');
    if (commentText !== undefined) {
      this.wordDoc.comments.push({
        text: text || 'Selected text',
        comment: commentText || 'Please review this section.',
        resolved: false,
        replies: []
      });
      this.wordDoc.logAction('comment', text || commentText);
      this.closeMenus();
      this.onEditorInput();
      return;
    }
    this.commentSelectedText.set(text || 'Selected text');
    this.newCommentText.set('Please review this section.');
    this.showCommentDialog.set(true);
    this.closeMenus();
  }

  confirmInsertComment(): void {
    const comment = this.newCommentText().trim() || 'Please review this section.';
    const text = this.commentSelectedText();
    this.wordDoc.comments.push({
      text: text || 'Selected text',
      comment: comment,
      resolved: false,
      replies: []
    });
    this.wordDoc.logAction('comment', text || comment);
    this.showCommentDialog.set(false);
    this.newCommentText.set('Please review this section.');
    this.onEditorInput();
  }

  cancelInsertComment(): void {
    this.showCommentDialog.set(false);
    this.newCommentText.set('Please review this section.');
  }

  toggleResolveComment(index: number): void {
    if (this.wordDoc.comments[index]) {
      this.wordDoc.comments[index].resolved = !this.wordDoc.comments[index].resolved;
      this.wordDoc.logAction('resolveComment');
      this.onEditorInput();
    }
  }

  addCommentReply(index: number, text: string): void {
    const trimmed = (text || '').trim();
    if (!trimmed || !this.wordDoc.comments[index]) return;
    if (!this.wordDoc.comments[index].replies) {
      this.wordDoc.comments[index].replies = [];
    }
    this.wordDoc.comments[index].replies!.push(trimmed);
    this.wordDoc.logAction('replyComment', trimmed);
    this.onEditorInput();
  }

  deleteComment(index: number): void {
    if (this.wordDoc.comments && index >= 0 && index < this.wordDoc.comments.length) {
      this.wordDoc.comments.splice(index, 1);
      this.wordDoc.logAction('deleteComment');
      this.onEditorInput();
    }
  }

  showVariablesPanel = signal<boolean>(false);
  openVariablesPanel(): void {
    this.showVariablesPanel.set(true);
    this.wordDoc.logAction('openVariables');
    this.closeMenus();
  }

  getVariablesList(): { key: string; val: string }[] {
    const vars = this.currentActivity()?.variables;
    if (!vars || typeof vars !== 'object') return [];
    return Object.entries(vars).map(([key, val]) => ({ key, val: String(val) }));
  }

  insertVariable(key: string, val: string): void {
    const span = document.createElement('span');
    span.dataset['element'] = 'variable';
    span.dataset['variableKey'] = key;
    span.className = 'doc-variable-token';
    span.textContent = val;
    span.tabIndex = 0;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(span);
      range.setStartAfter(span);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      this.editorRef?.nativeElement.append(span);
    }

    this.wordDoc.logAction('variable', key);
    this.showVariablesPanel.set(false);
    this.onEditorInput();
  }

  selectAllText(): void {
    if (!this.editorRef) return;
    const r = document.createRange();
    r.selectNodeContents(this.editorRef.nativeElement);
    const s = window.getSelection();
    if (s) {
      s.removeAllRanges();
      s.addRange(r);
    }
    this.wordDoc.logAction('select', this.editorRef.nativeElement.innerText || '');
    this.closeMenus();
  }

  setOrientation(o: 'portrait' | 'landscape'): void {
    this.wordDoc.setup.orientation = o;
    this.wordDoc.logAction('pageSetup');
    this.closeMenus();
  }

  openPageSetupDialog(): void {
    this.tempSetup = clone(this.wordDoc.setup);
    this.wordDoc.logAction('openPageSetup');
    this.showPageSetup.set(true);
    this.closeMenus();
  }

  savePageSetup(): void {
    this.wordDoc.setup = clone(this.tempSetup);
    this.wordDoc.logAction('pageSetup');
    this.showPageSetup.set(false);
  }

  findTab = signal<'headings' | 'find' | 'replace'>('replace');
  openFindReplaceDialog(tab: 'find' | 'replace' = 'replace'): void {
    this.findTab.set(tab);
    this.updateMatchCount();
    this.wordDoc.logAction('openFindReplace');
    this.showFindReplace.set(true);
  }

  updateMatchCount(): void {
    if (!this.findTextQuery || !this.editorRef) {
      this.matchCount.set(0);
      return;
    }
    const text = this.editorRef.nativeElement.innerText || '';
    const query = this.matchCase() ? this.findTextQuery : this.findTextQuery.toLowerCase();
    const source = this.matchCase() ? text : text.toLowerCase();
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, 'g');
    const matches = source.match(regex);
    this.matchCount.set(matches ? matches.length : 0);
  }

  findNext(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    this.wordDoc.logAction('find', this.findTextQuery);
    if ((window as any).find) {
      (window as any).find(this.findTextQuery, this.matchCase(), false, true);
    }
    this.updateMatchCount();
  }

  findPrev(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    this.wordDoc.logAction('find', this.findTextQuery);
    if ((window as any).find) {
      (window as any).find(this.findTextQuery, this.matchCase(), true, true);
    }
    this.updateMatchCount();
  }

  executeReplace(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    const html = this.editorRef.nativeElement.innerHTML;
    const escaped = this.findTextQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const flags = this.matchCase() ? '' : 'i';
    const replaced = html.replace(new RegExp(escaped, flags), this.replaceTextQuery);
    this.editorRef.nativeElement.innerHTML = replaced;
    this.wordDoc.logAction('find', this.findTextQuery);
    this.wordDoc.logAction('findReplace', this.findTextQuery);
    this.updateMatchCount();
    this.showFindReplace.set(false);
    this.onEditorInput();
  }

  executeReplaceAll(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    const html = this.editorRef.nativeElement.innerHTML;
    const escaped = this.findTextQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const flags = this.matchCase() ? 'g' : 'gi';
    const replaced = html.replace(new RegExp(escaped, flags), this.replaceTextQuery);
    this.editorRef.nativeElement.innerHTML = replaced;
    this.wordDoc.logAction('find', this.findTextQuery);
    this.wordDoc.logAction('findReplace', this.findTextQuery);
    this.updateMatchCount();
    this.showFindReplace.set(false);
    this.onEditorInput();
  }

  selectFontColorSwatch(colorKey: string): void {
    const hex = COLORS[colorKey] || colorKey;
    this.selectFontColor(hex, colorKey);
  }

  selectHighlightColorSwatch(colorKey: string): void {
    const hex = COLORS[colorKey] || colorKey;
    this.selectHighlightColor(hex, colorKey);
  }

  updateActiveFormatting(): void {
    this.wordDoc.rich?.capture();
    const s = window.getSelection();
    if (!s || !s.anchorNode) return;
    const el = (s.anchorNode.nodeType === Node.ELEMENT_NODE
      ? s.anchorNode
      : s.anchorNode.parentElement) as HTMLElement | null;
    if (el && this.editorRef?.nativeElement.contains(el)) {
      const cs = window.getComputedStyle(el);
      if (cs.color) {
        this.activeFontColor.set(cs.color);
      }
      if (cs.backgroundColor && cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent') {
        this.activeHighlightColor.set(cs.backgroundColor);
      }
    }
  }

  openDemoModal(): void {
    const curIdx = this.taskIndex();
    const existing = this.taskStartSnapshots.get(curIdx);
    if (existing) {
      this.currentTaskSnapshot.set(existing);
    }
    this.isDemoOpen.set(true);
  }

  onCheckWork(): void {
    if (this.bridge.isPaused() || this.isCompleted()) return;
    const task = this.currentTask();
    if (!task) return;

    let allPassed = false;
    try {
      allPassed = task.goals.length === 0 || task.goals.every((g) => this.wordDoc.check(g));
    } catch (err: any) {
      this.bridge.sendError('Goal check failed: ' + err?.message);
    }

    if (allPassed) {
      this.selectedImageEl.set(null);
      this.isPassed.set(true);
      this.feedbackMsg.set(task.goals.length === 0 ? '' : "Well done! You've got this!");
      this.confetti.celebrate();

      const passedSet = new Set(this.passedTasks());
      passedSet.add(this.taskIndex());
      this.passedTasks.set(passedSet);

      if (this.isAssessment()) {
        const isLast = this.taskIndex() === this.totalTasks() - 1;
        if (!isLast) {
          this.advanceTimer = setTimeout(() => {
            this.loadTaskAtIndex(this.taskIndex() + 1);
          }, 1000);
        } else {
          this.showFinalTealCard.set(true);
        }
      }
    } else {
      this.isPassed.set(false);
      this.feedbackMsg.set(task.hint || 'Not quite right. Try again.');
    }
  }

  onContinueNext(): void {
    this.selectedImageEl.set(null);
    const total = this.totalTasks();
    const curIdx = this.taskIndex();

    if (this.isAssessment()) {
      if (curIdx + 1 < total) {
        this.loadTaskAtIndex(curIdx + 1);
        return;
      }
      // Last task of assessment: send completed and stop!
      const passedCount = this.passedTasks().size || total;
      this.bridge.sendComplete(true, passedCount, total);
      this.isCompleted.set(true);
      return;
    }

    // In a lesson (single task): send completed once and stop!
    this.bridge.sendComplete(true, 1, 1);
    this.isCompleted.set(true);
  }

  // --- Movable Task Card & Mascot Group Positioning ---

  effectiveX(): number | null {
    return this.groupX();
  }

  effectiveY(): number | null {
    return this.groupY();
  }

  computeInitialPlacement(): void {
    if (!this.workLayoutRef || !this.paperRef) return;
    const layoutEl = this.workLayoutRef.nativeElement;
    const paperEl = this.paperRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    const paperRect = paperEl.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupWidth = groupEl?.offsetWidth || 330;
    const groupHeight = groupEl?.offsetHeight || 440;

    // Right side of the Word page
    let targetX = (paperRect.right - layoutRect.left) + 24;

    // If there is not enough room to the right of the paper, dock near the right edge
    if (targetX + groupWidth > layoutRect.width - 16) {
      targetX = layoutRect.width - groupWidth - 16;
    }
    if (targetX < 12) {
      targetX = 12;
    }

    // Vertically align near the top of the document page
    let targetY = Math.max(16, (paperRect.top - layoutRect.top) + 12);
    const maxY = Math.max(16, layoutRect.height - groupHeight - 16);
    if (targetY > maxY) {
      targetY = maxY;
    }

    this.groupX.set(Math.round(targetX));
    this.groupY.set(Math.round(targetY));
  }

  clampX(x: number): number {
    if (!this.workLayoutRef) return Math.max(8, x);
    const layoutRect = this.workLayoutRef.nativeElement.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupWidth = groupEl?.offsetWidth || 330;
    const minX = 8;
    const maxX = Math.max(minX, layoutRect.width - groupWidth - 8);
    return Math.min(Math.max(x, minX), maxX);
  }

  clampY(y: number): number {
    if (!this.workLayoutRef) return Math.max(8, y);
    const layoutRect = this.workLayoutRef.nativeElement.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupHeight = groupEl?.offsetHeight || 440;
    const minY = 8;
    const maxY = Math.max(minY, layoutRect.height - groupHeight - 8);
    return Math.min(Math.max(y, minY), maxY);
  }

  onGroupPointerDown(event: PointerEvent): void {
    const target = event.target as HTMLElement | null;
    const isDragHandle = target?.closest('.card-drag-handle');
    if (!isDragHandle) {
      return; // Do not initiate drag from regular buttons or card body
    }

    this.isDragging = true;
    this.dragStartPointerX = event.clientX;
    this.dragStartPointerY = event.clientY;

    const groupEl = this.movableGroupRef?.nativeElement;
    this.dragStartGroupX = this.groupX() ?? groupEl?.offsetLeft ?? 0;
    this.dragStartGroupY = this.groupY() ?? groupEl?.offsetTop ?? 0;

    try {
      (event.currentTarget as HTMLElement)?.setPointerCapture(event.pointerId);
    } catch {}
    event.preventDefault();
  }

  onGroupPointerMove(event: PointerEvent): void {
    if (!this.isDragging) return;
    const dx = event.clientX - this.dragStartPointerX;
    const dy = event.clientY - this.dragStartPointerY;
    const nextX = this.clampX(this.dragStartGroupX + dx);
    const nextY = this.clampY(this.dragStartGroupY + dy);
    this.groupX.set(Math.round(nextX));
    this.groupY.set(Math.round(nextY));
    this.hasUserMoved.set(true);
  }

  onGroupPointerUp(event: PointerEvent): void {
    if (this.isDragging) {
      this.isDragging = false;
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  onGroupPointerCancel(event: PointerEvent): void {
    if (this.isDragging) {
      this.isDragging = false;
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  onNudgePosition(delta: { dx: number; dy: number }): void {
    const groupEl = this.movableGroupRef?.nativeElement;
    const curX = this.groupX() ?? groupEl?.offsetLeft ?? 0;
    const curY = this.groupY() ?? groupEl?.offsetTop ?? 0;
    this.groupX.set(Math.round(this.clampX(curX + delta.dx)));
    this.groupY.set(Math.round(this.clampY(curY + delta.dy)));
    this.hasUserMoved.set(true);
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (!this.hasUserMoved()) {
      this.computeInitialPlacement();
    } else if (this.groupX() !== null && this.groupY() !== null) {
      this.groupX.set(Math.round(this.clampX(this.groupX()!)));
      this.groupY.set(Math.round(this.clampY(this.groupY()!)));
    }
  }
}
