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

interface MenuItem {
  label: string;
  action?: () => void;
  disabled?: boolean;
}

interface MenuCategory {
  name: string;
  items: MenuItem[];
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
  @ViewChild('workLayout') workLayoutRef?: ElementRef<HTMLElement>;
  @ViewChild('wordPaper') paperRef?: ElementRef<HTMLElement>;
  @ViewChild('movableGroup') movableGroupRef?: ElementRef<HTMLElement>;

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

  menus: MenuCategory[] = [];

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
        this.isDemoOpen.set(false);
        const total = this.totalTasks();
        const passedCount = this.passedTasks().size;
        const allDone = passedCount >= total;
        this.bridge.sendComplete(allDone, passedCount, total);
        this.isCompleted.set(true);
      }
    });
  }

  ngOnInit(): void {
    this.bridge.init('word');
    this.setupMenus();
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
    }, 50);
  }

  ngOnDestroy(): void {
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
    }
  }

  private setupMenus(): void {
    this.menus = [
      {
        name: 'File',
        items: [
          { label: 'Page setup', action: () => this.openPageSetupDialog() },
          { label: 'Print preview', disabled: true },
        ],
      },
      {
        name: 'Edit',
        items: [
          { label: 'Undo', action: () => document.execCommand('undo') },
          { label: 'Redo', action: () => document.execCommand('redo') },
          { label: 'Cut', action: () => this.wordDoc.rich?.copy(true) },
          { label: 'Copy', action: () => this.wordDoc.rich?.copy(false) },
          { label: 'Paste', action: () => this.wordDoc.rich?.paste() },
          { label: 'Find & replace', action: () => this.openFindReplaceDialog() },
        ],
      },
      {
        name: 'Insert',
        items: [
          {
            label: 'Image',
            action: () => {
              const file = (this.currentActivity()?.images as any)?.[0]?.file || 'photo.png';
              const img = document.createElement('img');
              const base = this.currentActivity()?.assetBaseUrl || '';
              img.src = base ? base + file : file;
              img.dataset['file'] = file;
              img.style.maxWidth = '300px';
              this.editorRef?.nativeElement.append(img);
              this.wordDoc.logAction('insertImage', file);
              this.onEditorInput();
            },
          },
          {
            label: 'Table (2x2)',
            action: () => {
              const table = document.createElement('table');
              table.className = 'doc-table';
              for (let i = 0; i < 2; i++) {
                const r = table.insertRow();
                for (let j = 0; j < 2; j++) {
                  const c = r.insertCell();
                  c.innerHTML = '<br>';
                }
              }
              this.editorRef?.nativeElement.append(table);
              this.wordDoc.logAction('insertTable');
              this.onEditorInput();
            },
          },
          {
            label: 'Page break',
            action: () => {
              const hr = document.createElement('hr');
              hr.dataset['element'] = 'pageBreak';
              this.editorRef?.nativeElement.append(hr);
              this.wordDoc.logAction('pageBreak');
              this.onEditorInput();
            },
          },
          {
            label: 'Page header',
            action: () => {
              const text = this.wordDoc.header || 'Header';
              this.wordDoc.header = text;
              this.wordDoc.logAction('header', text);
            },
          },
          {
            label: 'Page footer',
            action: () => {
              const text = this.wordDoc.footer || 'Footer';
              this.wordDoc.footer = text;
              this.wordDoc.logAction('footer', text);
            },
          },
          {
            label: 'Page numbers',
            action: () => {
              this.wordDoc.numbering = !this.wordDoc.numbering;
              this.wordDoc.logAction('pageNumbers');
            },
          },
        ],
      },
      {
        name: 'Format',
        items: [
          { label: 'Paragraph style: Heading 1', action: () => this.applyParagraphStyle('h1') },
          { label: 'Paragraph style: Heading 2', action: () => this.applyParagraphStyle('h2') },
          { label: 'Paragraph style: Normal', action: () => this.applyParagraphStyle('normal') },
          { label: 'Align: Left', action: () => this.formatCmd('justifyLeft') },
          { label: 'Align: Center', action: () => this.formatCmd('justifyCenter') },
          { label: 'Align: Right', action: () => this.formatCmd('justifyRight') },
        ],
      },
      {
        name: 'Help',
        items: [
          { label: 'Show me how', action: () => this.openDemoModal() },
        ],
      },
    ];
  }

  toggleMenu(name: string, event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.openMenuName.update(current => (current === name ? null : name));
  }

  closeMenus(): void {
    this.openMenuName.set(null);
    this.showTablePicker.set(false);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event?: MouseEvent): void {
    const target = event?.target as HTMLElement | null;
    if (!target?.closest('.ribbon-tab-menu')) {
      this.closeMenus();
    }
    if (!target?.closest('.color-picker-wrap')) {
      this.showFontColorPicker.set(false);
      this.showHighlightColorPicker.set(false);
    }
  }

  @HostListener('document:keydown.escape')
  onEscapeKey(): void {
    this.closeMenus();
    this.showFontColorPicker.set(false);
    this.showHighlightColorPicker.set(false);
  }

  onMenuItemClick(item: MenuItem): void {
    if (item.action) {
      item.action();
    }
    this.closeMenus();
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
      this.wordDoc.init(firstTask, normalizedAct, this.editorRef.nativeElement);
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

    if (task.startingDocument !== null && task.startingDocument !== undefined) {
      this.wordDoc.init(task, act, this.editorRef.nativeElement);
    } else {
      this.wordDoc.actions = [];
    }
    this.updateWordCount();

    this.captureTaskStartSnapshot(index, act, task);
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
    this.wordDoc.html = this.editorRef?.nativeElement.innerHTML || '';
    this.updateWordCount();
  }

  onHeaderInput(e: Event): void {
    this.wordDoc.header = (e.target as HTMLElement).textContent || '';
  }

  onFooterInput(e: Event): void {
    this.wordDoc.footer = (e.target as HTMLElement).textContent || '';
  }

  private updateWordCount(): void {
    let text = this.editorRef?.nativeElement.innerText || '';
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
    const target = event.target as HTMLElement;
    if (target && target.tagName === 'IMG') {
      this.selectedImageEl.set(target as HTMLImageElement);
    } else {
      this.selectedImageEl.set(null);
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
    }
  }

  resizeSelectedImage(width: number): void {
    const img = this.selectedImageEl();
    if (img) {
      img.width = width;
      img.style.width = width + 'px';
      this.onEditorInput();
    }
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
    const hr = document.createElement('hr');
    hr.dataset['element'] = 'pageBreak';
    this.editorRef?.nativeElement.append(hr);
    this.wordDoc.logAction('pageBreak');
    this.closeMenus();
    this.onEditorInput();
  }

  toggleHeader(): void {
    const text = this.wordDoc.header || 'Document Header';
    this.wordDoc.header = text;
    this.wordDoc.logAction('header', text);
    this.closeMenus();
  }

  toggleFooter(): void {
    const text = this.wordDoc.footer || 'Document Footer';
    this.wordDoc.footer = text;
    this.wordDoc.logAction('footer', text);
    this.closeMenus();
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
    toc.contentEditable = 'false';

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

    toc.innerHTML = `<strong>Table of Contents</strong><ul class="toc-list">${listHtml}</ul>`;
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
    fnRef.textContent = `[${noteNum}]`;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      range.insertNode(fnRef);
    } else {
      this.editorRef.nativeElement.append(fnRef);
    }

    let fnSection = this.editorRef.nativeElement.querySelector<HTMLElement>('.footnote-section');
    if (!fnSection) {
      fnSection = document.createElement('div');
      fnSection.className = 'footnote-section';
      fnSection.setAttribute('contenteditable', 'false');
      fnSection.innerHTML = '<hr class="footnote-divider"/>';
      this.editorRef.nativeElement.append(fnSection);
    }
    const fnItem = document.createElement('p');
    fnItem.className = 'footnote';
    fnItem.dataset['element'] = 'footnote';
    fnItem.innerHTML = `<sup>${noteNum}</sup> Footnote ${noteNum} details`;
    fnSection.append(fnItem);

    this.wordDoc.logAction('footnote');
    this.closeMenus();
    this.onEditorInput();
  }

  insertComment(): void {
    const sel = window.getSelection();
    const text = sel ? sel.toString().trim() : '';
    const commentText = prompt('Enter your comment:', 'Please review this section.');
    if (commentText) {
      this.wordDoc.comments.push({
        text: text || 'Selected text',
        comment: commentText,
        resolved: false,
        replies: []
      });
      this.wordDoc.logAction('comment', text || commentText);
      this.closeMenus();
      this.onEditorInput();
    }
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
    span.textContent = val;
    this.editorRef?.nativeElement.append(span);
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

  findTab = signal<'find' | 'replace'>('replace');
  openFindReplaceDialog(tab: 'find' | 'replace' = 'replace'): void {
    this.findTab.set(tab);
    this.findTextQuery = '';
    this.replaceTextQuery = '';
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
