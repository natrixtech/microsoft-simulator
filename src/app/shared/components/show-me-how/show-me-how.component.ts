import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnDestroy,
  OnInit,
  Output,
  ViewChild,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { TaskGoal } from '../../models/activity.types';
import { COLORS, findText, safeAsset } from '../../utils/editor.utils';

export interface DemoStep {
  label: string;
  state: 'pending' | 'active' | 'done';
}

@Component({
  selector: 'lib-show-me-how',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './show-me-how.component.html',
  styleUrls: ['./show-me-how.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ShowMeHowComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('replicaDocContent') replicaDocContentRef?: ElementRef<HTMLElement>;
  @ViewChild('replicaArea') replicaAreaRef?: ElementRef<HTMLElement>;

  @Input() appKind: 'word' | 'excel' | 'powerpoint' = 'word';
  @Input() title: string = '';
  @Input() instruction: string = '';
  @Input() goals: TaskGoal[] | any[] = [];
  @Input() startingDocument?: string | null = '';
  @Input() pageSetup?: any = null;
  @Input() comments?: any[] = [];
  @Input() variables?: Record<string, string> = {};
  @Input() images?: any = null;
  @Input() assetBaseUrl: string = '';
  @Input() isPaused: boolean = false;

  @Output() readonly close = new EventEmitter<void>();
  @Output() readonly closed = new EventEmitter<void>();

  readonly cursorPixelX = signal<number>(80);
  readonly cursorPixelY = signal<number>(80);
  readonly isClicking = signal<boolean>(false);
  readonly showToast = signal<boolean>(false);
  readonly steps = signal<DemoStep[]>([]);

  readonly activeAction = signal<string>('');
  readonly replicaDialog = signal<string | null>(null);

  // Tab switching state
  readonly replicaActiveTab = signal<'home' | 'insert' | 'layout' | 'tools'>('home');

  // Coherent scaling state
  readonly baseWidth = 680;
  readonly scaleTransform = signal<string>('none');
  readonly scalerWidth = signal<string>('100%');
  readonly scalerHeight = signal<string>('100%');
  private resizeObserver?: ResizeObserver;

  // Word replica UI state
  readonly replicaFont = signal<string>('Calibri');
  readonly replicaFontSize = signal<string>('11');
  readonly replicaStyle = signal<string>('normal');
  readonly replicaOrientation = signal<'portrait' | 'landscape'>('portrait');
  readonly replicaPaperSize = signal<string>('Letter');
  readonly replicaMargins = signal<{ top: number; bottom: number; left: number; right: number }>({
    top: 1,
    bottom: 1,
    left: 1,
    right: 1,
  });
  readonly replicaHeader = signal<string>('');
  readonly replicaFooter = signal<string>('');
  readonly replicaPageNumbers = signal<boolean>(false);
  readonly replicaComments = signal<any[]>([]);

  // Find & replace typing simulation
  readonly findWhatInput = signal<string>('');
  readonly replaceWithInput = signal<string>('');

  // Dialog specific helpers
  readonly replicaImageFile = signal<string>('photo.png');
  readonly replicaTableRows = signal<number>(2);
  readonly replicaTableCols = signal<number>(2);

  private isCancelled = false;
  private activeTimers: any[] = [];

  constructor(private elRef: ElementRef, private cdr: ChangeDetectorRef) {}

  ngOnInit(): void {
    this.setupSteps();
  }

  ngAfterViewInit(): void {
    this.resetReplicaState();
    this.updateScaling();
    this.cdr.detectChanges();
    if (typeof ResizeObserver !== 'undefined' && this.replicaAreaRef) {
      this.resizeObserver = new ResizeObserver(() => {
        this.updateScaling();
      });
      this.resizeObserver.observe(this.replicaAreaRef.nativeElement);
    }
    this.play();
  }

  ngOnDestroy(): void {
    this.cancelAll();
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = undefined;
    }
  }

  onClose(): void {
    this.cancelAll();
    this.close.emit();
    this.closed.emit();
  }

  onBackdropClick(e: MouseEvent): void {
    this.onClose();
  }

  replay(): void {
    this.cancelAll();
    this.resetReplicaState();
    this.setupSteps();
    this.play();
  }

  getToastTitle(): string {
    const t = (this.title || '').trim();
    if (!t) return 'complete this task';
    return t.toLowerCase();
  }

  getVariableEntries(): { key: string; value: string }[] {
    const v = this.variables || {};
    const entries = Object.entries(v);
    if (!entries.length) {
      return [{ key: 'companyName', value: 'Contoso Ltd.' }];
    }
    return entries.map(([key, value]) => ({ key, value }));
  }

  private cancelAll(): void {
    this.isCancelled = true;
    for (const t of this.activeTimers) {
      clearTimeout(t);
    }
    this.activeTimers = [];
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => {
      let timerId: any = null;

      const finish = () => {
        const idx = this.activeTimers.indexOf(timerId);
        if (idx >= 0) this.activeTimers.splice(idx, 1);
        if (this.isCancelled) {
          resolve();
          return;
        }
        if (this.isPaused) {
          timerId = setTimeout(finish, 50);
          this.activeTimers.push(timerId);
          return;
        }
        resolve();
      };

      timerId = setTimeout(finish, ms);
      this.activeTimers.push(timerId);
    });
  }

  private setupSteps(): void {
    const list: DemoStep[] = (this.goals || []).map((g) => ({
      label: this.getStepLabelForGoal(g),
      state: 'pending',
    }));
    this.steps.set(list);
  }

  private getStepLabelForGoal(g: any): string {
    switch (g?.type) {
      case 'textStyle':
        return g.text ? `Format "${g.text}"` : 'Apply text formatting';
      case 'paragraph':
        return g.style ? `Set style to ${g.style}` : 'Format paragraph';
      case 'list':
        return g.listType === 'number' ? 'Create numbered list' : 'Create bullet list';
      case 'textContains':
        return g.notContains?.length ? `Replace "${g.notContains[0]}"` : 'Edit document text';
      case 'textCount':
        return `Adjust "${g.text}" count to ${g.count}`;
      case 'textOrder':
        return 'Reorder paragraphs';
      case 'image':
        return 'Insert image';
      case 'pageSetup':
        return 'Change page setup';
      case 'table':
        return 'Insert table';
      case 'element':
        return `Insert ${g.element || 'element'}`;
      case 'action':
        return `Do ${g.action}`;
      default:
        return 'Complete step';
    }
  }

  private getDocElement(): HTMLElement | null {
    if (this.replicaDocContentRef?.nativeElement) {
      return this.replicaDocContentRef.nativeElement;
    }
    const container = this.elRef.nativeElement.querySelector('.replica-window') as HTMLElement;
    return container?.querySelector('.replica-doc-content') as HTMLElement;
  }

  private resetReplicaState(): void {
    this.isCancelled = false;
    this.showToast.set(false);
    this.activeAction.set('');
    this.replicaDialog.set(null);
    this.findWhatInput.set('');
    this.replaceWithInput.set('');

    // Restore Word document from snapshot
    const docEl = this.getDocElement();
    const rawDoc = this.startingDocument !== undefined && this.startingDocument !== null
      ? this.startingDocument
      : '';
    if (docEl) {
      docEl.innerHTML = rawDoc;
    }

    this.replicaFont.set('Calibri');
    this.replicaFontSize.set('11');
    this.replicaStyle.set('normal');

    this.replicaActiveTab.set('home');
    this.updateScaling();

    // Restore page setup
    const ps = this.pageSetup || {};
    this.replicaOrientation.set(ps.orientation === 'landscape' ? 'landscape' : 'portrait');
    this.replicaPaperSize.set(ps.paperSize || 'Letter');
    this.replicaMargins.set({
      top: ps.margins?.top ?? 1,
      bottom: ps.margins?.bottom ?? 1,
      left: ps.margins?.left ?? 1,
      right: ps.margins?.right ?? 1,
    });

    this.replicaHeader.set('');
    this.replicaFooter.set('');
    this.replicaPageNumbers.set(false);

    // Restore comments
    const initialComments = Array.isArray(this.comments) ? this.comments.map((c) => ({ ...c })) : [];
    this.replicaComments.set(initialComments);

    // Initial image selection
    let defaultImg = 'photo.png';
    if (this.images) {
      if (Array.isArray(this.images) && this.images.length > 0) {
        defaultImg = typeof this.images[0] === 'string' ? this.images[0] : (this.images[0].file || 'photo.png');
      } else if (typeof this.images === 'object') {
        const keys = Object.keys(this.images);
        if (keys.length > 0) defaultImg = keys[0];
      }
    }
    this.replicaImageFile.set(defaultImg);
  }

  setReplicaTab(tab: 'home' | 'insert' | 'layout' | 'tools'): void {
    this.replicaActiveTab.set(tab);
    this.cdr.markForCheck();
  }

  updateScaling(): void {
    if (!this.replicaAreaRef) return;
    const el = this.replicaAreaRef.nativeElement;
    const w = el.clientWidth;
    const h = el.clientHeight;
    if (w <= 0 || h <= 0) return;

    if (w >= this.baseWidth) {
      this.scaleTransform.set('none');
      this.scalerWidth.set('100%');
      this.scalerHeight.set('100%');
    } else {
      const scale = w / this.baseWidth;
      this.scaleTransform.set(`scale(${scale})`);
      this.scalerWidth.set(`${this.baseWidth}px`);
      this.scalerHeight.set(`${Math.round(h / scale)}px`);
    }
    this.cdr.markForCheck();
  }

  private positionCursorOver(targetRect: DOMRect): void {
    const area =
      this.replicaAreaRef?.nativeElement ||
      (this.elRef.nativeElement.querySelector('.demo-replica-area') as HTMLElement) ||
      (this.elRef.nativeElement.querySelector('.replica-window') as HTMLElement);
    if (!area) return;

    const areaRect = area.getBoundingClientRect();
    const x = targetRect.left - areaRect.left + (targetRect.width > 0 ? targetRect.width / 2 : 10);
    const y = targetRect.top - areaRect.top + (targetRect.height > 0 ? targetRect.height / 2 : 10);

    this.cursorPixelX.set(Math.round(x));
    this.cursorPixelY.set(Math.round(y));
    this.cdr.markForCheck();
  }

  private async moveTo(targetName: string): Promise<boolean> {
    if (this.isCancelled) return false;

    // Switch ribbon tab automatically if targeting tab buttons or tab-specific controls
    if (targetName === 'tab-insert') {
      this.replicaActiveTab.set('insert');
    } else if (targetName === 'tab-layout') {
      this.replicaActiveTab.set('layout');
    } else if (targetName === 'tab-tools') {
      this.replicaActiveTab.set('tools');
    } else if (targetName === 'tab-home') {
      this.replicaActiveTab.set('home');
    } else if (
      [
        'insert-image',
        'insert-table',
        'insert-page-break',
        'insert-header',
        'insert-footer',
        'insert-page-numbers',
        'insert-toc',
        'insert-footnote',
        'insert-comment',
      ].includes(targetName)
    ) {
      this.replicaActiveTab.set('insert');
    } else if (targetName === 'page-setup') {
      this.replicaActiveTab.set('layout');
    } else if (targetName === 'tools-variables') {
      this.replicaActiveTab.set('tools');
    } else if (
      [
        'bold',
        'italic',
        'underline',
        'font',
        'size',
        'color',
        'highlight',
        'bullets',
        'numbering',
        'indent',
        'align-left',
        'align-center',
        'align-right',
        'align-justify',
        'line-spacing',
        'paste',
        'cut',
        'copy',
        'find',
        'replace',
      ].includes(targetName) ||
      targetName.startsWith('style-')
    ) {
      this.replicaActiveTab.set('home');
    }
    this.cdr.markForCheck();

    const host = this.elRef.nativeElement;
    let target = host.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;
    if (!target) {
      for (let i = 0; i < 6; i++) {
        await this.delay(60);
        if (this.isCancelled) return false;
        target = host.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;
        if (target) break;
      }
    }
    if (!target) return false;

    this.positionCursorOver(target.getBoundingClientRect());

    await this.delay(300);
    if (this.isCancelled) return false;

    this.isClicking.set(true);
    await this.delay(100);
    this.isClicking.set(false);
    return true;
  }

  private async moveToRange(range: Range): Promise<boolean> {
    if (this.isCancelled) return false;

    const rects = range.getClientRects();
    const targetRect = rects.length > 0 ? rects[0] : range.getBoundingClientRect();
    this.positionCursorOver(targetRect);

    await this.delay(300);
    if (this.isCancelled) return false;

    this.isClicking.set(true);
    await this.delay(100);
    this.isClicking.set(false);
    return true;
  }

  private setStepState(idx: number, state: 'pending' | 'active' | 'done'): void {
    this.steps.update((s) =>
      s.map((step, i) => (i === idx ? { ...step, state } : step))
    );
  }

  async play(): Promise<void> {
    await this.delay(350);
    if (this.isCancelled) return;

    const goalsList = this.goals || [];
    const total = goalsList.length;

    if (total === 0) {
      await this.delay(300);
      if (this.isCancelled) return;
      this.showToast.set(true);
      return;
    }

    let allSuccess = true;
    for (let i = 0; i < total; i++) {
      if (this.isCancelled) return;
      this.setStepState(i, 'active');
      const g = goalsList[i];

      const ok = await this.executeWordGoal(g);
      if (this.isCancelled) return;

      if (ok) {
        this.setStepState(i, 'done');
      } else {
        allSuccess = false;
        // Do not mark step done if required target or action failed
      }
      await this.delay(300);
    }

    if (this.isCancelled) return;
    if (allSuccess) {
      this.showToast.set(true);
    }
  }

  private async executeWordGoal(g: any): Promise<boolean> {
    const type = g?.type;
    const docEl = this.getDocElement();
    if (!docEl) return false;

    if (type === 'textStyle') {
      const targetText = g.text;
      if (!targetText) return false;
      const range = findText(docEl, targetText);
      if (!range) return false;

      // 1. Animate cursor to exact text and select it
      await this.moveToRange(range);
      const selSpan = document.createElement('span');
      selSpan.className = 'demo-selection-box';
      selSpan.dataset['demoStyled'] = 'true';
      try {
        const contents = range.extractContents();
        selSpan.appendChild(contents);
        range.insertNode(selSpan);
      } catch {
        return false;
      }
      await this.delay(300);

      // 2. Bold (handling explicit false correctly)
      if (g.bold !== undefined) {
        this.activeAction.set('bold');
        await this.moveTo('bold');
        selSpan.style.fontWeight = g.bold ? 'bold' : 'normal';
        await this.delay(250);
      }

      // 3. Italic (handling explicit false correctly)
      if (g.italic !== undefined) {
        this.activeAction.set('italic');
        await this.moveTo('italic');
        selSpan.style.fontStyle = g.italic ? 'italic' : 'normal';
        await this.delay(250);
      }

      // 4. Underline (handling explicit false correctly)
      if (g.underline !== undefined) {
        this.activeAction.set('underline');
        await this.moveTo('underline');
        selSpan.style.textDecoration = g.underline ? 'underline' : 'none';
        await this.delay(250);
      }

      // 5. Font
      if (g.font) {
        this.replicaFont.set(g.font);
        await this.moveTo('font');
        selSpan.style.fontFamily = g.font;
        await this.delay(250);
      }

      // 6. Size
      if (g.size) {
        this.replicaFontSize.set(String(g.size));
        await this.moveTo('size');
        selSpan.style.fontSize = g.size + 'pt';
        await this.delay(250);
      }

      // 7. Color (mapping through COLORS palette)
      if (g.color) {
        this.activeAction.set('color');
        await this.moveTo('color');
        const mappedColor = COLORS[g.color] || g.color;
        selSpan.style.color = mappedColor;
        await this.delay(250);
      }

      // 8. Highlight (mapping through COLORS palette)
      if (g.highlight) {
        this.activeAction.set('highlight');
        await this.moveTo('highlight');
        const mappedHl = COLORS[g.highlight] || g.highlight;
        selSpan.style.backgroundColor = mappedHl;
        await this.delay(250);
      }

      // Remove selection highlight while preserving applied formatting
      selSpan.classList.remove('demo-selection-box');
      this.activeAction.set('');
      return true;
    }

    if (type === 'paragraph') {
      const targetText = g.text;
      if (!targetText) return false;
      const range = findText(docEl, targetText);
      if (!range) return false;

      const block = range.startContainer.parentElement?.closest('p, h1, h2, h3, li, div') as HTMLElement | null;
      if (!block) return false;

      // Animate cursor to paragraph
      await this.moveToRange(range);

      // Alignment
      if (g.align) {
        this.activeAction.set('align-' + g.align);
        await this.moveTo('align-' + g.align);
        block.style.textAlign = g.align;
        await this.delay(250);
      }

      // Indent (supporting zero indent)
      if (g.indent !== undefined) {
        this.activeAction.set('indent');
        await this.moveTo('indent');
        block.style.marginLeft = g.indent === 0 ? '0px' : g.indent * 40 + 'px';
        await this.delay(250);
      }

      // Line Spacing
      if (g.lineSpacing !== undefined) {
        this.activeAction.set('line-spacing');
        await this.moveTo('line-spacing');
        block.style.lineHeight = String(g.lineSpacing);
        block.dataset['lineSpacing'] = String(g.lineSpacing);
        await this.delay(250);
      }

      // Paragraph Style
      if (g.style) {
        this.replicaStyle.set(g.style);
        await this.moveTo('style-' + g.style);
        block.dataset['paragraphStyle'] = g.style;
        if (['h1', 'h2', 'h3'].includes(g.style)) {
          const heading = document.createElement(g.style);
          heading.innerHTML = block.innerHTML;
          heading.style.cssText = block.style.cssText;
          Object.assign(heading.dataset, block.dataset);
          heading.dataset['paragraphStyle'] = g.style;
          block.replaceWith(heading);
        } else if (g.style === 'title') {
          block.classList.add('doc-title');
        } else if (g.style === 'subtitle') {
          block.classList.add('doc-subtitle');
        }
        await this.delay(250);
      }

      this.activeAction.set('');
      return true;
    }

    if (type === 'list') {
      const items = (g.items || []) as string[];
      if (!items.length) return false;

      // Locate lines in document matching requested items
      const foundBlocks: HTMLElement[] = [];
      for (const it of items) {
        const r = findText(docEl, it);
        if (r) {
          const b = r.startContainer.parentElement?.closest('p, li, div') as HTMLElement | null;
          if (b && !foundBlocks.includes(b)) foundBlocks.push(b);
        }
      }

      if (foundBlocks.length > 0) {
        await this.moveToRange(findText(docEl, items[0])!);
      } else {
        await this.moveTo('doc-canvas');
      }

      const listTarget = g.listType === 'number' ? 'numbering' : 'bullets';
      this.activeAction.set(listTarget);
      await this.moveTo(listTarget);

      const listTag = g.listType === 'number' ? 'ol' : 'ul';
      const listEl = document.createElement(listTag);
      if (g.level !== undefined && g.level > 0) {
        listEl.style.marginLeft = g.level * 24 + 'px';
      }

      items.forEach((it) => {
        const li = document.createElement('li');
        li.textContent = it;
        listEl.appendChild(li);
      });

      if (foundBlocks.length > 0) {
        foundBlocks[0].replaceWith(listEl);
        for (let j = 1; j < foundBlocks.length; j++) {
          foundBlocks[j].remove();
        }
      } else {
        docEl.appendChild(listEl);
      }

      await this.delay(300);
      this.activeAction.set('');
      return true;
    }

    if (type === 'textContains') {
      const notContains = (g.notContains || []) as string[];
      const contains = (g.contains || []) as string[];

      // Find & replace for each notContains item
      if (notContains.length > 0) {
        for (let k = 0; k < notContains.length; k++) {
          const wrong = notContains[k];
          const right = contains[k] !== undefined ? contains[k] : '';

          this.activeAction.set('replace');
          await this.moveTo('replace');
          this.replicaDialog.set('findReplace');
          await this.delay(250);

          // Type wrong word
          this.findWhatInput.set('');
          await this.moveTo('find-what-input');
          for (const ch of wrong) {
            if (this.isCancelled) return false;
            this.findWhatInput.update((v) => v + ch);
            await this.delay(50);
          }

          // Type replacement word
          this.replaceWithInput.set('');
          await this.moveTo('replace-with-input');
          for (const ch of right) {
            if (this.isCancelled) return false;
            this.replaceWithInput.update((v) => v + ch);
            await this.delay(50);
          }

          await this.moveTo('dialog-replace-btn');

          // In-document replacement
          let cur = docEl.innerHTML;
          if (cur.includes(wrong)) {
            docEl.innerHTML = cur.split(wrong).join(right);
          }
          await this.delay(250);
          this.replicaDialog.set(null);
        }
      }

      // Check missing contains words and type them
      if (contains.length > 0) {
        for (const wanted of contains) {
          if (!docEl.textContent?.includes(wanted)) {
            await this.moveTo('doc-canvas');
            const p = document.createElement('p');
            p.textContent = wanted;
            docEl.appendChild(p);
            await this.delay(250);
          }
        }
      }

      this.activeAction.set('');
      return true;
    }

    if (type === 'textCount') {
      const targetText = g.text;
      const needed = g.count;
      if (!targetText || typeof needed !== 'number') return false;

      let currentCount = (docEl.textContent || '').split(targetText).length - 1;

      if (currentCount < needed) {
        const r = findText(docEl, targetText);
        if (r) {
          await this.moveToRange(r);
          this.activeAction.set('copy');
          await this.moveTo('copy');
        } else {
          await this.moveTo('doc-canvas');
        }

        while (currentCount < needed) {
          await this.moveTo('doc-canvas');
          this.activeAction.set('paste');
          await this.moveTo('paste');
          const p = document.createElement('p');
          p.textContent = targetText;
          docEl.appendChild(p);
          currentCount++;
          await this.delay(200);
        }
      } else if (currentCount > needed) {
        while (currentCount > needed) {
          const r = findText(docEl, targetText);
          if (!r) break;
          await this.moveToRange(r);
          this.activeAction.set('cut');
          await this.moveTo('cut');
          r.deleteContents();
          currentCount--;
          await this.delay(200);
        }
      }

      this.activeAction.set('');
      return true;
    }

    if (type === 'textOrder') {
      const seq = (g.sequence || []) as string[];
      if (!seq.length) return false;

      for (const itemText of seq) {
        const r = findText(docEl, itemText);
        if (!r) return false;
      }

      // Demonstrate reordering
      for (let sIdx = 0; sIdx < seq.length; sIdx++) {
        const it = seq[sIdx];
        const r = findText(docEl, it);
        if (!r) continue;
        const block = r.startContainer.parentElement?.closest('p, h1, h2, h3, li, div') as HTMLElement | null;
        if (block) {
          await this.moveToRange(r);
          this.activeAction.set('cut');
          await this.moveTo('cut');
          block.remove();

          await this.moveTo('doc-canvas');
          this.activeAction.set('paste');
          await this.moveTo('paste');
          const p = document.createElement('p');
          p.textContent = it;
          docEl.appendChild(p);
          await this.delay(200);
        }
      }

      this.activeAction.set('');
      return true;
    }

    if (type === 'image') {
      let file = g.file;
      if (!file) {
        if (Array.isArray(this.images) && this.images.length > 0) {
          file = typeof this.images[0] === 'string' ? this.images[0] : (this.images[0].file || 'photo.png');
        } else {
          file = 'photo.png';
        }
      }
      this.replicaImageFile.set(file);

      this.activeAction.set('insert-image');
      await this.moveTo('insert-image');
      this.replicaDialog.set('insertImage');
      await this.delay(250);

      await this.moveTo('insert-image-ok-btn');
      this.replicaDialog.set(null);

      const img = document.createElement('img');
      img.src = safeAsset(this.assetBaseUrl, file);
      img.dataset['file'] = file;
      img.dataset['align'] = g.align || 'left';
      img.dataset['wrap'] = g.wrap || 'inline';
      if (g.minWidth) img.width = g.minWidth;
      if (g.width) img.width = g.width;
      if (g.maxWidth) img.style.maxWidth = g.maxWidth + 'px';

      const wrapDiv = document.createElement('div');
      wrapDiv.style.textAlign = g.align || 'left';
      wrapDiv.appendChild(img);
      docEl.appendChild(wrapDiv);

      await this.delay(250);
      this.activeAction.set('');
      return true;
    }

    if (type === 'pageSetup') {
      this.activeAction.set('page-setup');
      await this.moveTo('page-setup');
      this.replicaDialog.set('pageSetup');
      await this.delay(250);

      if (g.orientation) {
        const orientTarget = 'ps-orient-' + g.orientation;
        await this.moveTo(orientTarget);
        this.replicaOrientation.set(g.orientation);
        await this.delay(200);
      }

      if (g.paperSize) {
        await this.moveTo('ps-size');
        this.replicaPaperSize.set(g.paperSize);
        await this.delay(200);
      }

      if (g.margins) {
        for (const [side, val] of Object.entries(g.margins)) {
          const mTarget = 'ps-margin-' + side;
          await this.moveTo(mTarget);
          this.replicaMargins.update((cur) => ({ ...cur, [side]: val as number }));
          await this.delay(150);
        }
      }

      await this.moveTo('pagesetup-ok-btn');
      this.replicaDialog.set(null);
      this.activeAction.set('');
      return true;
    }

    if (type === 'table') {
      let rows = g.rows;
      let cols = g.columns;
      if ((!rows || !cols) && g.cells && g.cells.length) {
        rows = g.cells.length;
        cols = Math.max(...g.cells.map((r: string[]) => r.length));
      }
      rows = rows || 2;
      cols = cols || 2;
      this.replicaTableRows.set(rows);
      this.replicaTableCols.set(cols);

      this.activeAction.set('insert-table');
      await this.moveTo('insert-table');
      this.replicaDialog.set('insertTable');
      await this.delay(250);

      await this.moveTo('insert-table-ok-btn');
      this.replicaDialog.set(null);

      const tbl = document.createElement('table');
      tbl.className = 'doc-table';
      for (let r = 0; r < rows; r++) {
        const tr = tbl.insertRow();
        for (let c = 0; c < cols; c++) {
          const td = tr.insertCell();
          const cellText = g.cells?.[r]?.[c];
          td.textContent = cellText !== undefined ? cellText : '';
        }
      }
      docEl.appendChild(tbl);

      if (g.cells && g.cells.length) {
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            if (g.cells[r]?.[c]) {
              await this.moveTo('doc-canvas');
              break;
            }
          }
        }
      }

      await this.delay(250);
      this.activeAction.set('');
      return true;
    }

    if (type === 'element') {
      const el = g.element;
      if (!el) return false;

      if (el === 'pageBreak') {
        this.activeAction.set('insert-page-break');
        await this.moveTo('insert-page-break');
        const hr = document.createElement('hr');
        hr.className = 'page-break-line';
        hr.dataset['element'] = 'pageBreak';
        docEl.appendChild(hr);
      } else if (el === 'header') {
        this.activeAction.set('insert-header');
        await this.moveTo('insert-header');
        this.replicaHeader.set(g.text || 'Document Header');
      } else if (el === 'footer') {
        this.activeAction.set('insert-footer');
        await this.moveTo('insert-footer');
        this.replicaFooter.set(g.text || 'Page 1');
      } else if (el === 'pageNumbers') {
        this.activeAction.set('insert-page-numbers');
        await this.moveTo('insert-page-numbers');
        this.replicaPageNumbers.set(true);
      } else if (el === 'tableOfContents') {
        this.activeAction.set('insert-toc');
        await this.moveTo('insert-toc');
        const toc = document.createElement('div');
        toc.className = 'table-of-contents';
        toc.dataset['element'] = 'tableOfContents';
        toc.innerHTML = '<h4 class="toc-title">Table of Contents</h4><p>Heading 1 .......... 1</p>';
        docEl.prepend(toc);
      } else if (el === 'footnote') {
        this.activeAction.set('insert-footnote');
        await this.moveTo('insert-footnote');
        const fn = document.createElement('div');
        fn.className = 'footnote-section';
        fn.dataset['element'] = 'footnote';
        fn.innerHTML = `<hr class="footnote-divider"/><p><sup>1</sup> ${g.text || 'Footnote'}</p>`;
        docEl.appendChild(fn);
      } else if (el === 'comment') {
        if (g.text) {
          const r = findText(docEl, g.text);
          if (r) await this.moveToRange(r);
        }
        this.activeAction.set('insert-comment');
        await this.moveTo('insert-comment');

        const newComment = {
          text: g.text || 'Selected note',
          comment: g.text || 'Note content',
          resolved: false,
        };
        this.replicaComments.update((cur) => [...cur, newComment]);
        await this.delay(250);

        if (g.resolved === true) {
          const cIdx = this.replicaComments().length - 1;
          await this.moveTo('comment-resolve-' + cIdx);
          this.replicaComments.update((cur) =>
            cur.map((c, idx) => (idx === cIdx ? { ...c, resolved: true } : c))
          );
        }
      } else if (el === 'variable') {
        this.activeAction.set('tools-variables');
        await this.moveTo('tools-variables');
        this.replicaDialog.set('variables');
        await this.delay(250);
        await this.moveTo('variables-insert-btn');
        this.replicaDialog.set(null);

        const varKey = g.text || Object.keys(this.variables || {})[0] || 'variable';
        const span = document.createElement('span');
        span.className = 'doc-variable-token';
        span.dataset['element'] = 'variable';
        span.textContent = `{{${varKey}}}`;
        docEl.appendChild(span);
      } else {
        const div = document.createElement('div');
        div.dataset['element'] = el;
        div.textContent = g.text || el;
        docEl.appendChild(div);
      }

      await this.delay(250);
      this.activeAction.set('');
      return true;
    }

    if (type === 'action') {
      const act = g.action;
      if (!act) return false;

      if (act === 'select') {
        if (!g.text) return false;
        const r = findText(docEl, g.text);
        if (!r) return false;
        await this.moveToRange(r);
        const selSpan = document.createElement('span');
        selSpan.className = 'demo-selection-box';
        try {
          const contents = r.extractContents();
          selSpan.appendChild(contents);
          r.insertNode(selSpan);
        } catch {
          return false;
        }
      } else if (act === 'selectImage') {
        const img = docEl.querySelector('img');
        if (!img) return false;
        await this.moveTo('doc-canvas');
        img.classList.add('selected-image-outline');
      } else if (act === 'copy') {
        this.activeAction.set('copy');
        await this.moveTo('copy');
      } else if (act === 'cut') {
        this.activeAction.set('cut');
        await this.moveTo('cut');
        const sel = docEl.querySelector('.demo-selection-box');
        if (sel) sel.remove();
      } else if (act === 'paste') {
        await this.moveTo('doc-canvas');
        this.activeAction.set('paste');
        await this.moveTo('paste');
        if (g.text) {
          const p = document.createElement('p');
          p.textContent = g.text;
          docEl.appendChild(p);
        }
      } else if (act === 'undo') {
        await this.moveTo('undo');
      } else if (act === 'openFindReplace') {
        this.activeAction.set('replace');
        await this.moveTo('replace');
        this.replicaDialog.set('findReplace');
        await this.delay(300);
      } else if (act === 'find') {
        this.activeAction.set('find');
        await this.moveTo('find');
        if (g.text) {
          const r = findText(docEl, g.text);
          if (r) await this.moveToRange(r);
        }
      } else if (act === 'openPageSetup') {
        this.activeAction.set('page-setup');
        await this.moveTo('page-setup');
        this.replicaDialog.set('pageSetup');
        await this.delay(300);
      } else if (act === 'nextCell') {
        const tbl = docEl.querySelector('table');
        if (!tbl) return false;
        await this.moveTo('doc-canvas');
      } else if (act === 'tabNewRow') {
        const tbl = docEl.querySelector('table');
        if (!tbl) return false;
        tbl.insertRow().insertCell().textContent = '';
        await this.moveTo('doc-canvas');
      } else if (act === 'navigateToc') {
        const toc = docEl.querySelector('.table-of-contents');
        if (!toc) return false;
        await this.moveTo('doc-canvas');
      } else if (act === 'openVariables') {
        this.activeAction.set('tools-variables');
        await this.moveTo('tools-variables');
        this.replicaDialog.set('variables');
        await this.delay(300);
      } else {
        await this.moveTo('doc-canvas');
      }

      await this.delay(200);
      this.activeAction.set('');
      return true;
    }

    return false;
  }
}

