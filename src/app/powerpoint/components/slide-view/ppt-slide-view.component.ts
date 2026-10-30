import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  Optional,
  Output,
  SimpleChanges,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import {
  PlaceholderDef,
  PlaceholderType,
  ShapeKind,
  Slide,
  SlideElement,
  ThemeDef,
} from '../../models/ppt.types';
import { PPT_LAYOUTS } from '../../constants/ppt-layouts';
import { COLORS } from '../../constants/ppt-colors';
import { sanitizeRichText } from '../../utils/rich-text-sanitize';
import { PptPlaceholderComponent } from '../placeholder/ppt-placeholder.component';
import { PptSafeHtmlPipe } from '../../pipes/ppt-safe-html.pipe';
import { SLIDE_HEIGHT, SLIDE_WIDTH } from '../../utils/slide-geometry';
import { PptStateService } from '../../services/ppt-state.service';

@Component({
  selector: 'app-ppt-slide-view',
  standalone: true,
  imports: [CommonModule, PptPlaceholderComponent, PptSafeHtmlPipe],
  templateUrl: './ppt-slide-view.component.html',
  styleUrls: ['./ppt-slide-view.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptSlideViewComponent implements OnChanges {
  @Input() slide!: Slide;
  @Input() theme!: ThemeDef;
  @Input() isInteractive: boolean = false;
  @Input() selectedElementId: string | null = null;
  @Input() selectedElementIds: string[] = [];
  @Input() editingPlaceholderField: PlaceholderType | null = null;
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly selectElement = new EventEmitter<string>();
  @Output() readonly deselectElement = new EventEmitter<void>();
  @Output() readonly updateTableCellText = new EventEmitter<{
    elementId: string;
    row: number;
    col: number;
    text: string;
  }>();
  @Output() readonly updateTextBoxText = new EventEmitter<{
    elementId: string;
    text: string;
    html?: string;
  }>();

  @Output() readonly updatePlaceholder = new EventEmitter<{
    field: PlaceholderType;
    text: string;
    html?: string;
  }>();
  @Output() readonly startEditPlaceholder = new EventEmitter<PlaceholderType>();
  @Output() readonly stopEditPlaceholder = new EventEmitter<void>();

  readonly slideWidth = SLIDE_WIDTH;
  readonly slideHeight = SLIDE_HEIGHT;
  readonly failedImages = new Set<string>();

  focusedTextBoxId: string | null = null;
  private cachedTextBoxHtml = new Map<string, string>();

  focusedCellKey: string | null = null;
  private cachedTableCellText = new Map<string, string>();

  constructor(
    private elRef: ElementRef<HTMLElement>,
    @Optional() private state?: PptStateService
  ) {}

  focusTextBox(id: string): void {
    this.focusedTextBoxId = id;
    this.cachedTextBoxHtml.set(id, '');
    setTimeout(() => {
      const allTextboxes = this.elRef.nativeElement.querySelectorAll('.ppt-element-textbox');
      let target = this.elRef.nativeElement.querySelector('.ppt-element-textbox.selected') as HTMLElement | null;
      if (!target && allTextboxes.length > 0) {
        target = allTextboxes[allTextboxes.length - 1] as HTMLElement;
      }
      if (target) {
        target.contentEditable = 'true';
        target.focus();
        const sel = window.getSelection();
        if (sel) {
          const range = document.createRange();
          range.selectNodeContents(target);
          range.collapse(true);
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }
    }, 10);
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['selectedElementId']) {
      if (this.focusedTextBoxId && this.focusedTextBoxId !== this.selectedElementId) {
        this.focusedTextBoxId = null;
        this.cachedTextBoxHtml.clear();
      }
    }
  }

  isElementSelected(id: string): boolean {
    if (this.selectedElementId === id) return true;
    return this.selectedElementIds ? this.selectedElementIds.includes(id) : false;
  }

  getTextBoxContent(el: SlideElement): string {
    const raw = (this.asAny(el).html || this.asAny(el).text || '') as string;
    if (this.focusedTextBoxId === el.id) {
      if (!this.cachedTextBoxHtml.has(el.id)) {
        this.cachedTextBoxHtml.set(el.id, raw);
      }
      return this.cachedTextBoxHtml.get(el.id)!;
    }
    this.cachedTextBoxHtml.set(el.id, raw);
    return raw;
  }

  get layoutPlaceholders(): PlaceholderDef[] {
    if (!this.slide) return [];
    const layoutDef = PPT_LAYOUTS[this.slide.layout];
    return layoutDef ? layoutDef.placeholders : [];
  }

  get slideBackgroundStyle(): Record<string, string> {
    if (!this.slide) return { background: '#ffffff' };

    const bg = this.slide.background;
    if (bg) {
      if (bg in COLORS) {
        return { 'background-color': COLORS[bg] };
      }
      if (bg.startsWith('#') || bg.startsWith('rgb')) {
        return { 'background-color': bg };
      }
      if (bg.includes('.')) {
        const url = this.getAssetUrl(bg);
        return {
          'background-image': `url("${url}")`,
          'background-size': 'cover',
          'background-position': 'center',
        };
      }
    }

    return { 'background-color': this.theme?.background || '#ffffff' };
  }

  getPlaceholderText(field: PlaceholderType): string {
    if (!this.slide) return '';
    if (field === 'title') return this.slide.title || '';
    if (field === 'subtitle') return this.slide.subtitle || '';
    if (field === 'body') return this.slide.body || '';
    if (field === 'body2') return this.slide.body2 || '';
    return '';
  }

  getPlaceholderHtml(field: PlaceholderType): string | undefined {
    if (!this.slide) return undefined;
    if (field === 'title') return this.slide.titleHtml;
    if (field === 'subtitle') return this.slide.subtitleHtml;
    if (field === 'body') return this.slide.bodyHtml;
    if (field === 'body2') return this.slide.body2Html;
    return undefined;
  }

  getAssetUrl(file: string): string {
    if (!file) return '';
    const base = this.assetBaseUrl.endsWith('/') ? this.assetBaseUrl : `${this.assetBaseUrl}/`;
    return `${base}${file}`;
  }

  onImageError(file: string): void {
    this.failedImages.add(file);
  }

  hasImageFailed(file: string): boolean {
    return this.failedImages.has(file);
  }

  onPlaceholderTextChange(field: PlaceholderType, payload: { text: string; html: string }): void {
    this.updatePlaceholder.emit({ field, text: payload.text, html: payload.html });
  }

  onTextBoxInput(elementId: string, event: Event): void {
    const el = event.target as HTMLElement;
    const sanitizedHtml = sanitizeRichText(el.innerHTML);
    const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    this.updateTextBoxText.emit({ elementId, text: plain, html: sanitizedHtml });
  }

  onTextBoxPaste(event: ClipboardEvent): void {
    event.preventDefault();
  }

  onBlockClipboard(event: Event): void {
    event.preventDefault();
  }

  onTextBoxFocus(elementId: string, event: FocusEvent): void {
    this.focusedTextBoxId = elementId;
  }

  onTextBoxBlur(elementId: string, event: FocusEvent): void {
    const el = event.target as HTMLElement;
    const sanitizedHtml = sanitizeRichText(el.innerHTML);
    const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();

    const related = event.relatedTarget instanceof Element ? event.relatedTarget : null;
    if (
      related?.closest(
        '.ppt-ribbon, app-ppt-ribbon, .ppt-folded-panel, .ppt-ribbon-overflow-panel, app-ppt-color-picker, .ppt-ribbon-popup'
      )
    ) {
      return;
    }
    this.focusedTextBoxId = null;
    this.cachedTextBoxHtml.delete(elementId);

    // An empty, unfocused text box that loses focus with no text is deleted (like PowerPoint)
    if (!plain) {
      queueMicrotask(() => {
        this.state?.deleteElement(elementId);
      });
      return;
    }

    const storedEl = this.slide.elements?.find((e) => e.id === elementId);
    const storedHtml = storedEl ? ((storedEl as any).html || (storedEl as any).text || '') : '';

    if (sanitizedHtml !== storedHtml) {
      queueMicrotask(() => {
        this.updateTextBoxText.emit({ elementId, text: plain, html: sanitizedHtml });
      });
    }
  }

  onTextBoxKeyDown(event: KeyboardEvent): void {
    if (event.key === 'Escape') {
      event.stopPropagation();
      const el = event.target as HTMLElement;
      const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
      this.focusedTextBoxId = null;
      this.cachedTextBoxHtml.clear();
      el?.blur();
      if (!plain) {
        this.state?.deleteElement(this.selectedElementId || '');
      } else {
        this.deselectElement.emit();
      }
    }
  }

  @HostListener('window:keydown', ['$event'])
  onWindowKeyDown(event: KeyboardEvent): void {
    if (event.key === 'Escape' && this.focusedTextBoxId) {
      this.focusedTextBoxId = null;
      this.cachedTextBoxHtml.clear();
      const activeEl = document.activeElement as HTMLElement | null;
      if (activeEl?.classList.contains('ppt-element-textbox')) {
        activeEl.blur();
      }
      this.deselectElement.emit();
    }
  }

  onElementClick(id: string, event: MouseEvent): void {
    if (!this.isInteractive) return;
    event.stopPropagation();
    this.selectElement.emit(id);

    const el = this.slide.elements?.find((e) => e.id === id);
    if (el?.kind === 'textBox') {
      const target = event.currentTarget as HTMLElement;
      this.focusedTextBoxId = id;
      if (!this.cachedTextBoxHtml.has(id)) {
        this.cachedTextBoxHtml.set(id, (el as any).html || (el as any).text || '');
      }

      target.contentEditable = 'true';
      target.focus();

      const sel = window.getSelection();
      if (sel) {
        let range: Range | null = null;
        if ((document as any).caretRangeFromPoint) {
          range = (document as any).caretRangeFromPoint(event.clientX, event.clientY);
        } else if ((document as any).caretPositionFromPoint) {
          const pos = (document as any).caretPositionFromPoint(event.clientX, event.clientY);
          if (pos && pos.offsetNode) {
            range = document.createRange();
            range.setStart(pos.offsetNode, pos.offset);
            range.collapse(true);
          }
        }

        if (!range || !target.contains(range.startContainer)) {
          range = document.createRange();
          range.selectNodeContents(target);
          range.collapse(false);
        }

        sel.removeAllRanges();
        sel.addRange(range);
      }
    } else {
      this.focusedTextBoxId = null;
      this.cachedTextBoxHtml.clear();
      const target = event.target as HTMLElement | null;
      if (!target?.closest('td') && !target?.isContentEditable) {
        this.focusCanvasHost();
      }
    }
  }

  onElementPointerDown(el: SlideElement, event: PointerEvent): void {
    if (!this.isInteractive) return;
    const target = event.target as HTMLElement | null;
    const isEditable = target && (
      target.isContentEditable ||
      target.tagName === 'INPUT' ||
      target.tagName === 'TEXTAREA' ||
      target.tagName === 'TD' ||
      !!target.closest('[contenteditable="true"]') ||
      !!target.closest('td')
    );
    if (!isEditable && el.kind !== 'textBox') {
      this.focusCanvasHost();
    }
  }

  focusCanvasHost(): void {
    const canvasHost = document.querySelector('.ppt-canvas-area') as HTMLElement | null;
    canvasHost?.focus({ preventScroll: true });
  }

  getShapeFill(el: SlideElement): string {
    return this.asAny(el).fill || this.theme?.accentColors?.[3] || '#0078d4';
  }

  getShapeStroke(el: SlideElement): string {
    const s = this.asAny(el).stroke;
    if (s && s !== 'none') return s;
    let fill = this.asAny(el).fill || this.theme?.accentColors?.[3] || '#0078d4';
    if (fill in COLORS) {
      fill = (COLORS as Record<string, string>)[fill];
    }
    return this.darkenHex(fill, 0.25);
  }

  private darkenHex(hexStr: string, factor = 0.25): string {
    if (!hexStr) return '#005a9f';
    let hex = hexStr.trim().toLowerCase();
    if (hex.startsWith('#')) hex = hex.slice(1);
    if (hex.length === 3) {
      hex = hex.split('').map((c) => c + c).join('');
    }
    const num = parseInt(hex, 16);
    if (isNaN(num) || hex.length !== 6) {
      return '#005a9f';
    }
    const r = Math.max(0, Math.min(255, Math.round(((num >> 16) & 0xff) * (1 - factor))));
    const g = Math.max(0, Math.min(255, Math.round(((num >> 8) & 0xff) * (1 - factor))));
    const b = Math.max(0, Math.min(255, Math.round((num & 0xff) * (1 - factor))));
    const toHex = (n: number) => n.toString(16).padStart(2, '0');
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
  }

  getCellText(el: SlideElement, rIdx: number, cIdx: number): string {
    const key = `${el.id}:${rIdx}:${cIdx}`;
    const stored = (this.asAny(el).cells?.[rIdx]?.[cIdx] ?? '') as string;
    if (this.focusedCellKey === key) {
      if (!this.cachedTableCellText.has(key)) {
        this.cachedTableCellText.set(key, stored);
      }
      return this.cachedTableCellText.get(key)!;
    }
    return stored;
  }

  onTableCellFocus(elementId: string, row: number, col: number): void {
    const key = `${elementId}:${row}:${col}`;
    this.focusedCellKey = key;
    const storedEl = this.slide.elements?.find((e) => e.id === elementId);
    const storedCell = (storedEl as any)?.cells?.[row]?.[col] ?? '';
    this.cachedTableCellText.set(key, storedCell);
  }

  onSlideSurfaceClick(event: MouseEvent): void {
    if (!this.isInteractive) return;
    this.focusedTextBoxId = null;
    this.cachedTextBoxHtml.clear();
    this.focusedCellKey = null;
    this.cachedTableCellText.clear();
    this.deselectElement.emit();
    this.focusCanvasHost();
  }

  onTableCellClick(id: string, event: MouseEvent): void {
    if (!this.isInteractive) return;
    event.stopPropagation();
    if (this.selectedElementId !== id) {
      this.selectElement.emit(id);
    }

    const target = event.target as HTMLElement;
    const td = (target.tagName === 'TD' ? target : target.closest('td')) as HTMLTableCellElement | null;
    if (td && document.activeElement !== td) {
      td.focus();
      const sel = window.getSelection();
      if (sel) {
        let range: Range | null = null;
        if ((document as any).caretRangeFromPoint) {
          range = (document as any).caretRangeFromPoint(event.clientX, event.clientY);
        } else if ((document as any).caretPositionFromPoint) {
          const pos = (document as any).caretPositionFromPoint(event.clientX, event.clientY);
          if (pos && pos.offsetNode) {
            range = document.createRange();
            range.setStart(pos.offsetNode, pos.offset);
            range.collapse(true);
          }
        }
        if (range && td.contains(range.startContainer)) {
          sel.removeAllRanges();
          sel.addRange(range);
        } else {
          range = document.createRange();
          range.selectNodeContents(td);
          range.collapse(false);
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }
    }
  }

  onTableCellInput(elementId: string, row: number, col: number, event: Event): void {
    const target = event.target as HTMLElement;
    const text = target.innerText?.trim() || '';
    this.updateTableCellText.emit({ elementId, row, col, text });
  }

  onTableCellPaste(event: ClipboardEvent): void {
    event.preventDefault();
  }

  onTableCellBlur(elementId: string, row: number, col: number, event: FocusEvent): void {
    const key = `${elementId}:${row}:${col}`;
    if (this.focusedCellKey === key) {
      this.focusedCellKey = null;
    }
    this.cachedTableCellText.delete(key);

    const target = event.target as HTMLElement;
    const text = target.innerText?.trim() || '';

    const storedEl = this.slide.elements?.find((e) => e.id === elementId);
    const storedCell = storedEl ? (storedEl as any).cells?.[row]?.[col] || '' : '';

    if (text !== storedCell) {
      queueMicrotask(() => {
        this.updateTableCellText.emit({ elementId, row, col, text });
      });
    }
  }

  onTableCellKeyDown(event: KeyboardEvent): void {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
      this.focusedCellKey = null;
      this.cachedTableCellText.clear();
    }
  }

  asAny(el: SlideElement): any {
    return el;
  }

  trackByElementId(_: number, el: SlideElement): string {
    return el.id;
  }

  trackByPlaceholderType(_: number, ph: PlaceholderDef): string {
    return ph.type;
  }

  trackByIndex(index: number): number {
    return index;
  }
}
