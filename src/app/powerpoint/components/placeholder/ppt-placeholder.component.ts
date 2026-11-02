import {
  AfterViewInit,
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PlaceholderDef, ThemeDef } from '../../models/ppt.types';
import { plainTextToHtml, sanitizeRichText } from '../../utils/rich-text-sanitize';
import { PptSafeHtmlPipe } from '../../pipes/ppt-safe-html.pipe';

@Component({
  selector: 'app-ppt-placeholder',
  standalone: true,
  imports: [CommonModule, PptSafeHtmlPipe],
  templateUrl: './ppt-placeholder.component.html',
  styleUrls: ['./ppt-placeholder.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptPlaceholderComponent implements OnChanges, AfterViewInit {
  @ViewChild('editableRef', { static: false }) editableRef?: ElementRef<HTMLDivElement>;

  @Input() placeholder!: PlaceholderDef;
  @Input() text: string = '';
  @Input() html?: string;
  @Input() theme!: ThemeDef;
  @Input() isEditing: boolean = false;
  @Input() isInteractive: boolean = true;

  @Output() readonly textChange = new EventEmitter<{ text: string; html: string }>();
  @Output() readonly startEdit = new EventEmitter<void>();
  @Output() readonly stopEdit = new EventEmitter<void>();

  get isEmpty(): boolean {
    return !this.text || this.text.trim().length === 0;
  }

  get renderedHtml(): string {
    if (this.html) {
      return this.html;
    }
    return this.text ? plainTextToHtml(this.text, false) : '';
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isEditing'] && this.isEditing) {
      setTimeout(() => this.focusAndPopulateEditable(), 10);
    }
    if ((changes['text'] || changes['html']) && this.editableRef?.nativeElement && !this.isEditing) {
      this.editableRef.nativeElement.innerHTML = this.renderedHtml;
    }
  }

  ngAfterViewInit(): void {
    if (this.isEditing) {
      this.focusAndPopulateEditable();
    }
  }

  private pendingClickPoint: { x: number; y: number } | null = null;
  private pendingSelection: { start: number; end: number } | null = null;

  private textOffsetIn(root: Node, node: Node, offset: number): number {
    const r = document.createRange();
    r.selectNodeContents(root);
    r.setEnd(node, offset);
    return r.toString().length;
  }

  private rangeFromOffsets(root: HTMLElement, start: number, end: number): Range | null {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const r = document.createRange();
    let pos = 0;
    let startSet = false;
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const len = (n.textContent || '').length;
      if (!startSet && start <= pos + len) {
        r.setStart(n, start - pos);
        startSet = true;
      }
      if (startSet && end <= pos + len) {
        r.setEnd(n, end - pos);
        return r;
      }
      pos += len;
    }
    return null;
  }

  onPlaceholderClick(event: MouseEvent): void {
    if (!this.isInteractive) return;
    event.stopPropagation();
    if (!this.isEditing) {
      const box = event.currentTarget as HTMLElement | null;
      const rendered = box?.querySelector('.ppt-rendered-text');
      const sel = window.getSelection();
      if (
        rendered &&
        sel &&
        !sel.isCollapsed &&
        sel.rangeCount > 0
      ) {
        const r = sel.getRangeAt(0);
        if (rendered.contains(r.startContainer) && rendered.contains(r.endContainer)) {
          const start = this.textOffsetIn(rendered, r.startContainer, r.startOffset);
          const end = this.textOffsetIn(rendered, r.endContainer, r.endOffset);
          this.pendingSelection = { start, end };
          this.pendingClickPoint = null;
        } else {
          this.pendingSelection = null;
          this.pendingClickPoint = { x: event.clientX, y: event.clientY };
        }
      } else {
        this.pendingSelection = null;
        this.pendingClickPoint = { x: event.clientX, y: event.clientY };
      }
      this.startEdit.emit();
    }
  }

  onPlaceholderDblClick(event: MouseEvent): void {
    if (!this.isInteractive) return;
    setTimeout(() => {
      const el = this.editableRef?.nativeElement;
      const sel = window.getSelection();
      if (!el || !sel) return;
      let r: Range | null = null;
      if ((document as any).caretRangeFromPoint) {
        r = (document as any).caretRangeFromPoint(event.clientX, event.clientY);
      } else if ((document as any).caretPositionFromPoint) {
        const pos = (document as any).caretPositionFromPoint(event.clientX, event.clientY);
        if (pos && pos.offsetNode) {
          r = document.createRange();
          r.setStart(pos.offsetNode, pos.offset);
          r.collapse(true);
        }
      }
      if (!r || !el.contains(r.startContainer)) return;
      sel.removeAllRanges();
      sel.addRange(r);
      (sel as any).modify('move', 'backward', 'word');
      (sel as any).modify('extend', 'forward', 'word');
    }, 20);
  }

  onInput(event: Event): void {
    const el = event.target as HTMLDivElement;
    const sanitizedHtml = sanitizeRichText(el.innerHTML);
    const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    this.textChange.emit({ text: plain, html: sanitizedHtml });
  }

  onPaste(event: ClipboardEvent): void {
    event.preventDefault();
  }

  onCut(event: ClipboardEvent): void {
    event.preventDefault();
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
  }

  onBlur(e?: FocusEvent): void {
    if (this.isEditing) {
      const el = this.editableRef?.nativeElement;
      if (el) {
        const sanitizedHtml = sanitizeRichText(el.innerHTML);
        const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        this.textChange.emit({ text: plain, html: sanitizedHtml });
      }
      const related = e?.relatedTarget as HTMLElement | null;
      if (
        related?.closest(
          '.ppt-ribbon, app-ppt-ribbon, .ppt-folded-panel, .ppt-ribbon-overflow-panel, app-ppt-color-picker, .ppt-ribbon-popup'
        )
      ) {
        return;
      }
      this.stopEdit.emit();
    }
  }

  onKeyDown(event: KeyboardEvent): void {
    if (event.key === 'Escape') {
      this.onBlur();
    }
  }

  private focusAndPopulateEditable(): void {
    const el = this.editableRef?.nativeElement;
    if (!el) return;
    el.innerHTML = this.renderedHtml;
    el.focus();

    if (this.pendingSelection) {
      const { start, end } = this.pendingSelection;
      this.pendingSelection = null;
      this.pendingClickPoint = null;
      const sel = window.getSelection();
      if (sel) {
        const range = this.rangeFromOffsets(el, start, end);
        if (range) {
          sel.removeAllRanges();
          sel.addRange(range);
          return;
        }
      }
    }

    const point = this.pendingClickPoint;
    this.pendingClickPoint = null;

    const sel = window.getSelection();
    if (sel) {
      let range: Range | null = null;
      if (point) {
        if ((document as any).caretRangeFromPoint) {
          range = (document as any).caretRangeFromPoint(point.x, point.y);
        } else if ((document as any).caretPositionFromPoint) {
          const pos = (document as any).caretPositionFromPoint(point.x, point.y);
          if (pos && pos.offsetNode) {
            range = document.createRange();
            range.setStart(pos.offsetNode, pos.offset);
            range.collapse(true);
          }
        }
      }

      if (
        !range ||
        !el.contains(range.startContainer) ||
        (range.startContainer === el && range.startOffset === 0 && el.innerText.trim().length > 0)
      ) {
        range = document.createRange();
        range.selectNodeContents(el);
        range.collapse(false);
      }

      sel.removeAllRanges();
      sel.addRange(range);
    }
  }
}
