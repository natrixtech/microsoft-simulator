import { Injectable, signal } from '@angular/core';
import { PptStateService } from './ppt-state.service';
import { PptHistoryService } from './ppt-history.service';
import { PptActionLoggerService } from './ppt-action-logger.service';
import { getNextFontSize, PPT_FONT_SIZES } from '../constants/ppt-fonts';
import { htmlToPlainText, plainTextToHtml, sanitizeRichText } from '../utils/rich-text-sanitize';
import { PlaceholderType, TextBoxSlideElement } from '../models/ppt.types';

@Injectable({
  providedIn: 'root',
})
export class PptTextFormatService {
  readonly isBold = signal<boolean>(false);
  readonly isItalic = signal<boolean>(false);
  readonly isUnderline = signal<boolean>(false);
  readonly currentFont = signal<string>('Arial');
  readonly currentFontSize = signal<number>(24);
  readonly currentAlign = signal<'left' | 'center' | 'right'>('left');
  readonly currentLineSpacing = signal<number>(1.0);
  readonly lastFontColor = signal<string>('#c43e1c');
  readonly hasBullets = signal<boolean>(false);
  readonly hasNumbering = signal<boolean>(false);

  private savedRange: Range | null = null;
  private rootEl: HTMLElement | null = null;

  setRoot(el: HTMLElement): void {
    this.rootEl = el;
  }

  getRoot(): HTMLElement | null {
    return this.rootEl;
  }

  constructor(
    private state: PptStateService,
    private history: PptHistoryService,
    private actionLogger: PptActionLoggerService
  ) {
    if (typeof document !== 'undefined') {
      document.addEventListener('selectionchange', () => {
        this.onSelectionChange();
      });
    }
  }

  saveSelection(): void {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      this.savedRange = sel.getRangeAt(0).cloneRange();
    }
  }

  restoreSelection(): void {
    if (!this.savedRange) return;
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(this.savedRange);
    }
  }

  onSelectionChange(): void {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;

    const activeEl = document.activeElement;
    if (this.rootEl && activeEl && !this.rootEl.contains(activeEl)) {
      return;
    }
    const isInsideEditable =
      activeEl?.classList.contains('ppt-editable-text') ||
      activeEl?.classList.contains('ppt-element-textbox');

    if (!isInsideEditable) {
      // Check selected element in state
      const selectedEl = this.state.selectedElement();
      if (selectedEl?.kind === 'textBox') {
        const tb = selectedEl as TextBoxSlideElement;
        this.isBold.set(!!tb.bold);
        this.isItalic.set(!!tb.italic);
        this.isUnderline.set(!!tb.underline);
        this.currentFont.set(tb.fontFamily || 'Arial');
        this.currentFontSize.set(tb.fontSize || 20);
        this.currentAlign.set(tb.align || 'left');
      }
      return;
    }

    this.savedRange = sel.getRangeAt(0).cloneRange();

    // Query state from document
    try {
      this.isBold.set(document.queryCommandState('bold'));
      this.isItalic.set(document.queryCommandState('italic'));
      this.isUnderline.set(document.queryCommandState('underline'));

      const anchorElement = (sel.anchorNode?.nodeType === Node.ELEMENT_NODE
        ? sel.anchorNode
        : sel.anchorNode?.parentElement) as HTMLElement | null;

      const font = document.queryCommandValue('fontName');
      if (font) {
        const cleanedFont = font.replace(/['"]/g, '');
        this.currentFont.set(cleanedFont);
      } else if (anchorElement) {
        const computed = getComputedStyle(anchorElement);
        const firstFont = (computed.fontFamily || '').split(',')[0].trim().replace(/['"]/g, '');
        if (firstFont) {
          this.currentFont.set(firstFont);
        }
      }

      if (anchorElement) {
        const computed = getComputedStyle(anchorElement);
        const parsedSize = Math.round(parseFloat(computed.fontSize));
        if (parsedSize > 0) {
          this.currentFontSize.set(parsedSize);
        }
      }

      // Check alignment
      if (document.queryCommandState('justifyCenter')) {
        this.currentAlign.set('center');
      } else if (document.queryCommandState('justifyRight')) {
        this.currentAlign.set('right');
      } else {
        this.currentAlign.set('left');
      }

      this.hasBullets.set(document.queryCommandState('insertUnorderedList'));
      this.hasNumbering.set(document.queryCommandState('insertOrderedList'));
    } catch {
      // ignore
    }
  }

  toggleBold(): void {
    this.applyCommand('bold');
    this.isBold.update((v) => !v);
  }

  toggleItalic(): void {
    this.applyCommand('italic');
    this.isItalic.update((v) => !v);
  }

  toggleUnderline(): void {
    this.applyCommand('underline');
    this.isUnderline.update((v) => !v);
  }

  changeFont(font: string): void {
    this.currentFont.set(font);
    this.applyCommand('fontName', font);
  }

  changeFontSize(size: number): void {
    this.currentFontSize.set(size);
    this.applyCommand('fontSize', size);
  }

  increaseFontSize(): void {
    const cur = this.currentFontSize();
    const next = getNextFontSize(cur, 'increase');
    this.changeFontSize(next);
  }

  decreaseFontSize(): void {
    const cur = this.currentFontSize();
    const next = getNextFontSize(cur, 'decrease');
    this.changeFontSize(next);
  }

  setFontColor(color: string): void {
    this.lastFontColor.set(color);
    this.applyCommand('foreColor', color);
  }

  setBullets(): void {
    this.applyCommand('insertUnorderedList');
    this.hasBullets.update((v) => !v);
  }

  setNumbering(): void {
    this.applyCommand('insertOrderedList');
    this.hasNumbering.update((v) => !v);
  }

  setAlign(align: 'left' | 'center' | 'right'): void {
    this.currentAlign.set(align);
    const cmd = align === 'center' ? 'justifyCenter' : align === 'right' ? 'justifyRight' : 'justifyLeft';
    this.applyCommand(cmd, align);
  }

  setLineSpacing(spacing: number): void {
    this.currentLineSpacing.set(spacing);
    this.applyCommand('lineSpacing', spacing);
  }

  private applyCommand(command: string, value?: any): void {
    if (!this.state.isTextActive()) return;

    // Push history once before change
    this.history.push(this.state.getSnapshot());

    // Case 1: Active editable text inside document
    let activeEl = document.activeElement as HTMLElement | null;
    let isInsideEditable =
      activeEl &&
      (!this.rootEl || this.rootEl.contains(activeEl)) &&
      (activeEl.classList.contains('ppt-editable-text') ||
        activeEl.classList.contains('ppt-element-textbox'));

    if (!isInsideEditable) {
      const searchContainer = this.rootEl || document;
      const liveEditor = searchContainer.querySelector(
        '.ppt-editable-text[contenteditable="true"], .ppt-element-textbox[contenteditable="true"]'
      ) as HTMLElement | null;
      if (liveEditor) {
        liveEditor.focus();
        this.restoreSelection();
        this.executeOnActiveElement(liveEditor, command, value);
        this.syncActiveElementToState(liveEditor);
        return;
      }
    }

    if (isInsideEditable && activeEl) {
      this.executeOnActiveElement(activeEl, command, value);
      this.syncActiveElementToState(activeEl);
      return;
    }

    // Case 2: Selected element or placeholder in state (not actively typing)
    this.applyToSelectedTarget(command, value);
  }

  private executeOnActiveElement(el: HTMLElement, command: string, value?: any): void {
    try {
      document.execCommand('styleWithCSS', false, 'true');
    } catch {
      // ignore
    }

    if (command === 'fontSize') {
      const sizePt = Number(value) || 20;
      const sel = window.getSelection();
      if (sel && !sel.isCollapsed) {
        document.execCommand('fontSize', false, '7');
        const fontEls = el.querySelectorAll('font[size="7"], span[style*="xxx-large"]');
        fontEls.forEach((fontEl) => {
          (fontEl as HTMLElement).style.fontSize = `${sizePt}px`;
          fontEl.removeAttribute('size');
        });
      } else {
        el.style.fontSize = `${sizePt}px`;
      }
    } else if (command === 'lineSpacing') {
      const sel = window.getSelection();
      const range = (sel && sel.rangeCount > 0) ? sel.getRangeAt(0) : this.savedRange;
      const blocks = el.querySelectorAll('div, li, p');
      let appliedCount = 0;
      if (range && blocks.length > 0) {
        blocks.forEach((block) => {
          if (range.intersectsNode(block)) {
            (block as HTMLElement).style.lineHeight = String(value);
            appliedCount++;
          }
        });
      }
      if (appliedCount === 0) {
        el.style.lineHeight = String(value);
      }
    } else {
      document.execCommand(command, false, value);
    }
  }

  private syncActiveElementToState(el: HTMLElement): void {
    const sanitized = sanitizeRichText(el.innerHTML);
    const plain = (el.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

    const editingPh = this.state.editingPlaceholder();
    if (editingPh) {
      this.state.updatePlaceholderText(
        this.state.currentSlideIndex(),
        editingPh.field,
        plain,
        sanitized
      );
      this.actionLogger.log('formatText', { command: 'formatPlaceholder' });
      return;
    }

    const selId = this.state.selectedElementId();
    if (selId) {
      this.state.updateElement(selId, { text: plain, html: sanitized });
      this.actionLogger.log('formatText', { command: 'formatTextBox' });
    }
  }

  private applyToSelectedTarget(command: string, value?: any): void {
    const slide = this.state.currentSlide();
    if (!slide) return;

    // A) If a textBox element is selected
    const selId = this.state.selectedElementId();
    const selEl = selId ? slide.elements.find((e) => e.id === selId) : null;
    if (selEl && selEl.kind === 'textBox' && selId) {
      const tb = selEl as TextBoxSlideElement;
      let html = tb.html || plainTextToHtml(tb.text || '', false);
      const updates: Partial<TextBoxSlideElement> = {};

      if (command === 'bold') {
        updates.bold = !tb.bold;
        html = this.wrapOrToggleTag(html, 'b', updates.bold);
      } else if (command === 'italic') {
        updates.italic = !tb.italic;
        html = this.wrapOrToggleTag(html, 'i', updates.italic);
      } else if (command === 'underline') {
        updates.underline = !tb.underline;
        html = this.wrapOrToggleTag(html, 'u', updates.underline);
      } else if (command === 'fontName') {
        updates.fontFamily = String(value);
        html = this.applyStyleToHtml(html, 'font-family', String(value));
      } else if (command === 'fontSize') {
        updates.fontSize = Number(value);
        html = this.applyStyleToHtml(html, 'font-size', `${value}px`);
      } else if (command === 'foreColor') {
        updates.color = String(value);
        html = this.applyStyleToHtml(html, 'color', String(value));
      } else if (['justifyLeft', 'justifyCenter', 'justifyRight'].includes(command)) {
        const align = command === 'justifyCenter' ? 'center' : command === 'justifyRight' ? 'right' : 'left';
        updates.align = align;
        html = this.applyStyleToHtml(html, 'text-align', align);
      } else if (command === 'lineSpacing') {
        html = this.applyStyleToHtml(html, 'line-height', String(value));
      } else if (command === 'insertUnorderedList') {
        html = this.convertToList(html, 'ul');
      } else if (command === 'insertOrderedList') {
        html = this.convertToList(html, 'ol');
      }

      const plain = htmlToPlainText(html) || tb.text;
      updates.text = plain;
      updates.html = sanitizeRichText(html);
      this.state.updateElement(selId, updates);
      this.actionLogger.log('formatText', { command, value });
      return;
    }

    // B) If a placeholder is being edited / selected
    const editingPh = this.state.editingPlaceholder();
    if (editingPh) {
      const field = editingPh.field;
      let rawHtml = '';
      if (field === 'title') rawHtml = slide.titleHtml || plainTextToHtml(slide.title || '', false);
      else if (field === 'subtitle') rawHtml = slide.subtitleHtml || plainTextToHtml(slide.subtitle || '', false);
      else if (field === 'body') rawHtml = slide.bodyHtml || plainTextToHtml(slide.body || '', false);
      else if (field === 'body2') rawHtml = slide.body2Html || plainTextToHtml(slide.body2 || '', false);

      let html = rawHtml;
      if (command === 'bold') {
        const willBeBold = !this.isBold();
        this.isBold.set(willBeBold);
        html = this.wrapOrToggleTag(html, 'b', willBeBold);
      } else if (command === 'italic') {
        const willBeItalic = !this.isItalic();
        this.isItalic.set(willBeItalic);
        html = this.wrapOrToggleTag(html, 'i', willBeItalic);
      } else if (command === 'underline') {
        const willBeUnderline = !this.isUnderline();
        this.isUnderline.set(willBeUnderline);
        html = this.wrapOrToggleTag(html, 'u', willBeUnderline);
      } else if (command === 'fontName') {
        html = this.applyStyleToHtml(html, 'font-family', String(value));
      } else if (command === 'fontSize') {
        html = this.applyStyleToHtml(html, 'font-size', `${value}px`);
      } else if (command === 'foreColor') {
        html = this.applyStyleToHtml(html, 'color', String(value));
      } else if (['justifyLeft', 'justifyCenter', 'justifyRight'].includes(command)) {
        const align = command === 'justifyCenter' ? 'center' : command === 'justifyRight' ? 'right' : 'left';
        html = this.applyStyleToHtml(html, 'text-align', align);
      } else if (command === 'lineSpacing') {
        html = this.applyStyleToHtml(html, 'line-height', String(value));
      } else if (command === 'insertUnorderedList') {
        html = this.convertToList(html, 'ul');
      } else if (command === 'insertOrderedList') {
        html = this.convertToList(html, 'ol');
      }

      const sanitized = sanitizeRichText(html);
      const plain = htmlToPlainText(sanitized);
      this.state.updatePlaceholderText(
        this.state.currentSlideIndex(),
        field,
        plain,
        sanitized
      );
      this.actionLogger.log('formatText', { command, value });
    }
  }

  private wrapOrToggleTag(html: string, tag: 'b' | 'i' | 'u', enable: boolean): string {
    const div = document.createElement('div');
    div.innerHTML = html;
    if (enable) {
      return `<${tag}>${html}</${tag}>`;
    } else {
      const tags = div.querySelectorAll(tag);
      tags.forEach((t) => {
        while (t.firstChild) {
          t.parentNode?.insertBefore(t.firstChild, t);
        }
        t.remove();
      });
      return div.innerHTML;
    }
  }

  private applyStyleToHtml(html: string, styleProp: string, value: string): string {
    const div = document.createElement('div');
    div.innerHTML = html;
    if (styleProp === 'text-align' || styleProp === 'line-height') {
      // Apply to top level block elements or div itself
      const blocks = div.querySelectorAll('div, li, p');
      if (blocks.length > 0) {
        blocks.forEach((b) => ((b as HTMLElement).style as any)[styleProp] = value);
      } else {
        (div.style as any)[styleProp] = value;
        return `<div style="${styleProp}: ${value}">${div.innerHTML}</div>`;
      }
      return div.innerHTML;
    }

    // For font, size, color: wrap inside span
    return `<span style="${styleProp}: ${value}">${html}</span>`;
  }

  private convertToList(html: string, listType: 'ul' | 'ol'): string {
    const plain = htmlToPlainText(html);
    const lines = plain.split('\n').filter((l) => l.trim().length > 0);
    if (lines.length === 0) return '';
    const items = lines.map((l) => `<li>${l}</li>`).join('');
    return `<${listType}>${items}</${listType}>`;
  }
}
