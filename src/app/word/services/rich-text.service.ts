import { Injectable } from '@angular/core';
import { COLORS } from 'shared';

export function findText(root: HTMLElement, text: string): Range | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: { n: Text; start: number }[] = [];
  let n: Text | null;
  let all = '';

  while ((n = walker.nextNode() as Text | null)) {
    nodes.push({ n, start: all.length });
    all += n.textContent || '';
  }

  const index = all.indexOf(text);
  if (index < 0 || !text) return null;

  const start = nodes.find((x) => x.start + x.n.length > index);
  const end = nodes.find((x) => x.start + x.n.length >= index + text.length);
  if (!start || !end) return null;

  const r = document.createRange();
  r.setStart(start.n, index - start.start);
  r.setEnd(end.n, index + text.length - end.start);
  return r;
}

export function setSelection(r: Range | null): void {
  if (!r) return;
  const s = window.getSelection();
  if (s) {
    s.removeAllRanges();
    s.addRange(r);
  }
}

export function styleMatches(root: HTMLElement, g: Record<string, any>): boolean {
  const r = findText(root, g['text']);
  if (!r) return false;

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n: Text | null;

  while ((n = walker.nextNode() as Text | null)) {
    if (!n.textContent || !r.intersectsNode(n)) continue;

    const intersects =
      (n !== r.startContainer || r.startOffset < n.length) &&
      (n !== r.endContainer || r.endOffset > 0);
    if (!intersects) continue;

    const parent = n.parentElement;
    if (!parent) continue;

    const s = window.getComputedStyle(parent);
    const deco = (): boolean => {
      let e: HTMLElement | null = parent;
      while (e && e !== root.parentElement) {
        if (window.getComputedStyle(e).textDecorationLine.includes('underline')) {
          return true;
        }
        e = e.parentElement;
      }
      return false;
    };

    const isSuper = (): boolean => {
      let e: HTMLElement | null = parent;
      while (e && e !== root.parentElement) {
        if (e.tagName === 'SUP' || window.getComputedStyle(e).verticalAlign === 'super') {
          return true;
        }
        e = e.parentElement;
      }
      return false;
    };

    const isSub = (): boolean => {
      let e: HTMLElement | null = parent;
      while (e && e !== root.parentElement) {
        if (e.tagName === 'SUB' || window.getComputedStyle(e).verticalAlign === 'sub') {
          return true;
        }
        e = e.parentElement;
      }
      return false;
    };

    const actual: Record<string, any> = {
      bold: Number(s.fontWeight) >= 600 || s.fontWeight === 'bold',
      italic: s.fontStyle === 'italic',
      underline: deco(),
      superscript: isSuper(),
      subscript: isSub(),
      font: s.fontFamily.replaceAll('"', '').split(',')[0].trim(),
      size: parseFloat(s.fontSize) * 0.75,
      color: s.color,
      highlight: s.backgroundColor,
      align: s.textAlign === 'start' ? 'left' : s.textAlign,
    };

    for (const k of ['bold', 'italic', 'underline', 'superscript', 'subscript', 'font', 'size', 'align']) {
      if (g[k] !== undefined) {
        if (k === 'size' ? Math.abs(actual[k] - g[k]) > 0.15 : actual[k] !== g[k]) {
          return false;
        }
      }
    }

    for (const k of ['color', 'highlight']) {
      if (g[k] !== undefined) {
        const probe = document.createElement('span');
        probe.style.color = COLORS[g[k]] || g[k];
        root.append(probe);
        const wanted = window.getComputedStyle(probe).color;
        probe.remove();
        if (actual[k] !== wanted) return false;
      }
    }
  }

  return true;
}

function coversBlockText(range: Range, block: HTMLElement): boolean {
  if (range.collapsed) return false;
  try {
    const blockRange = document.createRange();
    blockRange.selectNodeContents(block);
    if (range.compareBoundaryPoints(Range.END_TO_START, blockRange) >= 0) return false;
    if (range.compareBoundaryPoints(Range.START_TO_END, blockRange) <= 0) return false;

    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let node: Node | null = walker.nextNode();
    while (node) {
      const len = node.textContent?.length || 0;
      if (len > 0) {
        const nr = document.createRange();
        nr.selectNodeContents(node);
        if (
          range.compareBoundaryPoints(Range.END_TO_START, nr) < 0 &&
          range.compareBoundaryPoints(Range.START_TO_END, nr) > 0
        ) {
          return true;
        }
      }
      node = walker.nextNode();
    }
    return false;
  } catch {
    return false;
  }
}

function isWholeBlock(range: Range, block: HTMLElement): boolean {
  try {
    const blockRange = document.createRange();
    blockRange.selectNodeContents(block);
    const startOk = range.compareBoundaryPoints(Range.START_TO_START, blockRange) <= 0;
    const endOk = range.compareBoundaryPoints(Range.END_TO_END, blockRange) >= 0;
    return startOk && endOk;
  } catch {
    return false;
  }
}

let internalClipboard = '';

export class RichEditor {
  saved: Range | null = null;
  private lastPasteTime = 0;
  private lastCopyTime = 0;
  taskStartHtml: string | null = null;
  getDocumentHtml?: () => string;
  customUndo?: () => boolean;

  private onInput = () => this.changed();
  private onMouseUp = () => {
    this.capture();
    setTimeout(() => this.capture(), 0);
  };
  private onKeyUp = () => this.capture();
  private onKeyDown = (e: KeyboardEvent) => {
    if (e.ctrlKey || e.metaKey) {
      const key = e.key.toLowerCase();
      if (key === 'b') {
        e.preventDefault();
        this.command('bold');
      } else if (key === 'i') {
        e.preventDefault();
        this.command('italic');
      } else if (key === 'u') {
        e.preventDefault();
        this.command('underline');
      } else if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) {
          this.command('redo');
          this.action('redo');
        } else {
          const currentHtml = this.getDocumentHtml ? this.getDocumentHtml() : this.root.innerHTML;
          if (this.taskStartHtml !== null && currentHtml === this.taskStartHtml) {
            return;
          }
          if (this.customUndo && this.customUndo()) {
            return;
          }
          this.command('undo');
          this.action('undo');
        }
      } else if (key === 'y') {
        e.preventDefault();
        this.command('redo');
        this.action('redo');
      } else if (key === 'a') {
        e.preventDefault();
        const r = document.createRange();
        r.selectNodeContents(this.root);
        setSelection(r);
        this.capture();
        this.action('select', this.root.innerText || '');
      } else if (key === 'c') {
        e.preventDefault();
        this.copy(false);
      } else if (key === 'x') {
        e.preventDefault();
        this.copy(true);
      } else if (key === 'v') {
        e.preventDefault();
        this.paste();
      }
    } else if (e.key === 'Tab') {
      const s = window.getSelection();
      const node = s?.anchorNode;
      const el = (node ? (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement) : null) as HTMLElement | null;
      const cell = el?.closest('td, th') as HTMLTableCellElement | null;
      if (cell) {
        e.preventDefault();
        const row = cell.closest('tr');
        const table = cell.closest('table');
        if (row && table) {
          const cells = Array.from(row.cells);
          const cellIdx = cells.indexOf(cell);
          if (!e.shiftKey) {
            if (cellIdx + 1 < cells.length) {
              const nextCell = cells[cellIdx + 1];
              const r = document.createRange();
              r.selectNodeContents(nextCell);
              r.collapse(true);
              setSelection(r);
              this.capture();
              this.action('nextCell');
            } else {
              const rows = Array.from(table.rows);
              const rowIdx = rows.indexOf(row);
              if (rowIdx + 1 < rows.length) {
                const firstCell = rows[rowIdx + 1].cells[0];
                const r = document.createRange();
                r.selectNodeContents(firstCell);
                r.collapse(true);
                setSelection(r);
                this.capture();
                this.action('nextCell');
              } else {
                const newRow = table.insertRow();
                for (let i = 0; i < cells.length; i++) {
                  const c = newRow.insertCell();
                  c.innerHTML = '<br>';
                }
                const r = document.createRange();
                r.selectNodeContents(newRow.cells[0]);
                r.collapse(true);
                setSelection(r);
                this.capture();
                this.changed();
                this.action('tabNewRow');
              }
            }
          }
        }
      } else {
        const inLi = el?.closest('li');
        if (inLi) {
          e.preventDefault();
          this.command(e.shiftKey ? 'outdent' : 'indent');
        }
      }
    }
  };

  private onCopy = (e: ClipboardEvent) => {
    e.preventDefault();
    this.copy(false);
  };

  private onCut = (e: ClipboardEvent) => {
    e.preventDefault();
    this.copy(true);
  };

  private onPaste = (e: ClipboardEvent) => {
    e.preventDefault();
    this.paste();
  };

  constructor(
    private root: HTMLElement,
    private changed: () => void = () => {},
    private action: (act: string, txt?: string) => void = () => {}
  ) {
    this.root.addEventListener('input', this.onInput);
    this.root.addEventListener('mouseup', this.onMouseUp);
    this.root.addEventListener('keyup', this.onKeyUp);
    this.root.addEventListener('keydown', this.onKeyDown);
    this.root.addEventListener('copy', this.onCopy);
    this.root.addEventListener('cut', this.onCut);
    this.root.addEventListener('paste', this.onPaste);
  }

  destroy(): void {
    this.root.removeEventListener('input', this.onInput);
    this.root.removeEventListener('mouseup', this.onMouseUp);
    this.root.removeEventListener('keyup', this.onKeyUp);
    this.root.removeEventListener('keydown', this.onKeyDown);
    this.root.removeEventListener('copy', this.onCopy);
    this.root.removeEventListener('cut', this.onCut);
    this.root.removeEventListener('paste', this.onPaste);
  }

  capture(): void {
    const s = window.getSelection();
    if (
      s &&
      s.rangeCount &&
      s.anchorNode &&
      s.focusNode &&
      this.root.contains(s.anchorNode) &&
      this.root.contains(s.focusNode)
    ) {
      this.saved = s.getRangeAt(0).cloneRange();
      if (s.toString()) this.action('select', s.toString());
    }
  }

  restore(): void {
    this.root.focus();
    if (
      this.saved &&
      this.saved.startContainer &&
      this.saved.endContainer &&
      this.root.contains(this.saved.startContainer) &&
      this.root.contains(this.saved.endContainer)
    ) {
      setSelection(this.saved);
    }
  }

  command(cmd: string, value?: string): void {
    if (cmd === 'indent' || cmd === 'outdent') {
      this.indent(cmd === 'indent');
      return;
    }
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    try {
      document.execCommand('styleWithCSS', false, 'true');
    } catch {}
    document.execCommand(cmd, false, value);
    this.capture();
    this.changed();
  }

  indent(increase: boolean = true): void {
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    const sel = window.getSelection();

    // In a list item: delegate directly to browser indent/outdent to nest sub-lists and adjust levels
    let inListItem = false;
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      const getLi = (n: Node | null): HTMLElement | null => {
        if (!n) return null;
        return (n.nodeType === Node.ELEMENT_NODE ? (n as Element) : n.parentElement)?.closest('li') as HTMLElement | null;
      };
      if (getLi(range.startContainer) || getLi(range.endContainer)) {
        inListItem = true;
      }
    }
    if (!inListItem) {
      const b = this.block();
      if (b?.closest('li')) {
        inListItem = true;
      }
    }

    if (inListItem) {
      try {
        document.execCommand('styleWithCSS', false, 'true');
      } catch {}
      document.execCommand(increase ? 'indent' : 'outdent');
      this.capture();
      this.changed();
      return;
    }

    // Normal blocks (p, h1-h3) not in a list and not wrapping a list
    let blocks: HTMLElement[] = [];
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      const all = Array.from(this.root.querySelectorAll<HTMLElement>('p, h1, h2, h3'));
      blocks = all.filter((el) => {
        if (el.closest('li, ul, ol') || el.querySelector('ul, ol, li')) return false;
        try {
          return coversBlockText(range, el);
        } catch {
          return false;
        }
      });
      if (blocks.length === 0) {
        const b = this.block();
        if (b && ['P', 'H1', 'H2', 'H3'].includes(b.tagName) && !b.closest('li, ul, ol') && !b.querySelector('ul, ol, li')) {
          blocks = [b];
        }
      }
    } else {
      const b = this.block();
      if (b && ['P', 'H1', 'H2', 'H3'].includes(b.tagName) && !b.closest('li, ul, ol') && !b.querySelector('ul, ol, li')) {
        blocks = [b];
      }
    }

    for (const b of blocks) {
      const cur = parseFloat(b.style.marginLeft) || 0;
      const next = increase ? cur + 40 : Math.max(0, cur - 40);
      b.style.marginLeft = next > 0 ? `${next}px` : '';
    }
    this.capture();
    this.changed();
  }

  private getCoveredBlocks(): HTMLElement[] {
    if (!this.saved || this.saved.collapsed) return [];
    const all = Array.from(this.root.querySelectorAll<HTMLElement>('p, h1, h2, h3, li'));
    return all.filter((el) => coversBlockText(this.saved!, el));
  }

  private getTargetWholeBlocks(): HTMLElement[] | null {
    if (!this.saved || this.saved.collapsed) return null;
    const covered = this.getCoveredBlocks();
    if (covered.length === 0) return null;

    if (covered.every((b) => isWholeBlock(this.saved!, b))) {
      return covered;
    }

    if (covered.length === 1) {
      const b = covered[0];
      const selText = this.saved.toString().trim();
      const blockText = (b.textContent || '').trim();
      if (selText && blockText && selText === blockText) {
        return [b];
      }
    }

    return null;
  }

  private applyToWholeBlocks(styles: Record<string, string>, targetBlocks: HTMLElement[]): void {
    this.root.dispatchEvent(new Event('beforeformat'));
    for (const blk of targetBlocks) {
      const span = document.createElement('span');
      Object.assign(span.style, styles);
      while (blk.firstChild) {
        span.appendChild(blk.firstChild);
      }
      blk.appendChild(span);
    }
    const newRange = document.createRange();
    newRange.setStart(targetBlocks[0], 0);
    newRange.setEnd(targetBlocks[targetBlocks.length - 1], targetBlocks[targetBlocks.length - 1].childNodes.length);
    setSelection(newRange);
    this.saved = newRange;
    this.changed();
  }

  private applyToMultiBlocks(styles: Record<string, string>): void {
    if (!this.saved) return;
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);

    // 1. Split text nodes at range boundaries so selected portions are whole text nodes
    if (range.endContainer.nodeType === Node.TEXT_NODE) {
      const endText = range.endContainer as Text;
      if (range.endOffset > 0 && range.endOffset < (endText.textContent?.length || 0)) {
        endText.splitText(range.endOffset);
        range.setEnd(endText, endText.textContent?.length || range.endOffset);
      }
    }

    if (range.startContainer.nodeType === Node.TEXT_NODE) {
      const startText = range.startContainer as Text;
      if (range.startOffset > 0 && range.startOffset < (startText.textContent?.length || 0)) {
        const after = startText.splitText(range.startOffset);
        range.setStart(after, 0);
      }
    }

    // 2. Collect all text nodes strictly within range
    const walker = document.createTreeWalker(this.root, NodeFilter.SHOW_TEXT);
    const textNodes: Text[] = [];
    let curr: Node | null = walker.nextNode();
    while (curr) {
      if (curr.nodeType === Node.TEXT_NODE && curr.textContent && curr.textContent.length > 0) {
        const nr = document.createRange();
        nr.selectNodeContents(curr);
        if (
          range.compareBoundaryPoints(Range.START_TO_START, nr) <= 0 &&
          range.compareBoundaryPoints(Range.END_TO_END, nr) >= 0
        ) {
          textNodes.push(curr as Text);
        }
      }
      curr = walker.nextNode();
    }

    if (textNodes.length === 0) return;

    // 3. Wrap each selected text node in a styled span inside its own block
    const createdSpans: HTMLSpanElement[] = [];
    for (const tn of textNodes) {
      const span = document.createElement('span');
      Object.assign(span.style, styles);
      tn.parentNode?.insertBefore(span, tn);
      span.appendChild(tn);
      createdSpans.push(span);
    }

    // 4. Update selection across all newly created spans
    if (createdSpans.length > 0) {
      const newRange = document.createRange();
      newRange.setStartBefore(createdSpans[0]);
      newRange.setEndAfter(createdSpans[createdSpans.length - 1]);
      setSelection(newRange);
      this.saved = newRange;
    }
    this.capture();
    this.changed();
  }

  applyColor(color: string): void {
    const whole = this.getTargetWholeBlocks();
    if (whole) {
      this.applyToWholeBlocks({ color }, whole);
      return;
    }
    const covered = this.getCoveredBlocks();
    if (covered.length >= 2) {
      this.applyToMultiBlocks({ color });
      return;
    }
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    try {
      document.execCommand('styleWithCSS', false, 'true');
    } catch {}
    document.execCommand('foreColor', false, color || '#000000');
    this.capture();
    this.changed();
  }

  applyHighlight(color: string): void {
    const whole = this.getTargetWholeBlocks();
    if (whole) {
      this.applyToWholeBlocks({ backgroundColor: color }, whole);
      return;
    }
    const covered = this.getCoveredBlocks();
    if (covered.length >= 2) {
      this.applyToMultiBlocks({ backgroundColor: color });
      return;
    }
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    try {
      document.execCommand('styleWithCSS', false, 'true');
    } catch {}
    const cmd = document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
    document.execCommand(cmd, false, color || 'transparent');
    this.capture();
    this.changed();
  }

  apply(styles: Record<string, string>): void {
    const whole = this.getTargetWholeBlocks();
    if (whole) {
      this.applyToWholeBlocks(styles, whole);
      return;
    }
    const covered = this.getCoveredBlocks();
    if (covered.length >= 2) {
      this.applyToMultiBlocks(styles);
      return;
    }
    if (styles['color'] && Object.keys(styles).length === 1) {
      this.applyColor(styles['color']);
      return;
    }
    if (styles['backgroundColor'] && Object.keys(styles).length === 1) {
      this.applyHighlight(styles['backgroundColor']);
      return;
    }
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    if (!this.saved) return;

    const span = document.createElement('span');
    Object.assign(span.style, styles);

    try {
      span.append(this.saved.extractContents());
      this.saved.insertNode(span);
      const r = document.createRange();
      r.selectNodeContents(span);
      setSelection(r);
      this.saved = r;
      this.changed();
    } catch {}
  }

  copy(cut: boolean = false): void {
    const now = Date.now();
    if (now - this.lastCopyTime < 40) return;
    this.lastCopyTime = now;

    this.capture();
    const s = window.getSelection();
    const str = s ? s.toString() : '';
    if (str) {
      internalClipboard = str;
      this.action(cut ? 'cut' : 'copy', internalClipboard);
    }
    if (cut) this.command('delete');
  }

  paste(): void {
    const now = Date.now();
    if (now - this.lastPasteTime < 60) return;
    this.lastPasteTime = now;

    if (!internalClipboard) return;
    this.capture();
    this.restore();
    document.execCommand('insertText', false, internalClipboard);
    this.capture();
    this.changed();
    this.action('paste', internalClipboard);
  }

  block(): HTMLElement | null {
    this.restore();
    const s = window.getSelection();
    const n = s ? s.anchorNode : null;
    if (!n) return null;
    const el = (n.nodeType === Node.TEXT_NODE ? n.parentElement : n) as HTMLElement | null;
    return el ? el.closest('p,h1,h2,h3,li,div') : null;
  }
}
