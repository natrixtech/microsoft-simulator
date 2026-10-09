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

    const actual: Record<string, any> = {
      bold: Number(s.fontWeight) >= 600 || s.fontWeight === 'bold',
      italic: s.fontStyle === 'italic',
      underline: deco(),
      font: s.fontFamily.replaceAll('"', '').split(',')[0].trim(),
      size: parseFloat(s.fontSize) * 0.75,
      color: s.color,
      highlight: s.backgroundColor,
      align: s.textAlign === 'start' ? 'left' : s.textAlign,
    };

    for (const k of ['bold', 'italic', 'underline', 'font', 'size', 'align']) {
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

let internalClipboard = '';

export class RichEditor {
  private saved: Range | null = null;
  private lastPasteTime = 0;
  private lastCopyTime = 0;

  private onInput = () => this.changed();
  private onMouseUp = () => this.capture();
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
        if (e.shiftKey) this.command('redo');
        else this.command('undo');
      } else if (key === 'y') {
        e.preventDefault();
        this.command('redo');
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
    this.root.dispatchEvent(new Event('beforeformat'));
    this.restore();
    document.execCommand(cmd, false, value);
    this.capture();
    this.changed();
  }

  apply(styles: Record<string, string>): void {
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
