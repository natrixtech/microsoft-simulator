import { Injectable } from '@angular/core';
import {
  ActivityPayload,
  BaseGoal,
  COLORS,
  FONTS,
  clone,
  matches,
  safeAsset,
  sanitizeHtml,
} from 'shared';
import {
  RichEditor,
  findText,
  setSelection,
  styleMatches,
} from './rich-text.service';

export interface PageSetup {
  orientation: 'portrait' | 'landscape';
  paperSize: 'Letter' | 'A4' | 'Legal';
  margins: { top: number; bottom: number; left: number; right: number };
}

export interface WordComment {
  text: string;
  comment: string;
  resolved: boolean;
  replies?: string[];
}

export interface ActionLog {
  action: string;
  text?: string;
}

@Injectable({
  providedIn: 'root',
})
export class WordDocumentService {
  html: string = '';
  setup: PageSetup = {
    orientation: 'portrait',
    paperSize: 'Letter',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
  };
  header: string = '';
  footer: string = '';
  numbering: boolean = false;
  comments: WordComment[] = [];
  styles: Record<string, any> = {};
  actions: ActionLog[] = [];

  editorElement: HTMLElement | null = null;
  rich: RichEditor | null = null;

  init(data: any, activity: ActivityPayload, editorEl: HTMLElement): void {
    this.editorElement = editorEl;
    this.actions = [];

    this.setup = {
      orientation: 'portrait',
      paperSize: 'Letter',
      margins: { top: 1, bottom: 1, left: 1, right: 1 },
      ...clone(data.pageSetup || {}),
    };
    this.setup.margins = {
      top: 1,
      bottom: 1,
      left: 1,
      right: 1,
      ...data.pageSetup?.margins,
    };

    this.header = '';
    this.footer = '';
    this.numbering = false;
    this.styles = {};
    this.comments = clone(data.comments || []).map((x: any) => ({
      ...x,
      resolved: !!x.resolved,
      replies: Array.isArray(x.replies) ? x.replies : [],
    }));

    this.html = sanitizeHtml(data.startingDocument || '', activity.assetBaseUrl);
    editorEl.innerHTML = this.html;

    if (this.rich) {
      this.rich.destroy();
      this.rich = null;
    }

    this.rich = new RichEditor(
      editorEl,
      () => {
        this.html = editorEl.innerHTML;
      },
      (action, text) => {
        this.actions.push({ action, text });
      }
    );
  }

  logAction(action: string, text?: string): void {
    this.actions.push({ action, text });
  }

  check(g: Record<string, any>): boolean {
    if (!this.editorElement) return false;
    const root = this.editorElement;
    const text = root.textContent || '';

    switch (g['type']) {
      case 'textStyle':
        return styleMatches(root, g);

      case 'textContains':
        return (
          (!g['contains'] || g['contains'].every((x: string) => text.includes(x))) &&
          (!g['notContains'] || g['notContains'].every((x: string) => !text.includes(x)))
        );

      case 'textCount':
        return g['text'] ? text.split(g['text']).length - 1 === g['count'] : false;

      case 'textOrder': {
        let i = -1;
        return g['sequence'].every((x: string) => {
          const j = text.indexOf(x, i + 1);
          i = j;
          return j >= 0;
        });
      }

      case 'paragraph': {
        const r = findText(root, g['text']);
        const b = r?.startContainer.parentElement?.closest('p,h1,h2,h3,li,div') as HTMLElement | null;
        if (!b) return false;

        const style =
          b.dataset['paragraphStyle'] ||
          ({ H1: 'h1', H2: 'h2', H3: 'h3' } as Record<string, string>)[b.tagName] ||
          'normal';
        const s = window.getComputedStyle(b);
        const lineSpacingMatch =
          g['lineSpacing'] === undefined ||
          b.style.lineHeight === String(g['lineSpacing']) ||
          parseFloat(b.style.lineHeight) === Number(g['lineSpacing']) ||
          b.dataset['lineSpacing'] === String(g['lineSpacing']);

        return (
          (g['style'] === undefined || style === g['style']) &&
          (g['align'] === undefined ||
            (s.textAlign === 'start' ? 'left' : s.textAlign) === g['align']) &&
          (g['indent'] === undefined ||
            Math.round(parseFloat(s.marginLeft) / 40) === g['indent']) &&
          lineSpacingMatch
        );
      }

      case 'list':
        return g['items'].every((t: string) => {
          const r = findText(root, t);
          const li = r?.startContainer.parentElement?.closest('li');
          if (!li || li.parentElement?.tagName !== (g['listType'] === 'bullet' ? 'UL' : 'OL')) {
            return false;
          }
          let level = 0;
          let p: HTMLElement | null = li.parentElement?.parentElement as HTMLElement | null;
          while (p && p !== root) {
            if (p.tagName === 'UL' || p.tagName === 'OL') level++;
            p = p.parentElement;
          }
          return g['level'] === undefined || g['level'] === level;
        });

      case 'pageSetup':
        return (
          matches(this.setup, g, ['orientation', 'paperSize']) &&
          (!g['margins'] ||
            Object.entries(g['margins']).every(
              ([k, v]) => (this.setup.margins as Record<string, number>)[k] === v
            ))
        );

      case 'table':
        return Array.from(root.querySelectorAll('table')).some((t) => {
          const rowsMatch = g['rows'] === undefined || t.rows.length === g['rows'];
          const colsMatch =
            g['columns'] === undefined ||
            Array.from(t.rows).every((r) => r.cells.length === g['columns']);
          const cellsMatch =
            !g['cells'] ||
            g['cells'].every((r: string[], i: number) =>
              r.every((c: string, j: number) => t.rows[i]?.cells[j]?.textContent === c)
            );
          return rowsMatch && colsMatch && cellsMatch;
        });

      case 'image':
        return Array.from(root.querySelectorAll('img')).some((im) => {
          const fileMatch =
            !g['file'] ||
            im.dataset['file'] === g['file'] ||
            im.src.endsWith('/' + g['file']);
          const alignMatch =
            g['align'] === undefined || (im.dataset['align'] || 'left') === g['align'];
          const wrapMatch =
            g['wrap'] === undefined || (im.dataset['wrap'] || 'inline') === g['wrap'];
          const minW = g['minWidth'] === undefined || im.width >= g['minWidth'];
          const maxW = g['maxWidth'] === undefined || im.width <= g['maxWidth'];
          return fileMatch && alignMatch && wrapMatch && minW && maxW;
        });

      case 'element':
        if (g['element'] === 'header' || g['element'] === 'footer') {
          const val = g['element'] === 'header' ? this.header : this.footer;
          return !!val && (!g['text'] || val.includes(g['text']));
        }
        if (g['element'] === 'pageNumbers') return this.numbering;
        if (
          g['element'] === 'comment' &&
          this.comments.some(
            (c) =>
              (!g['text'] || c.text.includes(g['text']) || c.comment.includes(g['text'])) &&
              (g['resolved'] === undefined || c.resolved === g['resolved'])
          )
        ) {
          return true;
        }
        if (g['element'] === 'tableOfContents') {
          return !!root.querySelector('.table-of-contents, [data-element="tableOfContents"], nav.toc');
        }
        if (g['element'] === 'footnote') {
          return !!root.querySelector('.footnote, [data-element="footnote"], sup.footnote-ref');
        }
        return Array.from(root.querySelectorAll('[data-element]')).some((n: any) => {
          return (
            n.dataset['element'] === g['element'] &&
            (!g['text'] || n.textContent?.includes(g['text'])) &&
            (g['resolved'] === undefined || (n.dataset['resolved'] === 'true') === g['resolved'])
          );
        });

      case 'action':
        return this.actions.some((a) => matches(a, g, ['action', 'text']));

      default:
        return false;
    }
  }

  demo(g: Record<string, any>, activity?: ActivityPayload): void {
    if (!this.editorElement) return;
    const root = this.editorElement;

    switch (g['type']) {
      case 'textStyle': {
        const r = findText(root, g['text']);
        if (r && this.rich) {
          setSelection(r);
          this.rich.capture();
          if (g['bold']) this.rich.command('bold');
          if (g['italic']) this.rich.command('italic');
          if (g['underline']) this.rich.command('underline');
          if (g['font']) this.rich.apply({ fontFamily: g['font'] });
          if (g['size']) this.rich.apply({ fontSize: g['size'] + 'pt' });
          if (g['color']) this.rich.apply({ color: COLORS[g['color']] || g['color'] });
          if (g['highlight'])
            this.rich.apply({ backgroundColor: COLORS[g['highlight']] || g['highlight'] });
        }
        break;
      }

      case 'paragraph': {
        const r = findText(root, g['text']);
        if (r) {
          const b = r.startContainer.parentElement?.closest('p,h1,h2,h3,li,div') as HTMLElement | null;
          if (b) {
            if (g['align']) b.style.textAlign = g['align'];
            if (g['indent'] !== undefined) b.style.marginLeft = g['indent'] * 40 + 'px';
            if (g['lineSpacing'] !== undefined) {
              b.style.lineHeight = String(g['lineSpacing']);
              b.dataset['lineSpacing'] = String(g['lineSpacing']);
            }
            if (g['style']) {
              b.dataset['paragraphStyle'] = g['style'];
              const tagName = (g['style'] === 'h1' || g['style'] === 'h2' || g['style'] === 'h3') ? g['style'] : 'p';
              const rep = document.createElement(tagName);
              rep.innerHTML = b.innerHTML;
              rep.dataset['paragraphStyle'] = g['style'];
              if (g['style'] === 'title') {
                rep.classList.add('doc-title');
              } else if (g['style'] === 'subtitle') {
                rep.classList.add('doc-subtitle');
              }
              b.replaceWith(rep);
            }
          }
        }
        break;
      }

      case 'list': {
        const ul = document.createElement(g['listType'] === 'bullet' ? 'ul' : 'ol');
        (g['items'] || []).forEach((it: string) => {
          const li = document.createElement('li');
          li.textContent = it;
          ul.append(li);
        });
        root.append(ul);
        break;
      }

      case 'table': {
        const rows = g['rows'] || 2;
        const cols = g['columns'] || 2;
        const tbl = document.createElement('table');
        tbl.className = 'doc-table';
        for (let i = 0; i < rows; i++) {
          const tr = tbl.insertRow();
          for (let j = 0; j < cols; j++) {
            const td = tr.insertCell();
            td.textContent = g['cells']?.[i]?.[j] || 'Cell';
          }
        }
        root.append(tbl);
        break;
      }

      case 'image': {
        const img = document.createElement('img');
        img.src = safeAsset(activity?.assetBaseUrl || '', g['file'] || 'photo.png');
        img.dataset['file'] = g['file'];
        if (g['align']) img.dataset['align'] = g['align'];
        if (g['wrap']) img.dataset['wrap'] = g['wrap'];
        if (g['minWidth']) img.width = g['minWidth'];
        root.append(img);
        break;
      }

      case 'pageSetup': {
        if (g['orientation']) this.setup.orientation = g['orientation'];
        if (g['paperSize']) this.setup.paperSize = g['paperSize'];
        if (g['margins']) Object.assign(this.setup.margins, g['margins']);
        break;
      }

      case 'element': {
        if (g['element'] === 'header') this.header = g['text'] || 'Header';
        else if (g['element'] === 'footer') this.footer = g['text'] || 'Footer';
        else if (g['element'] === 'pageNumbers') this.numbering = true;
        else if (g['element'] === 'tableOfContents') {
          const toc = document.createElement('div');
          toc.className = 'table-of-contents';
          toc.dataset['element'] = 'tableOfContents';
          toc.innerHTML = '<h3 class="toc-title">Table of Contents</h3><ul class="toc-list"><li>Heading 1</li></ul>';
          root.prepend(toc);
        } else if (g['element'] === 'footnote') {
          const fn = document.createElement('div');
          fn.className = 'footnote';
          fn.dataset['element'] = 'footnote';
          fn.innerHTML = '<hr><p><sup>1</sup> Footnote text</p>';
          root.append(fn);
        } else {
          const el = document.createElement('div');
          el.dataset['element'] = g['element'];
          el.textContent = g['text'] || g['element'];
          root.append(el);
        }
        break;
      }

      case 'textContains': {
        let currentHtml = root.innerHTML;
        const notList = (g['notContains'] || []) as string[];
        const conList = (g['contains'] || []) as string[];
        notList.forEach((wrong, idx) => {
          const replacement = conList[idx] || conList[0] || '';
          if (wrong && replacement) {
            currentHtml = currentHtml.split(wrong).join(replacement);
          }
        });
        root.innerHTML = currentHtml;
        const text = root.textContent || '';
        conList.forEach((req) => {
          if (!text.includes(req)) {
            const p = document.createElement('p');
            p.textContent = req;
            root.append(p);
          }
        });
        break;
      }

      case 'textCount': {
        const target = g['text'];
        const needed = g['count'];
        if (target && typeof needed === 'number') {
          let text = root.textContent || '';
          let currentCount = text.split(target).length - 1;
          while (currentCount < needed) {
            const p = document.createElement('p');
            p.textContent = target;
            root.append(p);
            currentCount++;
          }
          if (currentCount > needed) {
            let html = root.innerHTML;
            while (currentCount > needed && html.includes(target)) {
              html = html.replace(target, '');
              currentCount--;
            }
            root.innerHTML = html;
          }
        }
        break;
      }

      case 'textOrder': {
        const seq = (g['sequence'] || []) as string[];
        if (seq.length > 0) {
          const p = document.createElement('p');
          p.textContent = seq.join(' ');
          root.append(p);
        }
        break;
      }

      case 'action':
        this.logAction(g['action'], g['text']);
        break;
    }
  }

  snapshot(): Record<string, any> {
    return {
      html: this.editorElement ? this.editorElement.innerHTML : this.html,
      setup: clone(this.setup),
      header: this.header,
      footer: this.footer,
      numbering: this.numbering,
      comments: clone(this.comments),
      styles: clone(this.styles),
      actions: clone(this.actions),
    };
  }

  restore(s: Record<string, any>): void {
    this.html = s['html'] || '';
    if (this.editorElement) {
      this.editorElement.innerHTML = this.html;
    }
    this.setup = clone(s['setup']);
    this.header = s['header'] || '';
    this.footer = s['footer'] || '';
    this.numbering = !!s['numbering'];
    this.comments = clone(s['comments'] || []);
    this.styles = clone(s['styles'] || {});
    this.actions = clone(s['actions'] || []);
  }
}
