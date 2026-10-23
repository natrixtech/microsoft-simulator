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
import {
  DEFAULT_CHART_DATA,
  WordChartType,
  renderChartSvg,
} from '../chart-svg.helper';

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

export interface CustomStyleDefinition {
  name: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font?: string;
  size?: number;
  color?: string;
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
  customStyles: CustomStyleDefinition[] = [];
  actions: ActionLog[] = [];

  editorElement: HTMLElement | null = null;
  page2Element: HTMLElement | null = null;
  hasPageBreak: boolean = false;
  rich: RichEditor | null = null;

  init(data: any, activity: ActivityPayload, editorEl: HTMLElement): void {
    this.editorElement = editorEl;
    this.page2Element = null;
    this.hasPageBreak = false;
    this.actions = [];

    const effectivePageSetup = data?.pageSetup || activity?.pageSetup || {};
    this.setup = {
      orientation: 'portrait',
      paperSize: 'Letter',
      margins: { top: 1, bottom: 1, left: 1, right: 1 },
      ...clone(effectivePageSetup),
    };
    this.setup.margins = {
      top: 1,
      bottom: 1,
      left: 1,
      right: 1,
      ...effectivePageSetup?.margins,
    };

    this.header = '';
    this.footer = '';
    this.numbering = false;
    this.styles = {};
    const rawComments = data?.comments || activity?.comments || [];
    this.comments = clone(rawComments).map((x: any) => ({
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
    const p2 = this.page2Element;
    const text = (root.textContent || '') + (p2 ? ' ' + (p2.textContent || '') : '');

    switch (g['type']) {
      case 'textStyle':
        return styleMatches(root, g) || (p2 ? styleMatches(p2, g) : false);

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
        const r = findText(root, g['text']) || (p2 ? findText(p2, g['text']) : null);
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

        let totalMargin = parseFloat(s.marginLeft) || 0;
        let parentEl = b.parentElement;
        while (parentEl && parentEl !== root && parentEl !== p2) {
          if (parentEl.tagName === 'BLOCKQUOTE') {
            const bqStyle = window.getComputedStyle(parentEl);
            totalMargin += parseFloat(bqStyle.marginLeft) || 40;
          }
          parentEl = parentEl.parentElement;
        }

        return (
          (g['style'] === undefined || style === g['style']) &&
          (g['align'] === undefined ||
            (s.textAlign === 'start' ? 'left' : s.textAlign) === g['align']) &&
          (g['indent'] === undefined ||
            Math.round(totalMargin / 40) === g['indent']) &&
          lineSpacingMatch
        );
      }

      case 'list': {
        const checkList = (container: HTMLElement) =>
          g['items'].every((t: string) => {
            const r = findText(container, t);
            const li = r?.startContainer.parentElement?.closest('li');
            if (!li || li.parentElement?.tagName !== (g['listType'] === 'bullet' ? 'UL' : 'OL')) {
              return false;
            }
            let level = 0;
            let p: HTMLElement | null = li.parentElement?.parentElement as HTMLElement | null;
            while (p && p !== container) {
              if (p.tagName === 'UL' || p.tagName === 'OL') level++;
              p = p.parentElement;
            }
            return g['level'] === undefined || g['level'] === level;
          });
        return checkList(root) || (p2 ? checkList(p2) : false);
      }

      case 'pageSetup':
        return (
          matches(this.setup, g, ['orientation', 'paperSize']) &&
          (!g['margins'] ||
            Object.entries(g['margins']).every(
              ([k, v]) => (this.setup.margins as Record<string, number>)[k] === v
            ))
        );

      case 'table': {
        const tables = [...Array.from(root.querySelectorAll('table')), ...(p2 ? Array.from(p2.querySelectorAll('table')) : [])];
        return tables.some((t) => {
          const rowsMatch = g['rows'] === undefined || t.rows.length === g['rows'];
          const colsMatch =
            g['columns'] === undefined ||
            Array.from(t.rows).every((r) => r.cells.length === g['columns']);
          const cellsMatch =
            !g['cells'] ||
            g['cells'].every((r: string[], i: number) =>
              r.every((c: string, j: number) => t.rows[i]?.cells[j]?.textContent === c)
            );
          const styleMatch =
            g['style'] === undefined ||
            t.getAttribute('data-table-style') === g['style'];
          const colorsMatch =
            !g['cellColors'] ||
            g['cellColors'].every((cc: { row: number; col: number; color: string }) => {
              const cell = t.rows[cc.row]?.cells[cc.col];
              if (!cell) return false;
              const expectedRaw = COLORS[cc.color] || cc.color;
              const probe = document.createElement('span');
              probe.style.backgroundColor = expectedRaw;
              root.append(probe);
              const expectedComputed = window.getComputedStyle(probe).backgroundColor;
              probe.remove();
              const actualComputed = window.getComputedStyle(cell).backgroundColor;
              return actualComputed === expectedComputed;
            });
          return rowsMatch && colsMatch && cellsMatch && styleMatch && colorsMatch;
        });
      }

      case 'image': {
        const images = [...Array.from(root.querySelectorAll('img')), ...(p2 ? Array.from(p2.querySelectorAll('img')) : [])];
        return images.some((im) => {
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
      }

      case 'element':
        if (g['element'] === 'pageBreak') {
          return this.hasPageBreak || !!root.querySelector('[data-element="pageBreak"]') || !!p2?.querySelector('[data-element="pageBreak"]');
        }
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
          return !!root.querySelector('.table-of-contents, [data-element="tableOfContents"], nav.toc') ||
                 !!p2?.querySelector('.table-of-contents, [data-element="tableOfContents"], nav.toc');
        }
        if (g['element'] === 'footnote') {
          return !!root.querySelector('.footnote, [data-element="footnote"], sup.footnote-ref') ||
                 !!p2?.querySelector('.footnote, [data-element="footnote"], sup.footnote-ref');
        }
        const elements = [...Array.from(root.querySelectorAll('[data-element]')), ...(p2 ? Array.from(p2.querySelectorAll('[data-element]')) : [])];
        return elements.some((n: any) => {
          return (
            n.dataset['element'] === g['element'] &&
            (!g['text'] || n.textContent?.includes(g['text'])) &&
            (g['resolved'] === undefined || (n.dataset['resolved'] === 'true') === g['resolved'])
          );
        });

      case 'action':
        return this.actions.some((a) => matches(a, g, ['action', 'text']));

      case 'style': {
        const name = g['name'];
        if (!name) return false;
        const s = this.customStyles.find((cs) => cs.name === name);
        if (!s) return false;

        if (g['bold'] !== undefined && !!s.bold !== g['bold']) return false;
        if (g['italic'] !== undefined && !!s.italic !== g['italic']) return false;
        if (g['underline'] !== undefined && !!s.underline !== g['underline']) return false;
        if (g['font'] !== undefined && s.font !== g['font']) return false;
        if (g['size'] !== undefined && (s.size === undefined || Math.abs(s.size - g['size']) > 0.15)) return false;
        if (g['color'] !== undefined && s.color !== g['color']) return false;

        if (g['appliedTo']) {
          const r = findText(root, g['appliedTo']) || (p2 ? findText(p2, g['appliedTo']) : null);
          if (!r) return false;
          const b = r.startContainer.parentElement?.closest('p,h1,h2,h3,li,div') as HTMLElement | null;
          if (!b) return false;
          if (b.dataset['paragraphStyle'] !== name) return false;
        }

        return true;
      }

      case 'chart': {
        const charts = [
          ...Array.from(root.querySelectorAll<HTMLElement>('figure.doc-chart, [data-element="chart"]')),
          ...(p2 ? Array.from(p2.querySelectorAll<HTMLElement>('figure.doc-chart, [data-element="chart"]')) : []),
        ];

        if (g['count'] !== undefined) {
          if (charts.length !== Number(g['count'])) return false;
        }

        if (g['chartType'] !== undefined || g['title'] !== undefined) {
          return charts.some((ch) => {
            const typeMatch =
              !g['chartType'] ||
              (ch.getAttribute('data-chart-type') || '').toLowerCase() === String(g['chartType']).toLowerCase();
            const titleMatch =
              g['title'] === undefined ||
              (ch.getAttribute('data-chart-title') || '').trim() === String(g['title']).trim();
            return typeMatch && titleMatch;
          });
        }

        return g['count'] !== undefined ? true : charts.length > 0;
      }

      default:
        return false;
    }
  }

  demo(g: Record<string, any>, activity?: ActivityPayload): void {
    if (!this.editorElement) return;
    const root = this.editorElement;
    const p2 = this.page2Element;

    switch (g['type']) {
      case 'textStyle': {
        const r = findText(root, g['text']);
        if (r && this.rich) {
          setSelection(r);
          this.rich.capture();
          if (g['bold']) this.rich.command('bold');
          if (g['italic']) this.rich.command('italic');
          if (g['underline']) this.rich.command('underline');
          if (g['superscript']) this.rich.command('superscript');
          if (g['subscript']) this.rich.command('subscript');
          if (g['font']) this.rich.apply({ fontFamily: g['font'] });
          if (g['size']) this.rich.apply({ fontSize: g['size'] + 'pt' });
          if (g['color']) this.rich.apply({ color: COLORS[g['color']] || g['color'] });
          if (g['highlight'])
            this.rich.apply({ backgroundColor: COLORS[g['highlight']] || g['highlight'] });
        }
        break;
      }

      case 'style': {
        const name = g['name'];
        if (name) {
          const customDef: CustomStyleDefinition = {
            name,
            bold: g['bold'],
            italic: g['italic'],
            underline: g['underline'],
            font: g['font'],
            size: g['size'],
            color: g['color'],
          };
          const idx = this.customStyles.findIndex((cs) => cs.name === name);
          if (idx >= 0) {
            this.customStyles[idx] = customDef;
          } else {
            this.customStyles.push(customDef);
          }
          if (g['appliedTo']) {
            const r = findText(root, g['appliedTo']) || (p2 ? findText(p2, g['appliedTo']) : null);
            if (r) {
              const b = r.startContainer.parentElement?.closest('p,h1,h2,h3,li,div') as HTMLElement | null;
              if (b) {
                b.dataset['paragraphStyle'] = name;
                if (customDef.bold !== undefined) b.style.fontWeight = customDef.bold ? 'bold' : 'normal';
                if (customDef.italic !== undefined) b.style.fontStyle = customDef.italic ? 'italic' : 'normal';
                if (customDef.underline !== undefined) b.style.textDecoration = customDef.underline ? 'underline' : 'none';
                if (customDef.font) b.style.fontFamily = customDef.font;
                if (customDef.size) b.style.fontSize = customDef.size + 'pt';
                if (customDef.color) b.style.color = COLORS[customDef.color] || customDef.color;
              }
            }
          }
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
        let tbl = root.querySelector('table');
        if (!tbl) {
          tbl = document.createElement('table');
          tbl.className = 'doc-table';
          root.append(tbl);
        }
        if (g['style']) {
          tbl.setAttribute('data-table-style', g['style']);
        }
        const rows = g['rows'] || tbl.rows.length || 2;
        const cols = g['columns'] || tbl.rows[0]?.cells.length || 2;
        while (tbl.rows.length < rows) {
          const tr = tbl.insertRow();
          for (let j = 0; j < cols; j++) {
            const td = tr.insertCell();
            td.textContent = 'Cell';
          }
        }
        for (let i = 0; i < tbl.rows.length; i++) {
          while (tbl.rows[i].cells.length < cols) {
            const td = tbl.rows[i].insertCell();
            td.textContent = 'Cell';
          }
        }
        if (g['cells']) {
          for (let i = 0; i < g['cells'].length; i++) {
            for (let j = 0; j < g['cells'][i].length; j++) {
              if (tbl.rows[i]?.cells[j]) {
                tbl.rows[i].cells[j].textContent = g['cells'][i][j];
              }
            }
          }
        }
        if (g['cellColors']) {
          for (const cc of g['cellColors']) {
            const cell = tbl.rows[cc.row]?.cells[cc.col];
            if (cell) {
              cell.style.backgroundColor = COLORS[cc.color] || cc.color;
            }
          }
        }
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

      case 'chart': {
        const type = ((g['chartType'] as string) || 'column').toLowerCase() as WordChartType;
        const title = (g['title'] as string) || (type.charAt(0).toUpperCase() + type.slice(1) + ' Chart');
        const data = DEFAULT_CHART_DATA;
        const fig = document.createElement('figure');
        fig.className = 'doc-chart';
        fig.setAttribute('data-element', 'chart');
        fig.setAttribute('data-chart-type', type);
        fig.setAttribute('data-chart-title', title);
        fig.setAttribute('data-chart-data', JSON.stringify(data));
        fig.setAttribute('contenteditable', 'false');
        fig.innerHTML = renderChartSvg(type, title, data);
        root.append(fig);
        break;
      }
    }
  }

  snapshot(): Record<string, any> {
    return {
      html: this.editorElement ? this.editorElement.innerHTML : this.html,
      setup: clone(this.setup),
      header: this.header,
      footer: this.footer,
      numbering: this.numbering,
      hasPageBreak: this.hasPageBreak,
      page2Html: this.page2Element ? this.page2Element.innerHTML : '',
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
    this.hasPageBreak = !!s['hasPageBreak'];
    if (this.page2Element) {
      this.page2Element.innerHTML = s['page2Html'] || '';
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
