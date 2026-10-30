import { Injectable } from '@angular/core';
import {
  Deck,
  PptGoal,
  PptGoalContext,
  PptSlideTextGoal,
  PptTextStyleGoal,
  PptListGoal,
  PptSlideCountGoal,
  PptImageGoal,
  PptTextBoxGoal,
  PptShapeGoal,
  PptTableGoal,
  PptBackgroundGoal,
  PptLayoutGoal,
  PptThemeGoal,
  PptTransitionGoal,
  PptSlideOrderGoal,
  PptActionGoal,
  Slide,
  TextBoxSlideElement,
  ShapeSlideElement,
  TableSlideElement,
  ImageSlideElement,
} from '../models/ppt.types';
import { COLORS } from '../constants/ppt-colors';
import { PPT_THEMES } from '../constants/ppt-themes';
import { PPT_LAYOUTS } from '../constants/ppt-layouts';

@Injectable({
  providedIn: 'root',
})
export class PptGoalEvaluatorService {
  evaluate(goals: PptGoal[] | undefined, deck: Deck, ctx: PptGoalContext): boolean {
    if (!goals || goals.length === 0) return true;
    return goals.every((goal) => this.evaluateGoal(goal, deck, ctx));
  }

  private evaluateGoal(goal: PptGoal, deck: Deck, ctx: PptGoalContext): boolean {
    switch (goal.type) {
      case 'slideCount':
        return this.evaluateSlideCount(goal, deck);
      case 'theme':
        return this.evaluateTheme(goal, deck, ctx);
      case 'slideOrder':
        return this.evaluateSlideOrder(goal, deck);
      case 'action':
        return this.evaluateAction(goal, ctx);
      default: {
        const slide = this.resolveSlide(goal.slide, deck, ctx);
        if (!slide) return false;

        switch (goal.type) {
          case 'slideText':
            return this.evaluateSlideText(goal, slide);
          case 'textStyle':
            return this.evaluateTextStyle(goal, slide, deck);
          case 'list':
            return this.evaluateList(goal, slide);
          case 'image':
            return this.evaluateImage(goal, slide);
          case 'textBox':
            return this.evaluateTextBox(goal, slide);
          case 'shape':
            return this.evaluateShape(goal, slide);
          case 'table':
            return this.evaluateTable(goal, slide);
          case 'background':
            return this.evaluateBackground(goal, slide, deck);
          case 'layout':
            return this.evaluateLayout(goal, slide);
          case 'transition':
            return this.evaluateTransition(goal, slide);
          default:
            return true;
        }
      }
    }
  }

  private resolveSlide(slide1Based: number | undefined, deck: Deck, ctx: PptGoalContext): Slide | undefined {
    if (slide1Based !== undefined) {
      return deck.slides[slide1Based - 1];
    }
    return deck.slides[ctx.currentSlideIndex];
  }

  private normalizeText(text: string | undefined | null, exact?: boolean): string {
    if (!text) return '';
    if (exact) return text;
    return text
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  private extractPlainTextLines(htmlOrText: string | undefined | null): string[] {
    if (!htmlOrText) return [];
    try {
      const doc = new DOMParser().parseFromString(htmlOrText, 'text/html');
      let text = '';
      const walk = (node: Node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          text += node.textContent || '';
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          const el = node as Element;
          const tag = el.tagName.toLowerCase();
          if (tag === 'br') {
            text += '\n';
            return;
          }
          const isBlock = ['div', 'p', 'li', 'tr', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote'].includes(tag);
          if (isBlock) text += '\n';
          for (let i = 0; i < el.childNodes.length; i++) {
            walk(el.childNodes[i]);
          }
          if (isBlock) text += '\n';
        }
      };
      walk(doc.body);
      return text
        .split('\n')
        .map((l) => l.trim())
        .filter((l) => l.length > 0);
    } catch {
      return (htmlOrText || '')
        .split('\n')
        .map((l) => l.trim())
        .filter((l) => l.length > 0);
    }
  }

  private extractTextFromHtml(htmlOrText: string | undefined | null, exact?: boolean): string {
    const lines = this.extractPlainTextLines(htmlOrText);
    const joined = lines.join('\n');
    return this.normalizeText(joined, exact);
  }

  private getLineCount(htmlOrText: string | undefined | null): number {
    return this.extractPlainTextLines(htmlOrText).length;
  }

  private colorToHex(colorStr: string | undefined | null): string {
    if (!colorStr) return '';
    const trimmed = colorStr.trim().toLowerCase();
    if (COLORS[trimmed]) {
      return COLORS[trimmed].toLowerCase();
    }
    for (const [name, hex] of Object.entries(COLORS)) {
      if (name.toLowerCase() === trimmed) {
        return hex.toLowerCase();
      }
    }
    const rgbMatch = trimmed.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (rgbMatch) {
      const r = parseInt(rgbMatch[1], 10).toString(16).padStart(2, '0');
      const g = parseInt(rgbMatch[2], 10).toString(16).padStart(2, '0');
      const b = parseInt(rgbMatch[3], 10).toString(16).padStart(2, '0');
      return `#${r}${g}${b}`.toLowerCase();
    }
    if (trimmed.startsWith('#')) return trimmed;
    if (/^[0-9a-f]{6}$/i.test(trimmed)) return `#${trimmed}`;
    return trimmed;
  }

  private evaluateSlideText(goal: PptSlideTextGoal, slide: Slide): boolean {
    const candidates: string[] = [];
    const ph = goal.placeholder || 'any';

    if (ph === 'title') {
      candidates.push(slide.titleHtml || slide.title || '');
    } else if (ph === 'subtitle') {
      candidates.push(slide.subtitleHtml || slide.subtitle || '');
    } else if (ph === 'body') {
      candidates.push(slide.bodyHtml || slide.body || '');
    } else if (ph === 'body2') {
      candidates.push(slide.body2Html || slide.body2 || '');
    } else {
      candidates.push(slide.titleHtml || slide.title || '');
      candidates.push(slide.subtitleHtml || slide.subtitle || '');
      candidates.push(slide.bodyHtml || slide.body || '');
      candidates.push(slide.body2Html || slide.body2 || '');
      for (const el of slide.elements) {
        if (el.kind === 'textBox') {
          const tb = el as TextBoxSlideElement;
          candidates.push(tb.html || tb.text || '');
        }
      }
    }

    const checkCandidate = (raw: string): boolean => {
      const text = this.extractTextFromHtml(raw, goal.exact);

      if (goal.contains !== undefined) {
        const list = Array.isArray(goal.contains) ? goal.contains : [goal.contains];
        const allContained = list.every((item) =>
          text.includes(this.normalizeText(item, goal.exact))
        );
        if (!allContained) return false;
      }

      if (goal.equals !== undefined) {
        if (text !== this.normalizeText(goal.equals, goal.exact)) {
          return false;
        }
      }

      if (goal.notContains !== undefined && goal.notContains.length > 0) {
        const anyDisallowed = goal.notContains.some((item) =>
          text.includes(this.normalizeText(item, goal.exact))
        );
        if (anyDisallowed) return false;
      }

      if (goal.minLines !== undefined) {
        if (this.getLineCount(raw) < goal.minLines) {
          return false;
        }
      }

      return true;
    };

    if (ph === 'any') {
      return candidates.some((c) => c.trim().length > 0 && checkCandidate(c));
    } else {
      return checkCandidate(candidates[0] || '');
    }
  }

  private evaluateTextStyle(goal: PptTextStyleGoal, slide: Slide, deck: Deck): boolean {
    const targetText = goal.text || '';
    if (!targetText) return true;

    // Find candidate text and field/element
    interface TextCandidate {
      html: string;
      field?: 'title' | 'subtitle' | 'body' | 'body2';
      element?: TextBoxSlideElement;
    }

    const candidates: TextCandidate[] = [];
    const ph = goal.placeholder || 'any';

    if (ph === 'title') {
      candidates.push({ html: slide.titleHtml || slide.title || '', field: 'title' });
    } else if (ph === 'subtitle') {
      candidates.push({ html: slide.subtitleHtml || slide.subtitle || '', field: 'subtitle' });
    } else if (ph === 'body') {
      candidates.push({ html: slide.bodyHtml || slide.body || '', field: 'body' });
    } else if (ph === 'body2') {
      candidates.push({ html: slide.body2Html || slide.body2 || '', field: 'body2' });
    } else {
      candidates.push({ html: slide.titleHtml || slide.title || '', field: 'title' });
      candidates.push({ html: slide.subtitleHtml || slide.subtitle || '', field: 'subtitle' });
      candidates.push({ html: slide.bodyHtml || slide.body || '', field: 'body' });
      candidates.push({ html: slide.body2Html || slide.body2 || '', field: 'body2' });
      for (const el of slide.elements) {
        if (el.kind === 'textBox') {
          const tb = el as TextBoxSlideElement;
          candidates.push({ html: tb.html || tb.text || '', element: tb });
        }
      }
    }

    // Find first candidate containing goal.text
    const matchingCandidate = candidates.find((c) =>
      this.extractTextFromHtml(c.html, false).includes(this.normalizeText(targetText, false))
    );

    if (!matchingCandidate) return false;

    // Render into hidden off-screen div with base styles
    const offscreen = document.createElement('div');
    offscreen.style.position = 'fixed';
    offscreen.style.left = '-9999px';
    offscreen.style.top = '-9999px';
    offscreen.style.visibility = 'hidden';

    const theme = PPT_THEMES[deck.theme] || PPT_THEMES['simple-light'];
    const layout = PPT_LAYOUTS[slide.layout] || PPT_LAYOUTS['title'];
    const phDef = matchingCandidate.field
      ? layout.placeholders.find((p) => p.type === matchingCandidate.field)
      : undefined;

    const baseFont = matchingCandidate.field === 'title' ? theme.titleFont : theme.bodyFont;
    const baseColor = matchingCandidate.field === 'title' ? theme.titleColor : theme.bodyColor;
    const baseSize = phDef?.fontSize || matchingCandidate.element?.fontSize || 24;
    const baseAlign = phDef?.align || matchingCandidate.element?.align || 'left';

    offscreen.style.fontFamily = baseFont;
    offscreen.style.color = baseColor;
    offscreen.style.fontSize = `${baseSize}px`;
    offscreen.style.textAlign = baseAlign;

    if (matchingCandidate.element) {
      if (matchingCandidate.element.bold) offscreen.style.fontWeight = 'bold';
      if (matchingCandidate.element.italic) offscreen.style.fontStyle = 'italic';
      if (matchingCandidate.element.underline) offscreen.style.textDecoration = 'underline';
      if (matchingCandidate.element.fontFamily) offscreen.style.fontFamily = matchingCandidate.element.fontFamily;
      if (matchingCandidate.element.color) offscreen.style.color = matchingCandidate.element.color;
    }

    document.body.appendChild(offscreen);
    offscreen.innerHTML = matchingCandidate.html;

    try {
      const walker = document.createTreeWalker(offscreen, NodeFilter.SHOW_TEXT);
      const textNodes: Text[] = [];
      let n: Node | null;
      while ((n = walker.nextNode())) {
        textNodes.push(n as Text);
      }

      let fullText = '';
      const charMap: { node: Text; offset: number }[] = [];
      for (const tn of textNodes) {
        const str = tn.textContent || '';
        for (let i = 0; i < str.length; i++) {
          charMap.push({ node: tn, offset: i });
          fullText += str[i] === '\u00a0' ? ' ' : str[i];
        }
      }

      const matchIndex = fullText.toLowerCase().indexOf(targetText.replace(/\u00a0/g, ' ').toLowerCase());
      if (matchIndex === -1) {
        return false;
      }

      for (let i = matchIndex; i < matchIndex + targetText.length; i++) {
        const { node: tn } = charMap[i];
        const parent = tn.parentElement || offscreen;
        const cs = getComputedStyle(parent);

        if (goal.bold !== undefined) {
          const isBold = cs.fontWeight === 'bold' || parseInt(cs.fontWeight, 10) >= 600;
          if (isBold !== goal.bold) return false;
        }

        if (goal.italic !== undefined) {
          const isItalic = cs.fontStyle === 'italic';
          if (isItalic !== goal.italic) return false;
        }

        if (goal.underline !== undefined) {
          let isUnderline = false;
          let cur: HTMLElement | null = parent;
          while (cur) {
            const ccs = getComputedStyle(cur);
            if ((ccs.textDecorationLine || ccs.textDecoration || '').includes('underline')) { isUnderline = true; break; }
            if (cur === offscreen) break;
            cur = cur.parentElement;
          }
          if (isUnderline !== goal.underline) return false;
        }

        if (goal.font !== undefined) {
          const font = cs.fontFamily.toLowerCase().replace(/['"]/g, '');
          if (!font.includes(goal.font.toLowerCase())) return false;
        }

        if (goal.size !== undefined) {
          const size = parseFloat(cs.fontSize);
          if (Math.abs(size - goal.size) > 0.25) return false;
        }

        if (goal.color !== undefined) {
          const actualHex = this.colorToHex(cs.color);
          const expectedHex = this.colorToHex(goal.color);
          if (actualHex !== expectedHex) return false;
        }

        if (goal.align !== undefined) {
          if (cs.textAlign !== goal.align) return false;
        }
      }

      return true;
    } finally {
      document.body.removeChild(offscreen);
    }
  }

  private evaluateList(goal: PptListGoal, slide: Slide): boolean {
    const ph = goal.placeholder || 'body';
    let html = '';
    if (ph === 'body') html = slide.bodyHtml || slide.body || '';
    else if (ph === 'body2') html = slide.body2Html || slide.body2 || '';
    else if (ph === 'title') html = slide.titleHtml || slide.title || '';
    else if (ph === 'subtitle') html = slide.subtitleHtml || slide.subtitle || '';

    const tag = goal.listType === 'numbered' ? 'ol' : 'ul';
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const lists = Array.from(doc.querySelectorAll(tag));
    if (lists.length === 0) return false;
    if (goal.minItems !== undefined) {
      const items = lists
        .flatMap((l) => Array.from(l.querySelectorAll(':scope > li')))
        .map((li) => this.extractTextFromHtml(li.innerHTML))
        .filter((t) => t.length > 0);
      if (items.length < goal.minItems) return false;
    }

    return true;
  }

  private evaluateSlideCount(goal: PptSlideCountGoal, deck: Deck): boolean {
    if (goal.count !== undefined && deck.slides.length !== goal.count) return false;
    if (goal.min !== undefined && deck.slides.length < goal.min) return false;
    return true;
  }

  private evaluateImage(goal: PptImageGoal, slide: Slide): boolean {
    const images = slide.elements.filter((e) => e.kind === 'image') as ImageSlideElement[];
    if (images.length === 0) return false;

    return images.some((img) => {
      if (goal.file !== undefined && goal.file !== null) {
        const want = goal.file.split('/').pop()!.toLowerCase();
        const have = (img.file || '').split('/').pop()!.toLowerCase();
        if (!have || have !== want) {
          return false;
        }
      }

      if (goal.minWidth !== undefined && img.w < goal.minWidth) return false;
      if (goal.maxWidth !== undefined && img.w > goal.maxWidth) return false;

      if (goal.position !== undefined) {
        const cx = img.x + img.w / 2;
        const cy = img.y + img.h / 2;
        if (goal.position === 'left' && cx >= 480) return false;
        if (goal.position === 'right' && cx < 480) return false;
        if (goal.position === 'top' && cy >= 270) return false;
        if (goal.position === 'bottom' && cy < 270) return false;
        if (goal.position === 'center' && (cx < 240 || cx > 720 || cy < 135 || cy > 405)) return false;
      }

      return true;
    });
  }

  private evaluateTextBox(goal: PptTextBoxGoal, slide: Slide): boolean {
    const textboxes = slide.elements.filter((e) => e.kind === 'textBox') as TextBoxSlideElement[];
    if (textboxes.length === 0) return false;

    if (goal.contains !== undefined) {
      const expected = this.normalizeText(goal.contains);
      return textboxes.some((tb) =>
        this.extractTextFromHtml(tb.html || tb.text || '').includes(expected)
      );
    }

    return true;
  }

  private evaluateShape(goal: PptShapeGoal, slide: Slide): boolean {
    const shapes = slide.elements.filter((e) => e.kind === 'shape') as ShapeSlideElement[];
    const matching = goal.shape ? shapes.filter((s) => s.shape === goal.shape) : shapes;
    if (goal.count !== undefined) {
      return matching.length === goal.count;
    }
    return matching.length > 0;
  }

  private evaluateTable(goal: PptTableGoal, slide: Slide): boolean {
    const tables = slide.elements.filter((e) => e.kind === 'table') as TableSlideElement[];
    if (tables.length === 0) return false;

    return tables.some((tbl) => {
      if (goal.rows !== undefined && tbl.rows !== goal.rows) return false;
      const cols = (tbl as any).cols !== undefined ? (tbl as any).cols : tbl.columns;
      if (goal.columns !== undefined && cols !== goal.columns) return false;

      if (goal.cells !== undefined) {
        for (let r = 0; r < goal.cells.length; r++) {
          for (let c = 0; c < goal.cells[r].length; c++) {
            const expected = this.normalizeText(goal.cells[r][c]);
            if (expected === '') continue;
            const actual = this.extractTextFromHtml(tbl.cells?.[r]?.[c] || '');
            if (actual !== expected) return false;
          }
        }
      }
      return true;
    });
  }

  private evaluateBackground(goal: PptBackgroundGoal, slide: Slide, deck: Deck): boolean {
    const bg = (slide.background || '').trim();
    const theme = PPT_THEMES[deck.theme] || PPT_THEMES['simple-light'];

    if (goal.color !== undefined) {
      if (goal.color === 'any') {
        if (!bg || bg === 'transparent' || this.colorToHex(bg) === this.colorToHex(theme.background)) {
          return false;
        }
      } else {
        const expectedHex = this.colorToHex(goal.color);
        const actualHex = this.colorToHex(bg);
        if (actualHex !== expectedHex) return false;
      }
    }

    if (goal.file !== undefined) {
      const want = goal.file.split('/').pop()!.toLowerCase();
      const have = bg.split('/').pop()!.toLowerCase();
      if (!have || have !== want) {
        return false;
      }
    }

    return true;
  }

  private evaluateLayout(goal: PptLayoutGoal, slide: Slide): boolean {
    return slide.layout === goal.layout;
  }

  private evaluateTheme(goal: PptThemeGoal, deck: Deck, ctx: PptGoalContext): boolean {
    if (goal.theme === 'any') {
      return deck.theme !== ctx.startingTheme;
    }
    return deck.theme === goal.theme;
  }

  private evaluateTransition(goal: PptTransitionGoal, slide: Slide): boolean {
    const tr = slide.transition;
    const want = goal.transition === undefined || goal.transition === null || String(goal.transition).trim() === ''
      ? 'any'
      : goal.transition;
    const actualType = tr?.type || 'none';
    if (want === 'any') {
      return actualType !== 'none';
    }
    const norm = (want || '').toLowerCase().replace(/[\s\-_]/g, '');
    if (norm === 'none') return actualType === 'none';
    if (norm === 'slidefromright') {
      return tr?.type === 'push' && tr?.direction === 'fromRight';
    }
    if (norm === 'slidefromleft') {
      return tr?.type === 'push' && tr?.direction === 'fromLeft';
    }
    if (norm === 'push') {
      return tr?.type === 'push';
    }
    return actualType === want || actualType.toLowerCase() === norm;
  }

  private evaluateSlideOrder(goal: PptSlideOrderGoal, deck: Deck): boolean {
    let titleIdx = 0;
    for (const s of deck.slides) {
      const sTitle = this.extractTextFromHtml(s.titleHtml || s.title || '');
      const targetTitle = this.normalizeText(goal.titles[titleIdx]);
      if (targetTitle !== '' && sTitle === targetTitle) {
        titleIdx++;
        if (titleIdx >= goal.titles.length) return true;
      }
    }
    return titleIdx >= goal.titles.length;
  }

  private evaluateAction(goal: PptActionGoal, ctx: PptGoalContext): boolean {
    return ctx.actionLog.some((a) => {
      if (a.action !== goal.action) return false;
      if (goal.slide !== undefined && a.slide !== goal.slide) return false;
      if (goal.menu !== undefined && a.menu?.toLowerCase() !== goal.menu.toLowerCase()) return false;
      return true;
    });
  }
}

