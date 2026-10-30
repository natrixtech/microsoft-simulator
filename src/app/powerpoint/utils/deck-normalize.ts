import { Deck, ElementKind, Slide, SlideElement, SlideLayout, ThemeId } from '../models/ppt.types';
import { DEFAULT_THEME_ID, PPT_THEMES } from '../constants/ppt-themes';
import { PPT_LAYOUTS } from '../constants/ppt-layouts';
import { plainTextToHtml, sanitizeRichText } from './rich-text-sanitize';

let slideCounter = 1;
let elementCounter = 1;

export function generateSlideId(): string {
  return `slide-${Date.now()}-${slideCounter++}`;
}

export function generateElementId(): string {
  return `el-${Date.now()}-${elementCounter++}`;
}

export function normalizeThemeId(rawTheme: unknown): ThemeId {
  if (typeof rawTheme === 'string' && rawTheme in PPT_THEMES) {
    return rawTheme as ThemeId;
  }
  return DEFAULT_THEME_ID;
}

export function normalizeSlideLayout(rawLayout: unknown): SlideLayout {
  if (typeof rawLayout === 'string' && rawLayout in PPT_LAYOUTS) {
    return rawLayout as SlideLayout;
  }
  return 'titleAndBody';
}

export function normalizeElement(rawEl: any): SlideElement {
  const id = rawEl.id || generateElementId();
  const kind = rawEl.kind as ElementKind;
  const x = Number(rawEl.x) || 0;
  const y = Number(rawEl.y) || 0;
  const w = Number(rawEl.w) || 100;
  const h = Number(rawEl.h) || 100;

  if (kind === 'image') {
    return {
      id,
      kind: 'image',
      file: String(rawEl.file || ''),
      alt: rawEl.alt ? String(rawEl.alt) : undefined,
      x,
      y,
      w,
      h,
    };
  }

  if (kind === 'shape') {
    return {
      id,
      kind: 'shape',
      shape: rawEl.shape || 'rectangle',
      fill: rawEl.fill,
      stroke: rawEl.stroke,
      x,
      y,
      w,
      h,
    };
  }

  if (kind === 'table') {
    const rows = Number(rawEl.rows) || 2;
    const columns = Number(rawEl.columns) || 2;
    let cells: string[][] = [];
    if (Array.isArray(rawEl.cells)) {
      cells = rawEl.cells.map((row: any) =>
        Array.isArray(row) ? row.map((c) => String(c ?? '')) : []
      );
    } else {
      cells = Array.from({ length: rows }, () => Array.from({ length: columns }, () => ''));
    }
    return {
      id,
      kind: 'table',
      rows,
      columns,
      cells,
      x,
      y,
      w,
      h,
    };
  }

  if (kind === 'textBox') {
    const text = String(rawEl.text ?? '');
    const html = rawEl.html
      ? sanitizeRichText(String(rawEl.html))
      : (text ? plainTextToHtml(text, false) : undefined);

    return {
      id,
      kind: 'textBox',
      text,
      html,
      bold: !!rawEl.bold,
      italic: !!rawEl.italic,
      underline: !!rawEl.underline,
      fontSize: rawEl.fontSize ? Number(rawEl.fontSize) : undefined,
      fontFamily: rawEl.fontFamily,
      color: rawEl.color,
      align: rawEl.align,
      x,
      y,
      w,
      h,
    };
  }

  throw new Error(`Unknown element kind: ${rawEl.kind}`);
}

export function normalizeSlide(rawSlide: any, index: number): Slide {
  const id = rawSlide.id || `slide-${index + 1}-${generateSlideId()}`;
  const layout = normalizeSlideLayout(rawSlide.layout || (index === 0 ? 'title' : 'titleAndBody'));
  const elements = Array.isArray(rawSlide.elements)
    ? rawSlide.elements.map(normalizeElement)
    : [];

  const title = rawSlide.title !== undefined ? String(rawSlide.title) : '';
  const subtitle = rawSlide.subtitle !== undefined ? String(rawSlide.subtitle) : '';
  const body = rawSlide.body !== undefined ? String(rawSlide.body) : '';
  const body2 = rawSlide.body2 !== undefined ? String(rawSlide.body2) : '';

  return {
    id,
    layout,
    title,
    titleHtml: rawSlide.titleHtml
      ? sanitizeRichText(String(rawSlide.titleHtml))
      : (title ? plainTextToHtml(title, false) : undefined),
    subtitle,
    subtitleHtml: rawSlide.subtitleHtml
      ? sanitizeRichText(String(rawSlide.subtitleHtml))
      : (subtitle ? plainTextToHtml(subtitle, false) : undefined),
    body,
    bodyHtml: rawSlide.bodyHtml
      ? sanitizeRichText(String(rawSlide.bodyHtml))
      : (body ? plainTextToHtml(body, false) : undefined),
    body2,
    body2Html: rawSlide.body2Html
      ? sanitizeRichText(String(rawSlide.body2Html))
      : (body2 ? plainTextToHtml(body2, false) : undefined),
    background: rawSlide.background ? String(rawSlide.background) : undefined,
    elements,
    transition: rawSlide.transition ? { ...rawSlide.transition } : { type: 'none' },
  };
}

export function normalizeDeck(rawDeck: any): Deck {
  const theme = normalizeThemeId(rawDeck?.theme);
  let slides: Slide[] = [];

  if (rawDeck && Array.isArray(rawDeck.slides) && rawDeck.slides.length > 0) {
    slides = rawDeck.slides.map(normalizeSlide);
  } else {
    slides = [createDefaultSlide('title')];
  }

  return {
    theme,
    slides,
  };
}

export function createDefaultSlide(layout: SlideLayout = 'titleAndBody'): Slide {
  return {
    id: generateSlideId(),
    layout,
    title: '',
    subtitle: '',
    body: '',
    body2: '',
    elements: [],
    transition: { type: 'none' },
  };
}
