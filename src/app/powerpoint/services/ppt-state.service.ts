import { Injectable, computed, signal } from '@angular/core';
import {
  Deck,
  PlaceholderType,
  PptSnapshot,
  Slide,
  SlideElement,
  SlideLayout,
  SlideTransition,
  TableSlideElement,
  ThemeDef,
  ThemeId,
} from '../models/ppt.types';
import { DEFAULT_THEME_ID, PPT_THEMES } from '../constants/ppt-themes';
import { clone } from '../../shared/utils/editor.utils';
import { createDefaultSlide, generateSlideId } from '../utils/deck-normalize';
import { PptActionLoggerService } from './ppt-action-logger.service';
import { PptHistoryService } from './ppt-history.service';

@Injectable({
  providedIn: 'root',
})
export class PptStateService {
  readonly deck = signal<Deck>({
    theme: DEFAULT_THEME_ID,
    slides: [createDefaultSlide('title')],
  });

  readonly currentSlideIndex = signal<number>(0);
  readonly selectedElementId = signal<string | null>(null);
  readonly selectedElementIds = signal<string[]>([]);
  readonly editingPlaceholder = signal<{
    slideId: string;
    field: PlaceholderType;
  } | null>(null);
  readonly isTextActive = signal<boolean>(false);
  readonly clipboardSlide = signal<Slide | null>(null);

  readonly currentSlide = computed<Slide | null>(() => {
    const d = this.deck();
    const idx = this.currentSlideIndex();
    if (idx >= 0 && idx < d.slides.length) {
      return d.slides[idx];
    }
    return d.slides[0] || null;
  });

  readonly selectedElement = computed<SlideElement | null>(() => {
    const slide = this.currentSlide();
    const elId = this.selectedElementId();
    if (!slide || !elId) return null;
    return slide.elements.find((el) => el.id === elId) || null;
  });

  readonly selectedElements = computed<SlideElement[]>(() => {
    const slide = this.currentSlide();
    const ids = this.selectedElementIds();
    if (!slide || ids.length === 0) return [];
    return slide.elements.filter((el) => ids.includes(el.id));
  });

  readonly totalSlides = computed<number>(() => this.deck().slides.length);

  readonly currentTheme = computed<ThemeDef>(() => {
    const themeId = this.deck().theme;
    return PPT_THEMES[themeId] || PPT_THEMES[DEFAULT_THEME_ID];
  });

  constructor(private actionLogger: PptActionLoggerService, private history: PptHistoryService) {}

  initDeck(newDeck: Deck): void {
    this.deck.set(clone(newDeck));
    this.currentSlideIndex.set(0);
    this.selectedElementId.set(null);
    this.selectedElementIds.set([]);
    this.editingPlaceholder.set(null);
    this.isTextActive.set(false);
  }

  setCurrentSlideIndex(idx: number): void {
    const total = this.totalSlides();
    const clamped = Math.max(0, Math.min(total - 1, idx));
    this.currentSlideIndex.set(clamped);
    this.selectedElementId.set(null);
    this.selectedElementIds.set([]);
    this.editingPlaceholder.set(null);
    this.isTextActive.set(false);
    this.actionLogger.log('selectSlide', { slide: clamped + 1 });
  }

  selectSlide(idx: number): void {
    this.setCurrentSlideIndex(idx);
  }

  addSlide(layout: SlideLayout = 'titleAndBody', afterIndex?: number): Slide {
    const newSlide = createDefaultSlide(layout);
    const curIdx = afterIndex ?? this.currentSlideIndex();
    const d = clone(this.deck());
    const insertAt = Math.min(d.slides.length, curIdx + 1);
    d.slides.splice(insertAt, 0, newSlide);

    this.deck.set(d);
    this.setCurrentSlideIndex(insertAt);
    return newSlide;
  }

  
  moveSlide(fromIndex: number, toIndex: number): boolean {
    const d = clone(this.deck());
    const total = d.slides.length;
    if (
      fromIndex < 0 ||
      fromIndex >= total ||
      toIndex < 0 ||
      toIndex >= total ||
      fromIndex === toIndex
    ) {
      return false;
    }

    // Push undo history once before moving
    this.history.push(this.getSnapshot());

    const [moved] = d.slides.splice(fromIndex, 1);
    d.slides.splice(toIndex, 0, moved);

    this.deck.set(d);
    this.setCurrentSlideIndex(toIndex);
    this.actionLogger.log('moveSlide', { from: fromIndex + 1, to: toIndex + 1 });
    return true;
  }

  deleteSlide(slideIndex?: number): boolean {
    const d = clone(this.deck());
    if (d.slides.length <= 1) return false;

    const idx = slideIndex !== undefined ? slideIndex : this.currentSlideIndex();
    if (idx < 0 || idx >= d.slides.length) return false;

    d.slides.splice(idx, 1);
    const newIdx = Math.min(d.slides.length - 1, idx);
    this.deck.set(d);
    this.setCurrentSlideIndex(newIdx);
    return true;
  }

  setLayout(layout: SlideLayout): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].layout = layout;
      this.deck.set(d);
    }
  }

  setTheme(themeId: ThemeId): void {
    const d = clone(this.deck());
    d.theme = themeId;
    this.deck.set(d);
  }

  setSlideTransition(transition: SlideTransition): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].transition = clone(transition);
      this.deck.set(d);
    }
  }

  updatePlaceholderText(slideIndex: number, field: PlaceholderType, text: string, html?: string): void {
    const d = clone(this.deck());
    if (slideIndex >= 0 && slideIndex < d.slides.length) {
      const slide = d.slides[slideIndex];
      if (field === 'title') {
        slide.title = text;
        if (html !== undefined) slide.titleHtml = html;
      } else if (field === 'subtitle') {
        slide.subtitle = text;
        if (html !== undefined) slide.subtitleHtml = html;
      } else if (field === 'body') {
        slide.body = text;
        if (html !== undefined) slide.bodyHtml = html;
      } else if (field === 'body2') {
        slide.body2 = text;
        if (html !== undefined) slide.body2Html = html;
      }
      this.deck.set(d);
    }
  }

  setSlideBackground(background: string | undefined): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].background = background;
      this.deck.set(d);
    }
  }

  setEditingPlaceholder(ph: { slideId: string; field: PlaceholderType } | null): void {
    this.editingPlaceholder.set(ph);
    if (ph !== null) {
      this.selectedElementId.set(null);
      this.selectedElementIds.set([]);
      this.isTextActive.set(true);
    } else {
      const slide = this.currentSlide();
      const selId = this.selectedElementId();
      const el = selId ? slide?.elements.find((e) => e.id === selId) : null;
      this.isTextActive.set(el?.kind === 'textBox');
    }
  }

  setSelectedElement(id: string | null): void {
    this.selectedElementId.set(id);
    this.selectedElementIds.set(id !== null ? [id] : []);
    if (id !== null) {
      this.editingPlaceholder.set(null);
      const slide = this.currentSlide();
      const el = slide?.elements.find((e) => e.id === id);
      this.isTextActive.set(el?.kind === 'textBox');
      this.actionLogger.log('selectElement', { slide: this.currentSlideIndex() + 1 });
    } else {
      if (!this.editingPlaceholder()) {
        this.isTextActive.set(false);
      }
    }
  }

  selectAllElements(): void {
    const slide = this.currentSlide();
    if (!slide) return;
    const ids = slide.elements.map((el) => el.id);
    this.selectedElementIds.set(ids);
    this.selectedElementId.set(ids[0] || null);
    this.editingPlaceholder.set(null);
    this.isTextActive.set(false);
    if (ids.length > 0) {
      this.actionLogger.log('selectElement', { slide: this.currentSlideIndex() + 1 });
    }
  }

  addElement(element: SlideElement): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].elements.push(clone(element));
      this.deck.set(d);
      this.setSelectedElement(element.id);
    }
  }

  updateElement(id: string, updates: Partial<SlideElement>): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      const elIdx = d.slides[idx].elements.findIndex((e) => e.id === id);
      if (elIdx !== -1) {
        d.slides[idx].elements[elIdx] = {
          ...d.slides[idx].elements[elIdx],
          ...updates,
        } as SlideElement;
        this.deck.set(d);
      }
    }
  }

  deleteElement(id: string): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].elements = d.slides[idx].elements.filter((e) => e.id !== id);
      this.deck.set(d);
      this.selectedElementIds.set(this.selectedElementIds().filter((selId) => selId !== id));
      if (this.selectedElementId() === id) {
        this.selectedElementId.set(this.selectedElementIds()[0] || null);
      }
    }
  }

  deleteSelectedElements(): void {
    let ids = this.selectedElementIds();
    if (ids.length === 0 && this.selectedElementId()) {
      ids = [this.selectedElementId()!];
    }
    if (ids.length === 0) return;
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      d.slides[idx].elements = d.slides[idx].elements.filter((e) => !ids.includes(e.id));
      this.deck.set(d);
      this.selectedElementIds.set([]);
      this.selectedElementId.set(null);
      this.isTextActive.set(false);
    }
  }

  updateTableCell(elementId: string, row: number, col: number, text: string): void {
    const idx = this.currentSlideIndex();
    const d = clone(this.deck());
    if (idx >= 0 && idx < d.slides.length) {
      const el = d.slides[idx].elements.find((e) => e.id === elementId) as TableSlideElement | undefined;
      if (el && el.kind === 'table' && el.cells[row] && el.cells[row][col] !== undefined) {
        el.cells[row][col] = text;
        this.deck.set(d);
      }
    }
  }

  copyCurrentSlide(): void {
    const cur = this.currentSlide();
    if (!cur) return;
    this.clipboardSlide.set(clone(cur));
  }

  pasteSlide(): Slide | null {
    const copied = this.clipboardSlide();
    if (!copied) return null;

    const clonedSlide = clone(copied);
    clonedSlide.id = generateSlideId();
    const curIdx = this.currentSlideIndex();
    const d = clone(this.deck());
    const insertAt = Math.min(d.slides.length, curIdx + 1);
    d.slides.splice(insertAt, 0, clonedSlide);

    this.deck.set(d);
    this.setCurrentSlideIndex(insertAt);
    return clonedSlide;
  }

  getSnapshot(): PptSnapshot {
    return {
      deck: clone(this.deck()),
      currentSlideIndex: this.currentSlideIndex(),
      selectedElementId: this.selectedElementId(),
      editingPlaceholder: clone(this.editingPlaceholder()),
    };
  }

  restoreSnapshot(snap: PptSnapshot): void {
    this.deck.set(clone(snap.deck));
    this.currentSlideIndex.set(snap.currentSlideIndex);
    this.selectedElementId.set(snap.selectedElementId);
    this.selectedElementIds.set(snap.selectedElementId ? [snap.selectedElementId] : []);
    this.editingPlaceholder.set(clone(snap.editingPlaceholder));
    this.isTextActive.set(snap.editingPlaceholder !== null);
  }

  reset(): void {
    this.deck.set({
      theme: DEFAULT_THEME_ID,
      slides: [createDefaultSlide('title')],
    });
    this.currentSlideIndex.set(0);
    this.selectedElementId.set(null);
    this.selectedElementIds.set([]);
    this.editingPlaceholder.set(null);
    this.isTextActive.set(false);
    this.clipboardSlide.set(null);
  }
}
