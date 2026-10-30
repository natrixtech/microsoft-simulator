import { Injectable, NgZone, signal } from '@angular/core';
import { TAB_FOLD_ORDER, TAB_OVERFLOW_ORDER } from '../constants/ppt-ribbon';

@Injectable({
  providedIn: 'root',
})
export class PptRibbonFoldService {
  // Stage 1: compact state per tab (small buttons drop text)
  readonly isCompact = signal<Record<string, boolean>>({
    home: false,
    insert: false,
    design: false,
    transitions: false,
  });

  // Stage 2: folded group IDs per tab
  readonly foldedGroups = signal<Record<string, string[]>>({
    home: [],
    insert: [],
    design: [],
    transitions: [],
  });

  // Stage 3: overflow group IDs per tab
  readonly overflowGroups = signal<Record<string, string[]>>({
    home: [],
    insert: [],
    design: [],
    transitions: [],
  });

  readonly openFoldGroupId = signal<string | null>(null);
  readonly isRibbonOverflowOpen = signal<boolean>(false);

  private resizeObserver: ResizeObserver | null = null;
  private currentTab = 'home';
  private hostEl: HTMLElement | null = null;
  private tabGetter: (() => string) | null = null;
  private rafId: number | null = null;
  private lastWidth = 0;

  constructor(private ngZone: NgZone) {}

  initObserver(host: HTMLElement, tabGetter: () => string): void {
    this.hostEl = host;
    this.tabGetter = tabGetter;
    this.currentTab = tabGetter();

    if (typeof ResizeObserver === 'undefined') return;

    this.resizeObserver = new ResizeObserver((entries) => {
      this.ngZone.run(() => {
        if (!entries || !entries.length) return;
        const width = Math.round(entries[0].contentRect.width);
        if (Math.abs(width - this.lastWidth) < 1) {
          return;
        }
        this.lastWidth = width;
        this.currentTab = this.tabGetter ? this.tabGetter() : this.currentTab;
        this.triggerReFold();
      });
    });

    this.resizeObserver.observe(host);
    this.triggerReFold();
  }

  destroyObserver(): void {
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
      this.resizeObserver = null;
    }
    if (this.rafId !== null) {
      cancelAnimationFrame(this.rafId);
      this.rafId = null;
    }
    this.hostEl = null;
    this.tabGetter = null;
  }

  isTabCompact(tab: string): boolean {
    return !!this.isCompact()[tab];
  }

  isGroupFolded(tab: string, groupId: string): boolean {
    const list = this.foldedGroups()[tab] || [];
    return list.includes(groupId);
  }

  isGroupInOverflow(tab: string, groupId: string): boolean {
    const list = this.overflowGroups()[tab] || [];
    return list.includes(groupId);
  }

  toggleFold(groupId: string): void {
    if (this.openFoldGroupId() === groupId) {
      this.openFoldGroupId.set(null);
    } else {
      this.openFoldGroupId.set(groupId);
      this.isRibbonOverflowOpen.set(false);
    }
  }

  closeFold(): void {
    this.openFoldGroupId.set(null);
  }

  toggleRibbonOverflow(): void {
    this.isRibbonOverflowOpen.update((v) => !v);
    if (this.isRibbonOverflowOpen()) {
      this.openFoldGroupId.set(null);
    }
  }

  closeRibbonOverflow(): void {
    this.isRibbonOverflowOpen.set(false);
  }

  closeAll(): void {
    this.openFoldGroupId.set(null);
    this.isRibbonOverflowOpen.set(false);
  }

  setCurrentTab(tab: string): void {
    this.currentTab = tab;
    this.closeAll();
    this.triggerReFold();
  }

  recalculateFolding(): void {
    this.triggerReFold();
  }

  triggerReFold(): void {
    if (this.rafId !== null) {
      cancelAnimationFrame(this.rafId);
    }

    const tab = this.tabGetter ? this.tabGetter() : this.currentTab;

    // Reset everything for the current tab first
    this.isCompact.update((map) => ({ ...map, [tab]: false }));
    this.foldedGroups.update((map) => ({ ...map, [tab]: [] }));
    this.overflowGroups.update((map) => ({ ...map, [tab]: [] }));
    this.openFoldGroupId.set(null);
    this.isRibbonOverflowOpen.set(false);

    this.rafId = requestAnimationFrame(() => {
      this.rafId = requestAnimationFrame(() => {
        this.rafId = null;
        this.stepFoldCycle(0);
      });
    });
  }

  private stepFoldCycle(step: number): void {
    if (step > 30 || !this.hostEl) return;

    const tab = this.tabGetter ? this.tabGetter() : this.currentTab;
    const body = this.hostEl.classList.contains('ppt-ribbon-body')
      ? this.hostEl
      : (this.hostEl.querySelector('.ppt-ribbon-body') as HTMLElement) || this.hostEl;

    const clientW = body.clientWidth;
    if (clientW === 0) return;

    const isOverflowing = body.scrollWidth > body.clientWidth;
    if (!isOverflowing) {
      return;
    }

    // Stage 1: if not yet compact, switch to compact
    const compactMap = this.isCompact();
    if (!compactMap[tab]) {
      this.isCompact.update((m) => ({ ...m, [tab]: true }));
      this.rafId = requestAnimationFrame(() => {
        this.rafId = null;
        this.stepFoldCycle(step + 1);
      });
      return;
    }

    // Stage 2: fold big groups one at a time
    const foldCandidates = TAB_FOLD_ORDER[tab] || [];
    const currentFolded = [...(this.foldedGroups()[tab] || [])];
    const nextFold = foldCandidates.find((c) => !currentFolded.includes(c));

    if (nextFold) {
      currentFolded.push(nextFold);
      this.foldedGroups.update((map) => ({ ...map, [tab]: currentFolded }));
      this.rafId = requestAnimationFrame(() => {
        this.rafId = null;
        this.stepFoldCycle(step + 1);
      });
      return;
    }

    // Stage 3: move right-most groups into overflow if still overflowing
    const overflowCandidates = TAB_OVERFLOW_ORDER[tab] || [];
    const currentOverflow = [...(this.overflowGroups()[tab] || [])];
    const nextOverflow = overflowCandidates.find((c) => !currentOverflow.includes(c));

    if (nextOverflow) {
      currentOverflow.push(nextOverflow);
      this.overflowGroups.update((map) => ({ ...map, [tab]: currentOverflow }));
      this.rafId = requestAnimationFrame(() => {
        this.rafId = null;
        this.stepFoldCycle(step + 1);
      });
      return;
    }
  }
}
