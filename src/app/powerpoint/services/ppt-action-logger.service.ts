import { Injectable, signal } from '@angular/core';

export interface PptLoggedAction {
  action: string;
  slide?: number;
  menu?: string;
  [key: string]: any;
  timestamp: number;
}

@Injectable({
  providedIn: 'root',
})
export class PptActionLoggerService {
  readonly actions = signal<PptLoggedAction[]>([]);

  log(action: string, options?: Record<string, any>): void {
    this.actions.update((list) => [
      ...list,
      {
        action,
        ...options,
        timestamp: Date.now(),
      },
    ]);
  }

  hasAction(action: string, options?: { slide?: number; menu?: string }): boolean {
    const list = this.actions();
    return list.some((a) => {
      if (a.action !== action) return false;
      if (options?.slide !== undefined && a.slide !== options.slide) return false;
      if (options?.menu !== undefined && a.menu !== options.menu) return false;
      return true;
    });
  }

  clear(): void {
    this.actions.set([]);
  }
}

