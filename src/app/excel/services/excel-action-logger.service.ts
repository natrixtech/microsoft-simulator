import { Injectable, signal } from '@angular/core';
import { ExcelActionType, LoggedAction } from '../models/excel.types';

@Injectable({
  providedIn: 'root',
})
export class ExcelActionLoggerService {
  readonly actions = signal<LoggedAction[]>([]);

  log(action: ExcelActionType, range?: string): void {
    this.actions.update((list) => [
      ...list,
      {
        action,
        range: range ? range.toUpperCase().replace(/\$/g, '') : undefined,
        timestamp: Date.now(),
      },
    ]);
  }

  hasAction(action: ExcelActionType, targetRange?: string): boolean {
    const list = this.actions();
    if (!targetRange) {
      return list.some((a) => a.action === action);
    }

    const normTarget = targetRange.toUpperCase().replace(/\$/g, '');
    return list.some((a) => {
      if (a.action !== action) return false;
      if (!a.range) return true;
      return a.range === normTarget;
    });
  }

  clear(): void {
    this.actions.set([]);
  }
}
