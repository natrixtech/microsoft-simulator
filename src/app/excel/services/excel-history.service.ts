import { Injectable, computed, signal } from '@angular/core';
import { SheetSnapshot } from '../models/excel.types';
import { ExcelActionLoggerService } from './excel-action-logger.service';

@Injectable({
  providedIn: 'root',
})
export class ExcelHistoryService {
  private undoStack = signal<SheetSnapshot[]>([]);
  private redoStack = signal<SheetSnapshot[]>([]);

  readonly canUndo = computed(() => this.undoStack().length > 0);
  readonly canRedo = computed(() => this.redoStack().length > 0);

  constructor(private actionLogger: ExcelActionLoggerService) {}

  push(snapshot: SheetSnapshot): void {
    this.undoStack.update((stack) => [...stack, snapshot]);
    this.redoStack.set([]);
  }

  undo(currentSnapshot: SheetSnapshot): SheetSnapshot | null {
    const stack = this.undoStack();
    if (stack.length === 0) return null;

    const previous = stack[stack.length - 1];
    this.undoStack.set(stack.slice(0, -1));
    this.redoStack.update((rStack) => [...rStack, currentSnapshot]);

    this.actionLogger.log('undo');
    return previous;
  }

  redo(currentSnapshot: SheetSnapshot): SheetSnapshot | null {
    const rStack = this.redoStack();
    if (rStack.length === 0) return null;

    const next = rStack[rStack.length - 1];
    this.redoStack.set(rStack.slice(0, -1));
    this.undoStack.update((stack) => [...stack, currentSnapshot]);

    this.actionLogger.log('redo');
    return next;
  }

  clear(): void {
    this.undoStack.set([]);
    this.redoStack.set([]);
  }
}
