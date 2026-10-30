import { Injectable, computed, signal } from '@angular/core';
import { PptSnapshot } from '../models/ppt.types';

@Injectable({
  providedIn: 'root',
})
export class PptHistoryService {
  private undoStack = signal<PptSnapshot[]>([]);
  private redoStack = signal<PptSnapshot[]>([]);

  readonly canUndo = computed(() => this.undoStack().length > 0);
  readonly canRedo = computed(() => this.redoStack().length > 0);

  push(snapshot: PptSnapshot): void {
    this.undoStack.update((stack) => [...stack, snapshot]);
    this.redoStack.set([]);
  }

  undo(currentSnapshot: PptSnapshot): PptSnapshot | null {
    const stack = this.undoStack();
    if (stack.length === 0) return null;

    const previous = stack[stack.length - 1];
    this.undoStack.set(stack.slice(0, -1));
    this.redoStack.update((rStack) => [...rStack, currentSnapshot]);

    return previous;
  }

  redo(currentSnapshot: PptSnapshot): PptSnapshot | null {
    const rStack = this.redoStack();
    if (rStack.length === 0) return null;

    const next = rStack[rStack.length - 1];
    this.redoStack.set(rStack.slice(0, -1));
    this.undoStack.update((stack) => [...stack, currentSnapshot]);

    return next;
  }

  clear(): void {
    this.undoStack.set([]);
    this.redoStack.set([]);
  }
}

