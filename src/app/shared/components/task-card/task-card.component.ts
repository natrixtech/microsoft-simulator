import {
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  EventEmitter,
  Input,
  OnChanges,
  OnDestroy,
  Output,
  SimpleChanges,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { AppKind } from '../../models/activity.types';
import { MascotComponent, MascotMood } from '../mascot/mascot.component';

export interface MarkupToken {
  type: 'text' | 'code' | 'cell' | 'bold';
  content: string;
}

@Component({
  selector: 'lib-task-card',
  standalone: true,
  imports: [CommonModule, MascotComponent],
  templateUrl: './task-card.component.html',
  styleUrls: ['./task-card.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class TaskCardComponent implements OnChanges, OnDestroy {
  animating = true;
  private animTimer: any = null;

  constructor(private cdr: ChangeDetectorRef) {}

  ngOnChanges(changes: SimpleChanges): void {
    const isFirst =
      (changes['taskIndex']?.isFirstChange() ?? true) &&
      (changes['title']?.isFirstChange() ?? true);

    if (!isFirst && (changes['taskIndex'] || changes['title'] || changes['instructions'])) {
      this.replayAnimations();
    }
  }

  ngOnDestroy(): void {
    if (this.animTimer) {
      clearTimeout(this.animTimer);
      this.animTimer = null;
    }
  }

  replayAnimations(): void {
    this.animating = false;
    this.cdr.detectChanges();
    if (this.animTimer) {
      clearTimeout(this.animTimer);
    }
    this.animTimer = setTimeout(() => {
      this.animating = true;
      this.cdr.detectChanges();
    }, 30);
  }
  @Input() appKind: AppKind = 'word';
  @Input() title: string = '';
  @Input() explanation: string = '';
  @Input() instructions: string[] = [];
  @Input() hasGoals: boolean = true;
  @Input() hasShowMeHow: boolean = false;
  @Input() isSuccess: boolean = false;
  @Input() isDisabled: boolean = false;
  @Input() feedbackMessage: string = '';
  @Input() taskIndex: number = 0;
  @Input() totalTasks: number = 1;
  @Input() isAssessment: boolean = false;
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly checkWork = new EventEmitter<void>();
  @Output() readonly continueNext = new EventEmitter<void>();
  @Output() readonly startAgain = new EventEmitter<void>();
  @Output() readonly showDemo = new EventEmitter<void>();
  @Output() readonly selectReference = new EventEmitter<string>();
  @Output() readonly nudgePosition = new EventEmitter<{ dx: number; dy: number }>();

  get mascotMood(): MascotMood {
    if (this.isSuccess) return 'happy';
    if (this.feedbackMessage) return 'sad';
    return 'normal';
  }

  get eyebrowText(): string {
    if (this.isAssessment) {
      return `TASK ${this.taskIndex + 1} OF ${this.totalTasks} · ${this.title.toUpperCase()}`;
    }
    return `${this.appKind.toUpperCase()} PRACTICE`;
  }

  get buttonLabel(): string {
    if (!this.hasGoals) return 'Done';
    return 'Check My Work';
  }

  parseInstruction(text: string): MarkupToken[] {
    if (!text) return [];
    const tokens: MarkupToken[] = [];
    const parts = text.split(/(`[^`]+`|\b[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?\b|slide \d+|\b\w+ > \w+\b)/gi);

    for (const part of parts) {
      if (!part) continue;
      if (part.startsWith('`') && part.endsWith('`')) {
        const inner = part.slice(1, -1);
        if (this.appKind === 'excel' && /^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/i.test(inner)) {
          tokens.push({ type: 'cell', content: inner });
        } else {
          tokens.push({ type: 'code', content: inner });
        }
      } else if (this.appKind === 'excel' && /^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/i.test(part)) {
        tokens.push({ type: 'cell', content: part });
      } else if (/^slide \d+| > /i.test(part)) {
        tokens.push({ type: 'bold', content: part });
      } else {
        tokens.push({ type: 'text', content: part });
      }
    }
    return tokens;
  }

  onCellClick(cell: string): void {
    this.selectReference.emit(cell);
  }

  onPrimaryAction(): void {
    if (!this.isSuccess) {
      this.checkWork.emit();
    }
  }

  onDragHandleKeydown(event: KeyboardEvent): void {
    const step = 20;
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      this.nudgePosition.emit({ dx: -step, dy: 0 });
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      this.nudgePosition.emit({ dx: step, dy: 0 });
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      this.nudgePosition.emit({ dx: 0, dy: -step });
    } else if (event.key === 'ArrowDown') {
      event.preventDefault();
      this.nudgePosition.emit({ dx: 0, dy: step });
    }
  }
}
