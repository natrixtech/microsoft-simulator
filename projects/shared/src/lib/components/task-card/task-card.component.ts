import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
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
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="card-container">
      <section class="task-card" [class.success]="isSuccess">
        <!-- Drag Handle Header Area -->
        <div
          class="card-drag-handle"
          role="button"
          tabindex="0"
          aria-label="Movable task card. Use arrow keys to reposition."
          (keydown)="onDragHandleKeydown($event)"
          title="Drag to reposition card (Arrow keys to move)"
        >
          <div class="drag-grip-indicator" aria-hidden="true">
            <span class="grip-dot"></span>
            <span class="grip-dot"></span>
            <span class="grip-dot"></span>
            <span class="grip-dot"></span>
            <span class="grip-dot"></span>
            <span class="grip-dot"></span>
          </div>
          <span class="lesson-eyebrow">{{ eyebrowText }}</span>
          <span class="drag-pill" aria-hidden="true">Move ✥</span>
        </div>

        <h2>{{ title }}</h2>
        <p class="explanation">{{ explanation }}</p>

        <div class="task-box">
          <span class="task-label">YOUR TASK</span>
          <ng-container *ngIf="instructions.length === 1; else multiStep">
            <p>
              <ng-container *ngFor="let token of parseInstruction(instructions[0])">
                <code *ngIf="token.type === 'code'">{{ token.content }}</code>
                <button
                  type="button"
                  class="inline-reference"
                  *ngIf="token.type === 'cell'"
                  (click)="onCellClick(token.content)"
                >
                  {{ token.content }}
                </button>
                <strong *ngIf="token.type === 'bold'">{{ token.content }}</strong>
                <span *ngIf="token.type === 'text'">{{ token.content }}</span>
              </ng-container>
            </p>
          </ng-container>
          <ng-template #multiStep>
            <ol>
              <li *ngFor="let step of instructions">
                <ng-container *ngFor="let token of parseInstruction(step)">
                  <code *ngIf="token.type === 'code'">{{ token.content }}</code>
                  <button
                    type="button"
                    class="inline-reference"
                    *ngIf="token.type === 'cell'"
                    (click)="onCellClick(token.content)"
                  >
                    {{ token.content }}
                  </button>
                  <strong *ngIf="token.type === 'bold'">{{ token.content }}</strong>
                  <span *ngIf="token.type === 'text'">{{ token.content }}</span>
                </ng-container>
              </li>
            </ol>
          </ng-template>
        </div>

        <button
          type="button"
          class="help-button"
          *ngIf="hasShowMeHow"
          (click)="showDemo.emit()"
        >
          ▷ &nbsp; Show me how
        </button>

        <div
          class="feedback"
          *ngIf="feedbackMessage"
          [class.good]="isSuccess"
          [class.bad]="!isSuccess"
          role="status"
        >
          {{ feedbackMessage }}
        </div>

        <button
          type="button"
          class="primary check-button"
          [disabled]="isDisabled"
          (click)="onPrimaryAction()"
        >
          {{ buttonLabel }}
        </button>
      </section>

      <!-- Associated Mascot (Moves with Task Card) -->
      <lib-mascot [mood]="mascotMood" [assetBaseUrl]="assetBaseUrl"></lib-mascot>
    </div>
  `,
  styles: [`
    .card-container {
      display: flex;
      flex-direction: column;
      gap: 12px;
      width: 100%;
      max-width: 100%;
      user-select: text;
    }
    .task-card {
      position: relative;
      background: var(--surface, #ffffff);
      border: 1px solid var(--border-color, #d5deeb);
      border-radius: 12px;
      padding: 16px 18px;
      box-shadow: 0 6px 20px rgba(32, 48, 75, 0.1);
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    /* Speech-bubble pointer down to mascot */
    .task-card::after {
      content: "";
      position: absolute;
      bottom: -9px;
      left: 36px;
      width: 0;
      height: 0;
      border-left: 8px solid transparent;
      border-right: 8px solid transparent;
      border-top: 9px solid var(--surface, #ffffff);
      filter: drop-shadow(0 2px 2px rgba(32, 48, 75, 0.08));
      pointer-events: none;
    }
    .task-card.success {
      border-color: #34a853;
    }
    .card-drag-handle {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      padding: 4px 6px;
      margin: -6px -6px 2px;
      border-radius: 6px;
      cursor: grab;
      touch-action: none;
      user-select: none;
      background: rgba(0, 0, 0, 0.02);
      border: 1px dashed rgba(0, 0, 0, 0.1);
      outline: none;
      transition: background 0.15s ease, border-color 0.15s ease;
    }
    .card-drag-handle:hover,
    .card-drag-handle:focus-visible {
      background: rgba(39, 96, 209, 0.06);
      border-color: var(--primary, #2760d1);
    }
    .card-drag-handle:active {
      cursor: grabbing;
    }
    .drag-grip-indicator {
      display: grid;
      grid-template-columns: repeat(3, 4px);
      gap: 3px;
      align-items: center;
      opacity: 0.55;
    }
    .grip-dot {
      width: 3.5px;
      height: 3.5px;
      border-radius: 50%;
      background: #627288;
    }
    .lesson-eyebrow {
      font-size: 10.5px;
      letter-spacing: 1.2px;
      font-weight: 700;
      color: var(--primary, #2760d1);
      text-transform: uppercase;
      flex: 1;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .drag-pill {
      font-size: 10px;
      font-weight: 600;
      color: #627288;
      background: rgba(0, 0, 0, 0.04);
      padding: 2px 6px;
      border-radius: 4px;
    }
    h2 {
      font-size: 18px;
      font-weight: 650;
      margin: 0;
      color: var(--text-heading, #1f2d3d);
    }
    .explanation {
      font-size: 13px;
      line-height: 1.45;
      margin: 0;
      color: var(--text-muted, #5f6e82);
    }
    .task-box {
      background: var(--surface-secondary, #f4f7fc);
      border-radius: 8px;
      padding: 12px 14px;
    }
    .task-label {
      font-size: 10px;
      font-weight: 750;
      letter-spacing: 1.2px;
      color: var(--text-muted, #627288);
      display: block;
      margin-bottom: 6px;
    }
    p, ol {
      margin: 0;
      font-size: 13px;
      line-height: 1.5;
      color: var(--text-main, #24334a);
    }
    ol {
      padding-left: 20px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    code {
      font-family: Consolas, "Cascadia Code", monospace;
      background: #e4ecf7;
      border: 1px solid #c9d7ec;
      color: #174291;
      padding: 1px 5px;
      border-radius: 4px;
      font-size: 12px;
    }
    strong {
      color: var(--text-heading, #1f2d3d);
    }
    .inline-reference {
      background: #e2edfc;
      border: 1px solid #b5cced;
      color: #174291;
      padding: 1px 6px;
      border-radius: 4px;
      font-weight: 600;
      font-size: 12px;
      cursor: pointer;
      display: inline-block;
      margin: 0 2px;
    }
    .inline-reference:hover {
      background: #cddff8;
    }
    .help-button {
      background: none;
      border: none;
      padding: 2px 0;
      color: var(--primary, #2760d1);
      font-weight: 650;
      font-size: 13px;
      cursor: pointer;
      text-align: left;
    }
    .feedback {
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 13px;
      line-height: 1.4;
    }
    .feedback.good {
      background: #e6f4ea;
      color: #137333;
      border: 1px solid #a8dab5;
    }
    .feedback.bad {
      background: #fef7e0;
      color: #8f5300;
      border: 1px solid #f9d984;
    }
    .primary {
      background: var(--primary, #2760d1);
      color: white;
      border: none;
      border-radius: 8px;
      padding: 11px;
      font-size: 13.5px;
      font-weight: 650;
      cursor: pointer;
      transition: background 0.15s ease;
    }
    .primary:hover:not(:disabled) {
      background: #1c4fb3;
    }
    .primary:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    @media (max-height: 800px) {
      .card-container {
        gap: 6px;
      }
      .task-card {
        padding: 10px 12px;
        gap: 6px;
      }
      h2 {
        font-size: 15px;
      }
      .explanation {
        font-size: 12px;
      }
      .task-box {
        padding: 8px 10px;
      }
      p, ol {
        font-size: 12px;
      }
      .primary {
        padding: 8px;
        font-size: 12.5px;
      }
    }
  `],
})
export class TaskCardComponent {
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
    if (this.isSuccess) return 'Continue →';
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
    if (this.isSuccess) {
      this.continueNext.emit();
    } else {
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
