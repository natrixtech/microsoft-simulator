import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  forwardRef,
  Input,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptEditorComponent } from '../../powerpoint-editor.component';
import { PptStateService } from '../../services/ppt-state.service';
import { PptHistoryService } from '../../services/ppt-history.service';
import { PptTextFormatService } from '../../services/ppt-text-format.service';
import { PptActionLoggerService } from '../../services/ppt-action-logger.service';
import { PptRibbonFoldService } from '../../services/ppt-ribbon-fold.service';
import { PptGoalEvaluatorService } from '../../services/ppt-goal-evaluator.service';
import { ActivityPayload } from '../../../shared/models/activity.types';
import { Deck } from '../../models/ppt.types';

@Component({
  selector: 'app-ppt-demo-host',
  standalone: true,
  imports: [CommonModule, forwardRef(() => PptEditorComponent)],
  providers: [
    PptStateService,
    PptHistoryService,
    PptTextFormatService,
    PptActionLoggerService,
    PptRibbonFoldService,
    PptGoalEvaluatorService,
  ],
  template: `
    <div
      class="ppt-demo-host-container"
      [style.width.px]="1380"
      [style.height.px]="height"
      style="pointer-events: none; position: relative; overflow: hidden; background: var(--ppt-bg-window, #f3f2f1); display: flex; flex-direction: column;"
    >
      <app-ppt-editor
        [demoMode]="true"
        [demoDeck]="demoDeck"
        [demoActivity]="demoActivity"
        style="width: 100%; height: 100%; display: flex; flex-direction: column;"
      ></app-ppt-editor>
    </div>
  `,
  styles: [
    `
      :host {
        display: block;
        width: 1380px;
        height: auto;
        transform-origin: top left;
      }
      .ppt-demo-host-container {
        width: 1380px;
      }
    `,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptDemoHostComponent {
  @Input() height = 860;
  @Input() demoDeck: Deck | null = null;
  @Input() demoActivity: ActivityPayload | null = null;
  @ViewChild(PptEditorComponent) editorComponent?: PptEditorComponent;

  constructor(
    public elRef: ElementRef<HTMLElement>,
    public state: PptStateService,
    public history: PptHistoryService,
    public textFormat: PptTextFormatService,
    public actionLogger: PptActionLoggerService,
    public goalEvaluator: PptGoalEvaluatorService,
    public ribbonFold: PptRibbonFoldService
  ) {}
}
