import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-ppt-status-bar',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-status-bar.component.html',
  styleUrls: ['./ppt-status-bar.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptStatusBarComponent {
  @Input() currentSlideIndex: number = 0;
  @Input() totalSlides: number = 1;
  @Input() fitPercent: number = 100;

  @Output() readonly refit = new EventEmitter<void>();
}

