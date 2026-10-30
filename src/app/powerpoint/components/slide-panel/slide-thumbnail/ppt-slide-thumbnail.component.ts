import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptSlideViewComponent } from '../../slide-view/ppt-slide-view.component';
import { Slide, ThemeDef } from '../../../models/ppt.types';

@Component({
  selector: 'app-ppt-slide-thumbnail',
  standalone: true,
  imports: [CommonModule, PptSlideViewComponent],
  templateUrl: './ppt-slide-thumbnail.component.html',
  styleUrls: ['./ppt-slide-thumbnail.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptSlideThumbnailComponent {
  @Input() slide!: Slide;
  @Input() slideIndex: number = 0;
  @Input() isSelected: boolean = false;
  @Input() isDragging: boolean = false;
  @Input() theme!: ThemeDef;
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly selectSlide = new EventEmitter<number>();
  @Output() readonly contextMenu = new EventEmitter<{
    index: number;
    clientX: number;
    clientY: number;
  }>();
  @Output() readonly pointerDownThumb = new EventEmitter<PointerEvent>();

  readonly thumbWidth = 144;
  readonly thumbHeight = 81;
  readonly thumbScale = 144 / 960; // 0.15

  onPointerDown(event: PointerEvent): void {
    if (event.button !== 0) return;
    this.pointerDownThumb.emit(event);
  }

  onSelect(event: MouseEvent): void {
    event.stopPropagation();
    this.selectSlide.emit(this.slideIndex);
  }

  onContextMenu(event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.selectSlide.emit(this.slideIndex);
    this.contextMenu.emit({
      index: this.slideIndex,
      clientX: event.clientX,
      clientY: event.clientY,
    });
  }
}
