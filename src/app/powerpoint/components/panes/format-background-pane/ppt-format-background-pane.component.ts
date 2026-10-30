import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnChanges,
  OnInit,
  Output,
  SimpleChanges,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptColorPickerComponent } from '../../menus/color-picker/ppt-color-picker.component';
import { computePopupPosition, PopupPosition } from '../../../utils/popup-position';

@Component({
  selector: 'app-ppt-format-background-pane',
  standalone: true,
  imports: [CommonModule, PptColorPickerComponent],
  templateUrl: './ppt-format-background-pane.component.html',
  styleUrls: ['./ppt-format-background-pane.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptFormatBackgroundPaneComponent implements OnInit, OnChanges {
  @ViewChild('colorBtn', { static: false }) colorBtnRef?: ElementRef<HTMLElement>;

  @Input() currentBackground?: string;
  @Input() defaultBackground: string = '#ffffff';

  @Output() readonly close = new EventEmitter<void>();
  @Output() readonly setBackground = new EventEmitter<string | undefined>();
  @Output() readonly openPictureDialog = new EventEmitter<void>();
  @Output() readonly resetBackground = new EventEmitter<void>();

  fillType: 'solid' | 'picture' = 'solid';
  isColorPickerOpen = false;
  colorPickerPos: PopupPosition = { left: 0, top: 0 };

  ngOnInit(): void {
    this.updateFillType();
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['currentBackground']) {
      this.updateFillType();
    }
  }

  private updateFillType(): void {
    if (this.currentBackground && (this.currentBackground.includes('.') || this.currentBackground.includes('/'))) {
      this.fillType = 'picture';
    } else {
      this.fillType = 'solid';
    }
  }

  onSelectFillType(type: 'solid' | 'picture'): void {
    this.fillType = type;
    this.isColorPickerOpen = false;
  }

  onToggleColorPicker(event: MouseEvent): void {
    event.stopPropagation();
    const btn = event.currentTarget as HTMLElement;
    this.colorPickerPos = computePopupPosition(btn, { width: 176, height: 210 });
    this.isColorPickerOpen = !this.isColorPickerOpen;
  }

  onPickColor(color: string): void {
    this.isColorPickerOpen = false;
    this.setBackground.emit(color === 'transparent' ? undefined : color);
  }

  onInsertPicture(): void {
    this.openPictureDialog.emit();
  }

  onReset(): void {
    this.isColorPickerOpen = false;
    this.resetBackground.emit();
  }
}

