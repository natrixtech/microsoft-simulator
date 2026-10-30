import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { COLORS } from '../../../constants/ppt-colors';

export interface ColorSwatch {
  name: string;
  label: string;
  hex: string;
}

@Component({
  selector: 'app-ppt-color-picker',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-color-picker.component.html',
  styleUrls: ['./ppt-color-picker.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptColorPickerComponent {
  @Input() allowTransparent: boolean = false;
  @Input() selectedColor: string = '';

  @Output() readonly selectColor = new EventEmitter<string>();
  @Output() readonly close = new EventEmitter<void>();

  readonly colorList: ColorSwatch[] = Object.entries(COLORS).map(([name, hex]) => ({
    name,
    label: this.formatLabel(name),
    hex: String(hex),
  }));

  constructor(private elementRef: ElementRef) {}

  @HostListener('document:click', ['$event'])
  onDocClick(event: MouseEvent): void {
    if (!this.elementRef.nativeElement.contains(event.target)) {
      this.close.emit();
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.close.emit();
  }

  onPickColor(color: string, event: MouseEvent): void {
    event.stopPropagation();
    this.selectColor.emit(color);
  }

  onPickTransparent(event: MouseEvent): void {
    event.stopPropagation();
    this.selectColor.emit('transparent');
  }

  private formatLabel(key: string): string {
    const withSpaces = key.replace(/([A-Z])/g, ' $1');
    return withSpaces.charAt(0).toUpperCase() + withSpaces.slice(1);
  }
}
