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

@Component({
  selector: 'app-ppt-slide-context-menu',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-context-menu.component.html',
  styleUrls: ['./slide-context-menu.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptSlideContextMenuComponent {
  @Input() isOpen: boolean = false;
  @Input() x: number = 0;
  @Input() y: number = 0;
  @Input() canDelete: boolean = true;
  @Input() canPaste: boolean = false;

  @Output() readonly newSlide = new EventEmitter<void>();
  @Output() readonly deleteSlide = new EventEmitter<void>();
  @Output() readonly copy = new EventEmitter<void>();
  @Output() readonly paste = new EventEmitter<void>();
  @Output() readonly close = new EventEmitter<void>();

  constructor(private elementRef: ElementRef) {}

  @HostListener('document:click', ['$event'])
  onDocClick(event: MouseEvent): void {
    if (this.isOpen && !this.elementRef.nativeElement.contains(event.target)) {
      this.close.emit();
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.isOpen) {
      this.close.emit();
    }
  }

  onAction(action: 'new' | 'delete' | 'copy' | 'paste'): void {
    this.close.emit();
    if (action === 'new') this.newSlide.emit();
    else if (action === 'delete') {
      if (this.canDelete) this.deleteSlide.emit();
    } else if (action === 'copy') this.copy.emit();
    else if (action === 'paste') {
      if (this.canPaste) this.paste.emit();
    }
  }
}

