import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
} from '@angular/core';
import { CommonModule } from '@angular/common';

export interface PictureOption {
  file: string;
  alt?: string;
}

@Component({
  selector: 'app-ppt-insert-picture-dialog',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './insert-picture-dialog.component.html',
  styleUrls: ['./insert-picture-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptInsertPictureDialogComponent implements OnChanges {
  @Input() isOpen: boolean = false;
  @Input() images: PictureOption[] = [];
  @Input() assetBaseUrl: string = './assets/';

  @Output() readonly insert = new EventEmitter<string>();
  @Output() readonly close = new EventEmitter<void>();

  selectedFile: string | null = null;
  readonly failedImages = new Set<string>();

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['isOpen'] && this.isOpen) {
      if (this.images && this.images.length > 0) {
        this.selectedFile = this.images[0].file;
      } else {
        this.selectedFile = null;
      }
    }
  }

  getAssetUrl(file: string): string {
    const base = this.assetBaseUrl.endsWith('/') ? this.assetBaseUrl : `${this.assetBaseUrl}/`;
    return `${base}${file}`;
  }

  onImageError(file: string): void {
    this.failedImages.add(file);
  }

  hasImageFailed(file: string): boolean {
    return this.failedImages.has(file);
  }

  onSelectPicture(file: string): void {
    this.selectedFile = file;
  }

  onInsert(): void {
    if (!this.selectedFile) return;
    this.insert.emit(this.selectedFile);
    this.close.emit();
  }

  onCancel(): void {
    this.close.emit();
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    if (this.isOpen) {
      this.close.emit();
    }
  }
}
