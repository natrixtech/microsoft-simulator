import { ChangeDetectionStrategy, Component, EventEmitter, Input, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelStateService } from '../../services/excel-state.service';

@Component({
  selector: 'app-comment-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './comment-dialog.component.html',
  styleUrls: ['./comment-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class CommentDialogComponent implements OnInit {
  @Input() cellAddress: string = 'A1';
  @Output() readonly saveComment = new EventEmitter<{ cell: string; text: string }>();
  @Output() readonly deleteComment = new EventEmitter<string>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly commentText = signal<string>('');

  constructor(private state: ExcelStateService) {}

  ngOnInit(): void {
    const existing = this.state.comments().get(this.cellAddress.toUpperCase()) || '';
    this.commentText.set(existing);
  }

  onSave(): void {
    this.saveComment.emit({
      cell: this.cellAddress.toUpperCase(),
      text: this.commentText().trim(),
    });
    this.closeDialog.emit();
  }

  onDelete(): void {
    this.deleteComment.emit(this.cellAddress.toUpperCase());
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
