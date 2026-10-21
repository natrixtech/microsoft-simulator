import {
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
  ViewChild,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelStateService } from '../../services/excel-state.service';
import { ExcelActionLoggerService } from '../../services/excel-action-logger.service';

export interface FormulaCommitPayload {
  value: string;
  dRow?: number;
  dCol?: number;
}

@Component({
  selector: 'app-formula-bar',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './formula-bar.component.html',
  styleUrls: ['./formula-bar.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class FormulaBarComponent implements OnChanges {
  @ViewChild('formulaInput') formulaInputRef?: ElementRef<HTMLInputElement>;
  @ViewChild('nameBoxInput') nameBoxInputRef?: ElementRef<HTMLInputElement>;

  @Input() activeCellAddr: string = 'A1';
  @Input() rawValue: string = '';
  @Input() isEditing: boolean = false;

  @Output() readonly commitEdit = new EventEmitter<string | FormulaCommitPayload>();
  @Output() readonly cancelEdit = new EventEmitter<void>();
  @Output() readonly valueChange = new EventEmitter<string>();
  @Output() readonly openFunctionPicker = new EventEmitter<void>();
  @Output() readonly navigateAddress = new EventEmitter<string>();

  readonly nameBoxValue = signal<string>('A1');
  readonly inputValue = signal<string>('');

  constructor(
    private state: ExcelStateService,
    private actionLogger: ExcelActionLoggerService,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['activeCellAddr']) {
      this.nameBoxValue.set(this.activeCellAddr);
    }
    if (changes['rawValue']) {
      this.inputValue.set(this.rawValue || '');
      this.cdr.markForCheck();
    }
  }

  onNameBoxKeydown(event: KeyboardEvent): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      const val = this.nameBoxValue().trim().toUpperCase();
      if (val) {
        this.navigateAddress.emit(val);
        this.actionLogger.log('goToCell', val);
      }
    } else if (event.key === 'Escape') {
      this.nameBoxValue.set(this.activeCellAddr);
      this.nameBoxInputRef?.nativeElement.blur();
    }
  }

  onInput(event: Event): void {
    const val = (event.target as HTMLInputElement).value;
    this.inputValue.set(val);
    this.valueChange.emit(val);
  }

  onInputKeydown(event: KeyboardEvent): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      this.commitEdit.emit({ value: this.inputValue(), dRow: event.shiftKey ? -1 : 1, dCol: 0 });
    } else if (event.key === 'Tab') {
      event.preventDefault();
      this.commitEdit.emit({ value: this.inputValue(), dRow: 0, dCol: event.shiftKey ? -1 : 1 });
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.cancelEdit.emit();
    }
  }

  onCancel(): void {
    this.cancelEdit.emit();
  }

  onCommit(): void {
    this.commitEdit.emit({ value: this.inputValue(), dRow: 0, dCol: 0 });
  }

  onFxClick(): void {
    if (!this.inputValue().startsWith('=')) {
      this.inputValue.set('=' + this.inputValue());
      this.valueChange.emit(this.inputValue());
    }
    this.openFunctionPicker.emit();
  }
}
