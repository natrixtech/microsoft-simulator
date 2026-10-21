import { ChangeDetectionStrategy, Component, EventEmitter, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelIconComponent } from '../excel-icon.component';

@Component({
  selector: 'app-formulas-tab',
  standalone: true,
  imports: [CommonModule, ExcelIconComponent],
  templateUrl: './formulas-tab.component.html',
  styleUrls: ['./formulas-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class FormulasTabComponent {
  @Output() readonly openInsertFunction = new EventEmitter<void>();
  @Output() readonly insertNamedFunction = new EventEmitter<string>();
  @Output() readonly autoSum = new EventEmitter<void>();

  readonly openCategory = signal<string | null>(null);
  readonly catMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });

  toggleCategory(cat: string, event: MouseEvent): void {
    if (this.openCategory() === cat) {
      this.openCategory.set(null);
    } else {
      const el = event.currentTarget as HTMLElement;
      const rect = el.getBoundingClientRect();
      this.catMenuPos.set({ top: rect.bottom + 2, left: rect.left });
      this.openCategory.set(cat);
    }
  }

  onPickFunction(fnName: string): void {
    this.insertNamedFunction.emit(fnName);
    this.openCategory.set(null);
  }
}
