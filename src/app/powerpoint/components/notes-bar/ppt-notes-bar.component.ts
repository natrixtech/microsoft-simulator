import { ChangeDetectionStrategy, Component } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-ppt-notes-bar',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-notes-bar.component.html',
  styleUrls: ['./ppt-notes-bar.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptNotesBarComponent {}

