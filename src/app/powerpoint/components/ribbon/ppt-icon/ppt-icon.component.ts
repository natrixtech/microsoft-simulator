import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-ppt-icon',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-icon.component.html',
  styleUrls: ['./ppt-icon.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptIconComponent {
  @Input() name: string = '';
  @Input() size: number = 16;
}

