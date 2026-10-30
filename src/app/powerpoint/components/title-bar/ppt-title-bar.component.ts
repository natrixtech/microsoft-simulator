import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-ppt-title-bar',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ppt-title-bar.component.html',
  styleUrls: ['./ppt-title-bar.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptTitleBarComponent {
  @Input() title: string = '';
}

