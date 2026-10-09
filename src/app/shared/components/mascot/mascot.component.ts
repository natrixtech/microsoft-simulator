import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';

export type MascotMood = 'normal' | 'sad' | 'happy';

@Component({
  selector: 'lib-mascot',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './mascot.component.html',
  styleUrls: ['./mascot.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class MascotComponent {
  @Input() mood: MascotMood = 'normal';
  @Input() assetBaseUrl: string = './assets/';

  get mascotSrc(): string {
    const base = this.assetBaseUrl ? (this.assetBaseUrl.endsWith('/') ? this.assetBaseUrl : this.assetBaseUrl + '/') : './assets/';
    switch (this.mood) {
      case 'happy':
        return `${base}robot-happy.svg`;
      case 'sad':
        return `${base}robot-sad.svg`;
      case 'normal':
      default:
        return `${base}robot-neutral.svg`;
    }
  }

  get mascotAlt(): string {
    switch (this.mood) {
      case 'happy':
        return 'Happy assistant mascot';
      case 'sad':
        return 'Thinking assistant mascot';
      case 'normal':
      default:
        return 'Assistant mascot';
    }
  }

  get messageText(): string {
    switch (this.mood) {
      case 'happy':
        return 'Great job! Keep it up!';
      case 'sad':
        return 'Take your time. You can try again.';
      case 'normal':
      default:
        return 'A little practice goes a long way.';
    }
  }
}
