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
  @Input() assetBaseUrl?: string;

  get mascotSrc(): string {
    switch (this.mood) {
      case 'happy':
        return './assets/robot-happy.svg';
      case 'sad':
        return './assets/robot-sad.svg';
      case 'normal':
      default:
        return './assets/robot-neutral.svg';
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
