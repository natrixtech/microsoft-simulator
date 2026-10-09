import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';

export type MascotMood = 'normal' | 'sad' | 'happy';

@Component({
  selector: 'lib-mascot',
  standalone: true,
  imports: [CommonModule],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="mascot-area">
      <div class="mascot-figure" [attr.data-mood]="mood" aria-label="Practice helper robot">
        <img
          [src]="mascotSrc"
          [alt]="mascotAlt"
          class="mascot-svg"
          loading="eager"
        />
      </div>
      <p class="mascot-message">
        {{ messageText }}
      </p>
    </div>
  `,
  styles: [`
    .mascot-area {
      display: flex;
      align-items: center;
      gap: 12px;
      margin: 12px 0 0;
      color: var(--text-muted, #627087);
      font-size: 13px;
      line-height: 1.4;
    }
    @media (max-height: 800px) {
      .mascot-area {
        margin: 6px 0 0;
        gap: 8px;
        font-size: 11.5px;
      }
      .mascot-svg {
        width: 44px;
        height: auto;
      }
    }
    .mascot-figure {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
      transition: transform 0.2s cubic-bezier(0.34, 1.56, 0.64, 1);
    }
    .mascot-figure[data-mood='happy'] {
      transform: scale(1.05);
    }
    .mascot-figure[data-mood='sad'] {
      transform: translateY(2px);
    }
    .mascot-svg {
      width: 52px;
      height: auto;
      aspect-ratio: 228 / 204;
      object-fit: contain;
      display: block;
      filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.08));
    }
    .mascot-message {
      margin: 0;
      flex: 1;
      font-weight: 500;
    }
  `],
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
