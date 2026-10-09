import { Injectable, signal } from '@angular/core';
import { ThemePreference } from '../models/activity.types';

@Injectable({
  providedIn: 'root',
})
export class ThemeService {
  readonly currentTheme = signal<ThemePreference>('LIGHT');

  setTheme(theme: ThemePreference): void {
    this.currentTheme.set(theme);
    if (typeof document !== 'undefined') {
      document.documentElement.dataset['theme'] = theme === 'DARK' ? 'dark' : 'light';
    }
  }

  toggleTheme(): void {
    this.setTheme(this.currentTheme() === 'LIGHT' ? 'DARK' : 'LIGHT');
  }
}
