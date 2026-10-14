import { Injectable, NgZone, signal } from '@angular/core';
import {
  ActivityPayload,
  AppKind,
  IncomingPluginMessage,
  PluginEventMessage,
  ThemePreference,
} from '../models/activity.types';
import { validateActivity } from '../validators/contract.validator';

@Injectable({
  providedIn: 'root',
})
export class PluginBridgeService {
  readonly activity = signal<ActivityPayload | null>(null);
  readonly isPaused = signal<boolean>(false);
  readonly theme = signal<ThemePreference>('LIGHT');
  readonly isEmbedded = signal<boolean>(typeof window !== 'undefined' && window.parent !== window);
  readonly isCompleted = signal<boolean>(false);

  private activityCallbacks: Array<(activity: ActivityPayload) => void> = [];
  private forceFinishCallbacks: Array<() => void> = [];

  private hostOrigin: string | null = null;
  private currentKind: AppKind | null = null;

  constructor(private ngZone: NgZone) {}

  onActivity(callback: (activity: ActivityPayload) => void): () => void {
    this.activityCallbacks.push(callback);
    return () => {
      this.activityCallbacks = this.activityCallbacks.filter((cb) => cb !== callback);
    };
  }

  onForceFinish(callback: () => void): () => void {
    this.forceFinishCallbacks.push(callback);
    return () => {
      this.forceFinishCallbacks = this.forceFinishCallbacks.filter((cb) => cb !== callback);
    };
  }

  init(kind: AppKind): void {
    if (typeof window === 'undefined') return;

    this.currentKind = kind;
    const params = new URLSearchParams(window.location.search);
    this.hostOrigin = params.get('parentOrigin') || null;

    if (!this.hostOrigin && document.referrer) {
      try {
        this.hostOrigin = new URL(document.referrer).origin;
      } catch {
        this.hostOrigin = null;
      }
    }

    window.addEventListener('message', (event) => this.handleMessage(event));

    if (this.isEmbedded()) {
      this.sendReady();
    }
  }

  sendReady(): void {
    this.postToParent({
      type: 'PLUGIN_EVENT',
      name: 'pluginReady',
    });
  }

  resetCompleted(): void {
    this.isCompleted.set(false);
  }

  sendComplete(completed: boolean, tasksPassed?: number, tasksTotal?: number): void {
    if (this.isCompleted()) return; // strictly guard so it cannot be sent twice

    this.isCompleted.set(true);

    const payload = {
      completed,
      tasksPassed: tasksPassed ?? (completed ? 1 : 0),
      tasksTotal: tasksTotal ?? 1,
    };

    this.postToParent({
      type: 'PLUGIN_EVENT',
      name: 'activityCompleted',
      payload,
    });
  }

  sendError(message: string): void {
    this.postToParent({
      type: 'PLUGIN_EVENT',
      name: 'pluginError',
      message,
    });
  }

  loadActivity(data: unknown): void {
    if (!this.currentKind) return;
    try {
      const validated = validateActivity(this.currentKind, data);
      this.isCompleted.set(false);
      this.isPaused.set(false);
      this.theme.set(validated.theme || 'LIGHT');
      this.applyThemeToDom(validated.theme || 'LIGHT');
      this.activity.set(validated);
      this.ngZone.run(() => {
        this.activityCallbacks.forEach((cb) => cb(validated));
      });
    } catch (err: any) {
      this.sendError(err?.message || 'Invalid activity payload');
      throw err;
    }
  }

  private handleMessage(event: MessageEvent): void {
    if (!this.isEmbedded() || event.source !== window.parent) return;

    if (this.hostOrigin && this.hostOrigin !== 'null' && event.origin !== this.hostOrigin) {
      return;
    }

    const data = event.data as IncomingPluginMessage;
    if (!data || typeof data !== 'object') return;

    this.ngZone.run(() => {
      try {
        switch (data.type) {
          case 'PLUGIN_DATA_IN':
            if (!this.hostOrigin && event.origin) {
              this.hostOrigin = event.origin;
            }
            this.loadActivity(data.payload);
            break;

          case 'PAUSE':
            this.isPaused.set(!!data.payload?.pause);
            break;

          case 'FORCE_FINISH':
            this.forceFinishCallbacks.forEach((cb) => cb());
            break;

          case 'PREFERENCE_CHANGED':
            if (data.payload?.type === 'theme') {
              this.theme.set(data.payload.value);
              this.applyThemeToDom(data.payload.value);
            }
            break;
        }
      } catch (err: any) {
        this.sendError(err?.message || 'Error processing message');
      }
    });
  }

  private postToParent(message: PluginEventMessage): void {
    if (typeof window === 'undefined') return;

    const targetOrigin =
      this.hostOrigin && this.hostOrigin !== 'null' ? this.hostOrigin : '*';

    try {
      window.parent.postMessage(message, targetOrigin);
    } catch {
      // Ignore postMessage failure in isolated contexts
    }
  }

  private applyThemeToDom(theme: ThemePreference): void {
    if (typeof document !== 'undefined') {
      document.documentElement.dataset['theme'] = theme === 'DARK' ? 'dark' : 'light';
    }
  }
}
