import {
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ComponentRef,
  NgZone,
  OnInit,
  ViewChild,
  ViewContainerRef,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PluginBridgeService } from './shared/services/plugin-bridge.service';
import { ActivityPayload } from './shared/models/activity.types';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class AppComponent implements OnInit {
  @ViewChild('editorContainer', { read: ViewContainerRef, static: true })
  editorContainer!: ViewContainerRef;

  readonly loadedApp = signal<string | null>(null);
  readonly isLoading = signal<boolean>(false);

  private wordComponentRef: ComponentRef<any> | null = null;

  constructor(
    public bridge: PluginBridgeService,
    private ngZone: NgZone,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.bridge.onActivity((act: ActivityPayload) => {
      if (act.app === 'word' && !this.wordComponentRef) {
        this.loadWordEditor();
      }
    });

    this.bridge.init();
  }

  private loadWordEditor(): void {
    this.isLoading.set(true);
    this.cdr.markForCheck();

    import('./word/word-editor.component')
      .then((m) => {
        this.ngZone.run(() => {
          if (!this.wordComponentRef) {
            this.editorContainer.clear();
            this.wordComponentRef = this.editorContainer.createComponent(m.WordEditorComponent);
            this.loadedApp.set('word');
            this.isLoading.set(false);
            this.cdr.markForCheck();
          }
        });
      })
      .catch((err) => {
        this.ngZone.run(() => {
          this.isLoading.set(false);
          this.bridge.sendError('Failed to load Word editor: ' + (err?.message || err));
          this.cdr.markForCheck();
        });
      });
  }
}
