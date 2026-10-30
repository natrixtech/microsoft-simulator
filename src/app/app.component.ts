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
  private excelComponentRef: ComponentRef<any> | null = null;
  private powerpointComponentRef: ComponentRef<any> | null = null;

  constructor(
    public bridge: PluginBridgeService,
    private ngZone: NgZone,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.bridge.onActivity((act: ActivityPayload) => {
      if (act.app === 'word' && this.loadedApp() !== 'word') {
        this.loadWordEditor();
      } else if (act.app === 'excel' && this.loadedApp() !== 'excel') {
        this.loadExcelEditor();
      } else if (act.app === 'powerpoint' && this.loadedApp() !== 'powerpoint') {
        this.loadPowerpointEditor();
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
          if (this.excelComponentRef) {
            this.excelComponentRef.destroy();
            this.excelComponentRef = null;
          }
          if (this.powerpointComponentRef) {
            this.powerpointComponentRef.destroy();
            this.powerpointComponentRef = null;
          }
          this.editorContainer.clear();
          this.wordComponentRef = this.editorContainer.createComponent(m.WordEditorComponent);
          this.loadedApp.set('word');
          this.isLoading.set(false);
          this.cdr.markForCheck();
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

  private loadExcelEditor(): void {
    this.isLoading.set(true);
    this.cdr.markForCheck();

    import('./excel/excel-editor.component')
      .then((m) => {
        this.ngZone.run(() => {
          if (this.wordComponentRef) {
            this.wordComponentRef.destroy();
            this.wordComponentRef = null;
          }
          if (this.powerpointComponentRef) {
            this.powerpointComponentRef.destroy();
            this.powerpointComponentRef = null;
          }
          this.editorContainer.clear();
          this.excelComponentRef = this.editorContainer.createComponent(m.ExcelEditorComponent);
          this.loadedApp.set('excel');
          this.isLoading.set(false);
          this.cdr.markForCheck();
        });
      })
      .catch((err) => {
        this.ngZone.run(() => {
          this.isLoading.set(false);
          this.bridge.sendError('Failed to load Excel editor: ' + (err?.message || err));
          this.cdr.markForCheck();
        });
      });
  }

  private loadPowerpointEditor(): void {
    this.isLoading.set(true);
    this.cdr.markForCheck();

    import('./powerpoint/powerpoint-editor.component')
      .then((m) => {
        this.ngZone.run(() => {
          if (this.wordComponentRef) {
            this.wordComponentRef.destroy();
            this.wordComponentRef = null;
          }
          if (this.excelComponentRef) {
            this.excelComponentRef.destroy();
            this.excelComponentRef = null;
          }
          this.editorContainer.clear();
          this.powerpointComponentRef = this.editorContainer.createComponent(
            m.PowerpointEditorComponent || m.PptEditorComponent
          );
          this.loadedApp.set('powerpoint');
          this.isLoading.set(false);
          this.cdr.markForCheck();
        });
      })
      .catch((err) => {
        this.ngZone.run(() => {
          this.isLoading.set(false);
          this.bridge.sendError('Failed to load PowerPoint editor: ' + (err?.message || err));
          this.cdr.markForCheck();
        });
      });
  }
}
