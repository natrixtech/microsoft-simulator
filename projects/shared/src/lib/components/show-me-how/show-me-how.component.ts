import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnDestroy,
  OnInit,
  Output,
  computed,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { TaskGoal } from '../../models/activity.types';

export interface DemoStep {
  label: string;
  state: 'pending' | 'active' | 'done';
}

@Component({
  selector: 'lib-show-me-how',
  standalone: true,
  imports: [CommonModule],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="demo-backdrop" (click)="onBackdropClick($event)" role="dialog" aria-modal="true">
      <div class="demo-dialog" (click)="$event.stopPropagation()">
        <!-- Header: Task Title, First Instruction, and Close button only (No extra labels or bottom buttons) -->
        <header class="dialog-header">
          <div class="header-titles">
            <h2>{{ title }}</h2>
            <p class="header-instruction" *ngIf="instruction">{{ instruction }}</p>
          </div>
          <button type="button" class="close-button" (click)="onClose()" aria-label="Close">×</button>
        </header>

        <div class="demo-body">
          <!-- Left Column: Steps Checklist with Numbered Circles and Replay Button -->
          <div class="demo-steps-col">
            <ol class="demo-steps">
              <li *ngFor="let step of steps(); let i = index" [class]="step.state">
                <span class="step-num">
                  {{ step.state === 'done' ? '✓' : i + 1 }}
                </span>
                <span class="step-label">{{ step.label }}</span>
              </li>
            </ol>
            <button type="button" class="replay-btn" (click)="replay()">
              ↺ Replay
            </button>
          </div>

          <!-- Right Area: Scaled-down NON-interactive copy of the real app with animated cursor -->
          <div class="demo-replica-area">
            <div class="demo-toast" *ngIf="showToast()">
              ✨ That’s how you {{ getToastTitle() }}!
            </div>

            <!-- Replica Window Container -->
            <div class="replica-window" [attr.data-app]="appKind">
              <!-- Word 365 Blue Title Bar -->
              <ng-container *ngIf="appKind === 'word'">
                <div class="replica-word-titlebar">
                  <div class="rw-titlebar-left">
                    <span class="rw-app-icon" aria-hidden="true" title="Microsoft Word">
                      <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor">
                        <path d="M19.5 3h-15C3.67 3 3 3.67 3 4.5v15c0 .83.67 1.5 1.5 1.5h15c.83 0 1.5-.67 1.5-1.5v-15c0-.83-.67-1.5-1.5-1.5zm-6.25 14h-1.56l-1.9-6.8-1.9 6.8H6.33L4.5 9h1.72l1.1 5.3 1.83-5.3h1.7l1.83 5.3 1.1-5.3h1.72l-1.75 8z"/>
                      </svg>
                    </span>
                    <strong class="rw-doc-name">{{ title }} - Word</strong>
                  </div>
                </div>

                <!-- Word Ribbon Tabs -->
                <div class="replica-ribbon-tabs">
                  <span class="rw-tab">File</span>
                  <span class="rw-tab active">Home</span>
                  <span class="rw-tab" data-demo-target="tab-insert">Insert</span>
                  <span class="rw-tab" data-demo-target="tab-layout">Layout</span>
                  <span class="rw-tab disabled">Review</span>
                  <span class="rw-tab disabled">View</span>
                </div>

                <!-- Word Ribbon Command Bar with Targets -->
                <div class="replica-ribbon-commands">
                  <!-- Clipboard -->
                  <div class="rw-group">
                    <div class="rw-group-row">
                      <span class="rw-btn" data-demo-target="paste" [class.highlight]="activeAction() === 'paste'">Paste</span>
                      <span class="rw-btn micro" data-demo-target="cut" [class.highlight]="activeAction() === 'cut'">✂</span>
                      <span class="rw-btn micro" data-demo-target="copy" [class.highlight]="activeAction() === 'copy'">📋</span>
                    </div>
                    <span class="rw-group-label">Clipboard</span>
                  </div>

                  <!-- Font -->
                  <div class="rw-group">
                    <div class="rw-group-row">
                      <span class="rw-select" data-demo-target="font">{{ replicaFont() }}</span>
                      <span class="rw-select micro" data-demo-target="size">{{ replicaFontSize() }}</span>
                    </div>
                    <div class="rw-group-row">
                      <span class="rw-btn micro" data-demo-target="bold" [class.highlight]="activeAction() === 'bold'"><strong>B</strong></span>
                      <span class="rw-btn micro" data-demo-target="italic" [class.highlight]="activeAction() === 'italic'"><em>I</em></span>
                      <span class="rw-btn micro" data-demo-target="underline" [class.highlight]="activeAction() === 'underline'"><u>U</u></span>
                      <span class="rw-btn micro" data-demo-target="color" [class.highlight]="activeAction() === 'color'">A <small>▼</small></span>
                      <span class="rw-btn micro" data-demo-target="highlight" [class.highlight]="activeAction() === 'highlight'">🖊 <small>▼</small></span>
                    </div>
                    <span class="rw-group-label">Font</span>
                  </div>

                  <!-- Paragraph -->
                  <div class="rw-group">
                    <div class="rw-group-row">
                      <span class="rw-btn micro" data-demo-target="bullets" [class.highlight]="activeAction() === 'bullets'">•≡</span>
                      <span class="rw-btn micro" data-demo-target="numbering" [class.highlight]="activeAction() === 'numbering'">1≡</span>
                      <span class="rw-btn micro" data-demo-target="indent" [class.highlight]="activeAction() === 'indent'">⇥</span>
                    </div>
                    <div class="rw-group-row">
                      <span class="rw-btn micro" data-demo-target="align-left" [class.highlight]="activeAction() === 'align-left'">⯸</span>
                      <span class="rw-btn micro" data-demo-target="align-center" [class.highlight]="activeAction() === 'align-center'">≡</span>
                      <span class="rw-btn micro" data-demo-target="align-right" [class.highlight]="activeAction() === 'align-right'">⯹</span>
                      <span class="rw-btn micro" data-demo-target="align-justify" [class.highlight]="activeAction() === 'align-justify'">≣</span>
                      <span class="rw-btn micro" data-demo-target="line-spacing" [class.highlight]="activeAction() === 'line-spacing'">↕</span>
                    </div>
                    <span class="rw-group-label">Paragraph</span>
                  </div>

                  <!-- Styles Gallery -->
                  <div class="rw-group rw-styles-group">
                    <div class="rw-styles-strip">
                      <span class="rw-style-tile" data-demo-target="style-normal" [class.active]="replicaStyle() === 'normal'">Normal</span>
                      <span class="rw-style-tile" data-demo-target="style-title" [class.active]="replicaStyle() === 'title'">Title</span>
                      <span class="rw-style-tile" data-demo-target="style-subtitle" [class.active]="replicaStyle() === 'subtitle'">Subtitle</span>
                      <span class="rw-style-tile" data-demo-target="style-h1" [class.active]="replicaStyle() === 'h1'">Heading 1</span>
                      <span class="rw-style-tile" data-demo-target="style-h2" [class.active]="replicaStyle() === 'h2'">Heading 2</span>
                      <span class="rw-style-tile" data-demo-target="style-h3" [class.active]="replicaStyle() === 'h3'">Heading 3</span>
                    </div>
                    <span class="rw-group-label">Styles</span>
                  </div>

                  <!-- Editing -->
                  <div class="rw-group">
                    <div class="rw-group-row vertical">
                      <span class="rw-btn micro-text" data-demo-target="find" [class.highlight]="activeAction() === 'find'">🔍 Find</span>
                      <span class="rw-btn micro-text" data-demo-target="replace" [class.highlight]="activeAction() === 'replace'">⇄ Replace</span>
                    </div>
                    <span class="rw-group-label">Editing</span>
                  </div>

                  <!-- Working Insert Tools -->
                  <div class="rw-group">
                    <div class="rw-group-row">
                      <span class="rw-btn micro-text" data-demo-target="insert-image" [class.highlight]="activeAction() === 'insert-image'">🖼 Pic</span>
                      <span class="rw-btn micro-text" data-demo-target="insert-table" [class.highlight]="activeAction() === 'insert-table'">田 Table</span>
                      <span class="rw-btn micro-text" data-demo-target="page-setup" [class.highlight]="activeAction() === 'page-setup'">📄 Setup</span>
                    </div>
                    <span class="rw-group-label">Insert</span>
                  </div>
                </div>

                <!-- Word Document Canvas Sheet -->
                <div class="replica-doc-workspace">
                  <div class="replica-doc-ruler">
                    <span>1</span><span>2</span><span>3</span><span>4</span><span>5</span><span>6</span><span>7</span>
                  </div>
                  <div class="replica-doc-page">
                    <div
                      class="replica-doc-content"
                      data-demo-target="doc-canvas"
                      [innerHTML]="replicaDocHtml()"
                    ></div>
                  </div>
                </div>
              </ng-container>

              <!-- Excel Replica -->
              <ng-container *ngIf="appKind === 'excel'">
                <div class="replica-excel-bar">
                  <span class="re-tag">Excel</span>
                  <div class="re-tools">
                    <span class="rw-btn micro" data-demo-target="excel-bold"><strong>B</strong></span>
                    <span class="rw-btn micro" data-demo-target="excel-currency">$</span>
                    <span class="rw-btn micro" data-demo-target="excel-percent">%</span>
                    <span class="rw-btn micro" data-demo-target="excel-sum">∑</span>
                    <span class="rw-btn micro" data-demo-target="excel-sort">A→Z</span>
                    <span class="rw-btn micro" data-demo-target="excel-chart">📊</span>
                  </div>
                </div>
                <div class="replica-formula-bar">
                  <span class="r-name-box">{{ replicaFormulaCell() }}</span>
                  <span class="r-fx">fx</span>
                  <span class="r-fx-input">{{ replicaFormulaText() }}</span>
                </div>
                <div class="replica-sheet-wrap">
                  <table class="replica-sheet-grid">
                    <thead>
                      <tr><th></th><th>A</th><th>B</th><th>C</th></tr>
                    </thead>
                    <tbody>
                      <tr *ngFor="let row of [1, 2, 3, 4, 5]">
                        <th>{{ row }}</th>
                        <td [class.selected]="isCellSelected('A' + row)" [attr.data-demo-target]="'cell-A' + row">{{ getReplicaCell('A' + row) }}</td>
                        <td [class.selected]="isCellSelected('B' + row)" [attr.data-demo-target]="'cell-B' + row">{{ getReplicaCell('B' + row) }}</td>
                        <td [class.selected]="isCellSelected('C' + row)" [attr.data-demo-target]="'cell-C' + row">{{ getReplicaCell('C' + row) }}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </ng-container>

              <!-- PowerPoint Replica -->
              <ng-container *ngIf="appKind === 'powerpoint'">
                <div class="replica-ppt-bar">
                  <span class="rp-tag">PowerPoint</span>
                  <div class="rp-tools">
                    <span class="rw-btn micro" data-demo-target="ppt-newslide">+ Slide</span>
                    <span class="rw-btn micro" data-demo-target="ppt-theme">Theme</span>
                    <span class="rw-btn micro" data-demo-target="ppt-bullets">• List</span>
                    <span class="rw-btn micro" data-demo-target="ppt-shape">Shape</span>
                  </div>
                </div>
                <div class="replica-ppt-layout">
                  <div class="replica-ppt-thumbs">
                    <div class="thumb active">1</div>
                    <div class="thumb" *ngIf="replicaSlideCount() > 1">2</div>
                  </div>
                  <div class="replica-ppt-slide" [class.streamline]="replicaTheme() === 'streamline'">
                    <div class="r-slide-title" data-demo-target="ppt-title" [class.selected]="activeAction() === 'editTitle'">
                      {{ replicaSlideTitle() }}
                    </div>
                    <div class="r-slide-body" data-demo-target="ppt-body">
                      <ul *ngIf="replicaSlideBulletList()">
                        <li>Learn key concepts</li>
                        <li>Create dynamic slides</li>
                      </ul>
                      <p *ngIf="!replicaSlideBulletList()">Learn key concepts<br>Create dynamic slides</p>
                    </div>
                  </div>
                </div>
              </ng-container>

              <!-- In-Replica Dialogs Driven Step by Step -->
              <div class="replica-dialog-overlay" *ngIf="replicaDialog() === 'findReplace'">
                <div class="replica-dialog-box">
                  <div class="rd-header">Find and Replace</div>
                  <div class="rd-body">
                    <div class="rd-field">
                      <label>Find what:</label>
                      <input
                        type="text"
                        [value]="findWhatInput()"
                        data-demo-target="find-what-input"
                        readonly
                      />
                    </div>
                    <div class="rd-field">
                      <label>Replace with:</label>
                      <input
                        type="text"
                        [value]="replaceWithInput()"
                        data-demo-target="replace-with-input"
                        readonly
                      />
                    </div>
                    <div class="rd-actions">
                      <span class="rd-btn primary" data-demo-target="dialog-replace-btn">Replace</span>
                      <span class="rd-btn" data-demo-target="dialog-close-btn">Close</span>
                    </div>
                  </div>
                </div>
              </div>

              <div class="replica-dialog-overlay" *ngIf="replicaDialog() === 'pageSetup'">
                <div class="replica-dialog-box">
                  <div class="rd-header">Page Setup</div>
                  <div class="rd-body">
                    <div class="rd-field"><label>Orientation:</label><span>Landscape</span></div>
                    <div class="rd-actions">
                      <span class="rd-btn primary" data-demo-target="pagesetup-ok-btn">OK</span>
                    </div>
                  </div>
                </div>
              </div>

              <div class="replica-dialog-overlay" *ngIf="replicaDialog() === 'insertImage'">
                <div class="replica-dialog-box">
                  <div class="rd-header">Insert Picture</div>
                  <div class="rd-body">
                    <div class="rd-pic-preview">🖼 pizza.png</div>
                    <div class="rd-actions">
                      <span class="rd-btn primary" data-demo-target="insert-image-ok-btn">Insert</span>
                    </div>
                  </div>
                </div>
              </div>

              <div class="replica-dialog-overlay" *ngIf="replicaDialog() === 'insertTable'">
                <div class="replica-dialog-box">
                  <div class="rd-header">Insert Table</div>
                  <div class="rd-body">
                    <div class="rd-table-grid">2 x 2 Table</div>
                    <div class="rd-actions">
                      <span class="rd-btn primary" data-demo-target="insert-table-ok-btn">Insert</span>
                    </div>
                  </div>
                </div>
              </div>

              <!-- Animated Cursor Glide Element -->
              <div
                class="animated-cursor"
                [style.left.px]="cursorPixelX()"
                [style.top.px]="cursorPixelY()"
                [class.clicking]="isClicking()"
              >
                <svg viewBox="0 0 24 24" width="22" height="22">
                  <path fill="#202124" stroke="#ffffff" stroke-width="1.5" d="M3 2l18 10-9 2-4 8z" />
                </svg>
                <span class="click-pulse" *ngIf="isClicking()"></span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .demo-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(16, 24, 40, 0.65);
      backdrop-filter: blur(4px);
      z-index: 25000;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
      animation: fadeIn 0.15s ease;
    }
    @keyframes fadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    .demo-dialog {
      border: none;
      border-radius: 12px;
      padding: 20px 24px;
      max-width: 860px;
      width: 100%;
      box-shadow: 0 24px 80px rgba(0, 0, 0, 0.4);
      background: #ffffff;
      color: #201f1e;
      font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
      max-height: 94vh;
      overflow-y: auto;
    }
    .dialog-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 12px;
      border-bottom: 1px solid #e1dfdd;
      padding-bottom: 8px;
    }
    h2 {
      margin: 0 0 4px;
      font-size: 18px;
      color: #201f1e;
      font-weight: 600;
    }
    .header-instruction {
      margin: 0;
      font-size: 13px;
      color: #605e5c;
    }
    .close-button {
      background: none;
      border: none;
      font-size: 26px;
      line-height: 1;
      cursor: pointer;
      color: #605e5c;
      padding: 4px;
      border-radius: 4px;
    }
    .close-button:hover {
      color: #a80000;
      background: #f3f2f1;
    }
    .demo-body {
      display: grid;
      grid-template-columns: 240px 1fr;
      gap: 16px;
      margin-bottom: 8px;
      min-height: 380px;
    }
    .demo-steps-col {
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      border-right: 1px solid #edebe9;
      padding-right: 14px;
    }
    .demo-steps {
      list-style: none;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .demo-steps li {
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 12.5px;
      color: #605e5c;
      padding: 7px 10px;
      border-radius: 6px;
      transition: all 0.2s ease;
    }
    .demo-steps li.active {
      background: #eff6fc;
      color: #185abd;
      font-weight: 600;
      box-shadow: 0 1px 3px rgba(24, 90, 189, 0.15);
    }
    .demo-steps li.done {
      color: #107c41;
      font-weight: 500;
    }
    .step-num {
      width: 22px;
      height: 22px;
      border-radius: 50%;
      background: #f3f2f1;
      color: #605e5c;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 11px;
      font-weight: 700;
      flex-shrink: 0;
    }
    li.active .step-num {
      background: #185abd;
      color: white;
    }
    li.done .step-num {
      background: #107c41;
      color: white;
    }
    .replay-btn {
      padding: 7px 12px;
      border-radius: 4px;
      border: 1px solid #c8c6c4;
      background: #f3f2f1;
      color: #201f1e;
      cursor: pointer;
      font-size: 12.5px;
      font-weight: 600;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      margin-top: 14px;
      transition: background 0.15s;
    }
    .replay-btn:hover {
      background: #e1dfdd;
    }
    .demo-replica-area {
      position: relative;
      display: flex;
      flex-direction: column;
    }
    .demo-toast {
      position: absolute;
      top: 12px;
      left: 50%;
      transform: translateX(-50%);
      background: #182230;
      color: #ffffff;
      padding: 8px 18px;
      border-radius: 20px;
      font-size: 13px;
      font-weight: 600;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.35);
      z-index: 5000;
      animation: toastDrop 0.25s ease-out;
      pointer-events: none;
      white-space: nowrap;
    }
    @keyframes toastDrop {
      from { transform: translate(-50%, -10px); opacity: 0; }
      to { transform: translate(-50%, 0); opacity: 1; }
    }
    .replica-window {
      border: 1px solid #d2d0ce;
      border-radius: 6px;
      overflow: hidden;
      background: #ffffff;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
      position: relative;
      min-height: 360px;
      display: flex;
      flex-direction: column;
    }
    /* Word Replica Styles */
    .replica-word-titlebar {
      background: #185abd;
      color: white;
      padding: 4px 10px;
      display: flex;
      align-items: center;
      font-size: 11px;
    }
    .rw-titlebar-left {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .rw-app-icon {
      display: inline-flex;
      align-items: center;
    }
    .replica-ribbon-tabs {
      background: #ffffff;
      border-bottom: 1px solid #e1dfdd;
      display: flex;
      padding: 0 8px;
      gap: 12px;
      font-size: 11.5px;
    }
    .rw-tab {
      padding: 4px 4px 3px;
      color: #323130;
      cursor: default;
    }
    .rw-tab.active {
      color: #185abd;
      font-weight: 700;
      border-bottom: 2px solid #185abd;
    }
    .rw-tab.disabled {
      color: #a19f9d;
    }
    .replica-ribbon-commands {
      background: #f3f2f1;
      border-bottom: 1px solid #d2d0ce;
      display: flex;
      padding: 4px 6px;
      gap: 8px;
      overflow-x: auto;
    }
    .rw-group {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: space-between;
      border-right: 1px solid #e1dfdd;
      padding: 0 6px;
    }
    .rw-group-row {
      display: flex;
      align-items: center;
      gap: 3px;
    }
    .rw-group-row.vertical {
      flex-direction: column;
      align-items: flex-start;
      gap: 2px;
    }
    .rw-group-label {
      font-size: 9.5px;
      color: #605e5c;
      margin-top: 2px;
    }
    .rw-btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 2px 5px;
      font-size: 11px;
      color: #201f1e;
      border: 1px solid transparent;
      border-radius: 3px;
      background: #ffffff;
      user-select: none;
    }
    .rw-btn.micro {
      min-width: 18px;
      height: 18px;
      padding: 0 3px;
      font-size: 10.5px;
    }
    .rw-btn.micro-text {
      font-size: 10px;
      padding: 1px 4px;
    }
    .rw-btn.highlight {
      background: #c7e0f4;
      border-color: #185abd;
    }
    .rw-select {
      background: white;
      border: 1px solid #c8c6c4;
      font-size: 10px;
      padding: 1px 4px;
      border-radius: 2px;
    }
    .rw-styles-strip {
      display: flex;
      gap: 4px;
    }
    .rw-style-tile {
      background: white;
      border: 1px solid #c8c6c4;
      padding: 2px 5px;
      font-size: 10px;
      border-radius: 2px;
    }
    .rw-style-tile.active {
      border-color: #185abd;
      background: #eff6fc;
      font-weight: 700;
    }
    .replica-doc-workspace {
      flex: 1;
      background: #e8e6e4;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 8px;
      overflow-y: auto;
    }
    .replica-doc-ruler {
      width: 90%;
      height: 12px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      display: flex;
      justify-content: space-around;
      font-size: 8px;
      color: #8a8886;
      margin-bottom: 6px;
    }
    .replica-doc-page {
      width: 90%;
      min-height: 200px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12);
      padding: 16px;
      box-sizing: border-box;
      font-size: 12.5px;
      line-height: 1.5;
      color: #201f1e;
    }
    .replica-doc-content table {
      border-collapse: collapse;
      width: 100%;
      margin: 6px 0;
    }
    .replica-doc-content td {
      border: 1px solid #c8c6c4;
      padding: 4px;
      font-size: 11px;
    }
    .replica-doc-content img {
      max-width: 140px;
      border-radius: 2px;
    }
    .demo-selection-box {
      background: #b4d5fe;
      border-radius: 2px;
    }
    /* Excel & PPT Stubs */
    .replica-excel-bar, .replica-ppt-bar {
      background: #107c41;
      color: white;
      padding: 4px 8px;
      display: flex;
      align-items: center;
      gap: 12px;
      font-size: 11px;
    }
    .replica-ppt-bar {
      background: #c43e1c;
    }
    .re-tools, .rp-tools {
      display: flex;
      gap: 4px;
    }
    .replica-formula-bar {
      display: flex;
      align-items: center;
      gap: 6px;
      background: #f3f2f1;
      border-bottom: 1px solid #d2d0ce;
      padding: 2px 8px;
      font-size: 11px;
    }
    .r-name-box {
      background: white;
      border: 1px solid #c8c6c4;
      padding: 1px 4px;
      font-size: 10px;
    }
    .r-fx-input {
      background: white;
      flex: 1;
      border: 1px solid #c8c6c4;
      padding: 1px 4px;
      font-size: 10px;
    }
    .replica-sheet-wrap {
      flex: 1;
      padding: 8px;
      background: white;
      overflow: auto;
    }
    .replica-sheet-grid {
      border-collapse: collapse;
      width: 100%;
      font-size: 11px;
    }
    .replica-sheet-grid th, .replica-sheet-grid td {
      border: 1px solid #d2d0ce;
      padding: 3px 6px;
    }
    .replica-sheet-grid td.selected {
      outline: 2px solid #107c41;
      background: #e7f5ed;
    }
    .replica-ppt-layout {
      flex: 1;
      display: flex;
      gap: 8px;
      padding: 8px;
      background: #f3f2f1;
    }
    .replica-ppt-thumbs {
      width: 40px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .replica-ppt-thumbs .thumb {
      height: 24px;
      background: white;
      border: 1px solid #c8c6c4;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 9px;
    }
    .replica-ppt-thumbs .thumb.active {
      border-color: #c43e1c;
    }
    .replica-ppt-slide {
      flex: 1;
      background: white;
      border: 1px solid #c8c6c4;
      padding: 12px;
      font-size: 11px;
    }
    /* Replica In-Dialogs */
    .replica-dialog-overlay {
      position: absolute;
      inset: 0;
      background: rgba(0, 0, 0, 0.25);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    }
    .replica-dialog-box {
      background: white;
      border-radius: 4px;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
      width: 260px;
      overflow: hidden;
      font-size: 11.5px;
    }
    .rd-header {
      background: #f3f2f1;
      padding: 6px 10px;
      font-weight: 600;
      border-bottom: 1px solid #e1dfdd;
    }
    .rd-body {
      padding: 10px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .rd-field {
      display: flex;
      flex-direction: column;
      gap: 2px;
    }
    .rd-field input {
      border: 1px solid #c8c6c4;
      padding: 3px 6px;
      border-radius: 2px;
      font-size: 11px;
    }
    .rd-actions {
      display: flex;
      justify-content: flex-end;
      gap: 6px;
      margin-top: 6px;
    }
    .rd-btn {
      padding: 3px 8px;
      border: 1px solid #c8c6c4;
      border-radius: 2px;
      background: #f3f2f1;
      cursor: default;
    }
    .rd-btn.primary {
      background: #185abd;
      color: white;
      border-color: #185abd;
    }
    /* Animated Cursor */
    .animated-cursor {
      position: absolute;
      pointer-events: none;
      z-index: 9999;
      transition: left 0.65s cubic-bezier(0.25, 0.8, 0.25, 1), top 0.65s cubic-bezier(0.25, 0.8, 0.25, 1);
      filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.35));
    }
    .click-pulse {
      position: absolute;
      top: -6px;
      left: -6px;
      width: 24px;
      height: 24px;
      border-radius: 50%;
      border: 2px solid #ea4335;
      animation: pulse 0.45s ease-out forwards;
    }
    @keyframes pulse {
      0% { transform: scale(0.3); opacity: 1; }
      100% { transform: scale(1.6); opacity: 0; }
    }
    @media (max-width: 680px) {
      .demo-dialog {
        padding: 12px;
      }
      .demo-body {
        grid-template-columns: 1fr;
      }
      .demo-steps-col {
        border-right: none;
        border-bottom: 1px solid #edebe9;
        padding-bottom: 10px;
      }
    }
  `],
})
export class ShowMeHowComponent implements OnInit, OnDestroy {
  @Input() appKind: 'word' | 'excel' | 'powerpoint' = 'word';
  @Input() title: string = '';
  @Input() instruction: string = '';
  @Input() goals: TaskGoal[] | any[] = [];
  @Input() startingDocument?: string | null = '';
  @Input() startingSheet?: any = null;
  @Input() startingDeck?: any = null;
  @Input() assetBaseUrl: string = '';

  @Output() readonly close = new EventEmitter<void>();
  @Output() readonly closed = new EventEmitter<void>(); // Backward-compatible output alias

  readonly cursorPixelX = signal<number>(80);
  readonly cursorPixelY = signal<number>(80);
  readonly isClicking = signal<boolean>(false);
  readonly showToast = signal<boolean>(false);
  readonly steps = signal<DemoStep[]>([]);

  readonly activeAction = signal<string>('');
  readonly replicaDialog = signal<string | null>(null);

  // Word replica state
  readonly replicaDocHtml = signal<string>('');
  readonly replicaFont = signal<string>('Calibri');
  readonly replicaFontSize = signal<string>('11');
  readonly replicaStyle = signal<string>('normal');

  // Find & replace typing simulation
  readonly findWhatInput = signal<string>('');
  readonly replaceWithInput = signal<string>('');

  // Excel replica state
  readonly replicaFormulaCell = signal<string>('A1');
  readonly replicaFormulaText = signal<string>('');
  readonly replicaSelectedCell = signal<string>('');
  readonly replicaCells = signal<Record<string, any>>({});

  // PowerPoint replica state
  readonly replicaSlideTitle = signal<string>('My Presentation');
  readonly replicaSlideCount = signal<number>(1);
  readonly replicaTheme = signal<string>('simple-light');
  readonly replicaSlideBulletList = signal<boolean>(false);

  private isCancelled = false;
  private activeTimers: any[] = [];

  constructor(private elRef: ElementRef) {}

  ngOnInit(): void {
    this.setupSteps();
    this.resetReplicaState();
    this.play();
  }

  ngOnDestroy(): void {
    this.cancelAll();
  }

  onClose(): void {
    this.cancelAll();
    this.close.emit();
    this.closed.emit();
  }

  onBackdropClick(e: MouseEvent): void {
    this.onClose();
  }

  replay(): void {
    this.cancelAll();
    this.resetReplicaState();
    this.setupSteps();
    this.play();
  }

  getToastTitle(): string {
    const t = (this.title || '').trim();
    if (!t) return 'complete this task';
    return t.charAt(0).toLowerCase() + t.slice(1);
  }

  private cancelAll(): void {
    this.isCancelled = true;
    for (const t of this.activeTimers) {
      clearTimeout(t);
    }
    this.activeTimers = [];
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => {
      const t = setTimeout(() => resolve(), ms);
      this.activeTimers.push(t);
    });
  }

  private setupSteps(): void {
    const list: DemoStep[] = (this.goals || []).map((g) => ({
      label: this.getStepLabelForGoal(g),
      state: 'pending',
    }));
    if (list.length === 0) {
      list.push({ label: 'Complete the task', state: 'pending' });
    }
    this.steps.set(list);
  }

  private getStepLabelForGoal(g: any): string {
    switch (g?.type) {
      case 'textStyle':
        return 'Select the text, then Apply the format';
      case 'paragraph':
        return 'Format the paragraph';
      case 'list':
        return 'Make the list';
      case 'textContains':
        return 'Edit the text';
      case 'textCount':
        return 'Adjust the copies';
      case 'textOrder':
        return 'Move the text';
      case 'image':
        return 'Insert the picture';
      case 'pageSetup':
        return 'Change the page setup';
      case 'table':
        return 'Insert a table';
      case 'element':
        return 'Add the item';
      case 'action':
        return 'Do the action';
      default:
        return 'Perform the action';
    }
  }

  private resetReplicaState(): void {
    this.isCancelled = false;
    this.showToast.set(false);
    this.activeAction.set('');
    this.replicaDialog.set(null);
    this.findWhatInput.set('');
    this.replaceWithInput.set('');

    // Reset Word
    const rawDoc = this.startingDocument || '<p>I really like pizza.</p>';
    this.replicaDocHtml.set(rawDoc);
    this.replicaFont.set('Calibri');
    this.replicaFontSize.set('11');
    this.replicaStyle.set('normal');

    // Reset Excel
    const sheet = this.startingSheet;
    const cells: Record<string, any> = {
      A1: 'Item', B1: 'Price', C1: 'Category',
      A2: 'Pencil', B2: 0.5, C2: 'School',
      A3: 'Notebook', B3: 2.25, C3: 'School',
      A4: 'Backpack', B4: 24, C4: 'Travel',
      ...(sheet?.cells || {}),
    };
    this.replicaCells.set(cells);
    this.replicaFormulaCell.set('A1');
    this.replicaFormulaText.set('');
    this.replicaSelectedCell.set('');

    // Reset PowerPoint
    const deck = this.startingDeck;
    const slide1Title = deck?.slides?.[0]?.elements?.find((e: any) => e.type === 'title')?.text || 'My Presentation';
    this.replicaSlideTitle.set(slide1Title);
    this.replicaSlideCount.set(deck?.slides?.length || 1);
    this.replicaTheme.set(deck?.theme || 'simple-light');
    this.replicaSlideBulletList.set(false);
  }

  isCellSelected(cell: string): boolean {
    return this.replicaSelectedCell() === cell;
  }

  getReplicaCell(cell: string): string {
    const val = this.replicaCells()[cell];
    return val !== undefined && val !== null ? String(val) : '';
  }

  private async moveTo(targetName: string): Promise<boolean> {
    if (this.isCancelled) return false;
    const container = this.elRef.nativeElement.querySelector('.replica-window') as HTMLElement;
    if (!container) return false;

    let target = container.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;
    if (!target) {
      for (let i = 0; i < 6; i++) {
        await this.delay(80);
        if (this.isCancelled) return false;
        target = container.querySelector(`[data-demo-target="${targetName}"]`) as HTMLElement;
        if (target) break;
      }
    }
    if (!target) return false;

    const contRect = container.getBoundingClientRect();
    const targetRect = target.getBoundingClientRect();

    const x = targetRect.left - contRect.left + targetRect.width / 2;
    const y = targetRect.top - contRect.top + targetRect.height / 2;

    this.cursorPixelX.set(Math.round(x));
    this.cursorPixelY.set(Math.round(y));

    await this.delay(650);
    if (this.isCancelled) return false;

    this.isClicking.set(true);
    await this.delay(200);
    this.isClicking.set(false);
    return true;
  }

  private setStepState(idx: number, state: 'pending' | 'active' | 'done'): void {
    this.steps.update((s) =>
      s.map((step, i) => (i === idx ? { ...step, state } : step))
    );
  }

  async play(): Promise<void> {
    await this.delay(350);
    if (this.isCancelled) return;

    const goalsList = this.goals || [];
    const total = goalsList.length;

    if (total === 0) {
      this.setStepState(0, 'active');
      await this.moveTo('doc-canvas');
      await this.delay(400);
      this.setStepState(0, 'done');
      this.showToast.set(true);
      return;
    }

    for (let i = 0; i < total; i++) {
      if (this.isCancelled) return;
      this.setStepState(i, 'active');
      const g = goalsList[i];

      if (this.appKind === 'word') {
        await this.executeWordGoal(g);
      } else if (this.appKind === 'excel') {
        await this.executeExcelGoal(g);
      } else {
        await this.executePptGoal(g);
      }

      if (this.isCancelled) return;
      this.setStepState(i, 'done');
      await this.delay(300);
    }

    if (this.isCancelled) return;
    this.showToast.set(true);
  }

  private async executeWordGoal(g: any): Promise<void> {
    const type = g?.type;

    if (type === 'textStyle') {
      // 1. Drag across text to select it
      await this.moveTo('doc-canvas');
      const targetText = g.text || 'really';
      let html = this.replicaDocHtml();
      if (html.includes(targetText)) {
        this.replicaDocHtml.set(
          html.replace(targetText, `<span class="demo-selection-box">${targetText}</span>`)
        );
      }
      await this.delay(400);

      // 2. Click toolbar buttons for each listed setting
      if (g.bold) {
        this.activeAction.set('bold');
        await this.moveTo('bold');
        this.replicaDocHtml.update((h) => h.replace(new RegExp(`(<span class="demo-selection-box">)?${targetText}(</span>)?`), `<b>${targetText}</b>`));
      }
      if (g.italic) {
        this.activeAction.set('italic');
        await this.moveTo('italic');
        this.replicaDocHtml.update((h) => h.replace(new RegExp(`(<span class="demo-selection-box">)?${targetText}(</span>)?`), `<i>${targetText}</i>`));
      }
      if (g.underline) {
        this.activeAction.set('underline');
        await this.moveTo('underline');
        this.replicaDocHtml.update((h) => h.replace(new RegExp(`(<span class="demo-selection-box">)?${targetText}(</span>)?`), `<u>${targetText}</u>`));
      }
      if (g.font) {
        this.replicaFont.set(g.font);
        await this.moveTo('font');
        this.replicaDocHtml.update((h) => h.replace(targetText, `<span style="font-family: ${g.font}">${targetText}</span>`));
      }
      if (g.size) {
        this.replicaFontSize.set(String(g.size));
        await this.moveTo('size');
        this.replicaDocHtml.update((h) => h.replace(targetText, `<span style="font-size: ${g.size}pt">${targetText}</span>`));
      }
      if (g.color) {
        await this.moveTo('color');
        this.replicaDocHtml.update((h) => h.replace(targetText, `<span style="color: ${g.color}">${targetText}</span>`));
      }
      if (g.highlight) {
        await this.moveTo('highlight');
        this.replicaDocHtml.update((h) => h.replace(targetText, `<span style="background-color: ${g.highlight}">${targetText}</span>`));
      }
      this.activeAction.set('');
    } else if (type === 'paragraph') {
      await this.moveTo('doc-canvas');
      const targetText = g.text || '';
      if (g.align) {
        this.activeAction.set('align-' + g.align);
        await this.moveTo('align-' + g.align);
        this.replicaDocHtml.update((h) => `<div style="text-align: ${g.align}">${h}</div>`);
      }
      if (g.style) {
        this.replicaStyle.set(g.style);
        const targetBtn = 'style-' + g.style;
        await this.moveTo(targetBtn);
        const tag = g.style === 'normal' ? 'p' : g.style === 'title' ? 'h1' : g.style === 'subtitle' ? 'h2' : g.style;
        this.replicaDocHtml.update((h) => `<${tag} data-paragraph-style="${g.style}">${targetText || h}</${tag}>`);
      }
      if (g.lineSpacing) {
        await this.moveTo('line-spacing');
        this.replicaDocHtml.update((h) => `<div style="line-height: ${g.lineSpacing}">${h}</div>`);
      }
      if (g.indent) {
        await this.moveTo('indent');
        this.replicaDocHtml.update((h) => `<div style="margin-left: ${g.indent * 40}px">${h}</div>`);
      }
      this.activeAction.set('');
    } else if (type === 'list') {
      await this.moveTo('doc-canvas');
      const items = (g.items || ['Item 1', 'Item 2']) as string[];
      await this.moveTo('bullets');
      const listHtml = `<ul>${items.map((it) => `<li>${it}</li>`).join('')}</ul>`;
      this.replicaDocHtml.set(listHtml);
    } else if (type === 'textContains') {
      // Find and replace sequence specified in prompt:
      // 1. Move to Replace in Home > Editing and click it. Dialog opens.
      this.activeAction.set('replace');
      await this.moveTo('replace');
      this.replicaDialog.set('findReplace');
      await this.delay(300);

      // 2. Cursor clicks "Find what" and types wrong word letter by letter
      const notWord = g.notContains?.[0] || 'recieve';
      const rightWord = g.contains?.[0] || 'receive';
      await this.moveTo('find-what-input');
      for (const ch of notWord) {
        if (this.isCancelled) return;
        this.findWhatInput.update((v) => v + ch);
        await this.delay(70);
      }

      // 3. Cursor clicks "Replace with" and types right word letter by letter
      await this.moveTo('replace-with-input');
      for (const ch of rightWord) {
        if (this.isCancelled) return;
        this.replaceWithInput.update((v) => v + ch);
        await this.delay(70);
      }

      // 4. Cursor clicks Replace. Text in document is corrected.
      await this.moveTo('dialog-replace-btn');
      let html = this.replicaDocHtml();
      html = html.split(notWord).join(rightWord);
      this.replicaDocHtml.set(html);
      await this.delay(300);

      // 5. Dialog closes
      this.replicaDialog.set(null);
      this.activeAction.set('');

      // Remaining contains words typed at end if not present
      if (g.contains && g.contains.length > 1) {
        for (let j = 1; j < g.contains.length; j++) {
          const extra = g.contains[j];
          if (!this.replicaDocHtml().includes(extra)) {
            this.replicaDocHtml.update((h) => h + `<p>${extra}</p>`);
          }
        }
      }
    } else if (type === 'textCount') {
      const targetText = g.text || 'Pizza is great.';
      await this.moveTo('doc-canvas');
      // Copy
      this.activeAction.set('copy');
      await this.moveTo('copy');
      await this.delay(200);
      // Move to end and paste
      await this.moveTo('doc-canvas');
      this.activeAction.set('paste');
      await this.moveTo('paste');
      let current = this.replicaDocHtml();
      while (current.split(targetText).length - 1 < (g.count || 2)) {
        current += `<p>${targetText}</p>`;
      }
      this.replicaDocHtml.set(current);
      this.activeAction.set('');
    } else if (type === 'textOrder') {
      await this.moveTo('doc-canvas');
      this.activeAction.set('cut');
      await this.moveTo('cut');
      await this.delay(200);
      await this.moveTo('doc-canvas');
      this.activeAction.set('paste');
      await this.moveTo('paste');
      const seq = (g.sequence || []) as string[];
      this.replicaDocHtml.set(`<p>${seq.join(' ')}</p>`);
      this.activeAction.set('');
    } else if (type === 'image') {
      this.activeAction.set('insert-image');
      await this.moveTo('insert-image');
      this.replicaDialog.set('insertImage');
      await this.delay(300);
      await this.moveTo('insert-image-ok-btn');
      this.replicaDialog.set(null);
      const file = g.file || 'pizza.png';
      const align = g.align || 'center';
      this.replicaDocHtml.update((h) => h + `<div style="text-align: ${align}"><img src="${this.assetBaseUrl || './assets/'}${file}" data-file="${file}" data-align="${align}" /></div>`);
      this.activeAction.set('');
    } else if (type === 'pageSetup') {
      this.activeAction.set('page-setup');
      await this.moveTo('page-setup');
      this.replicaDialog.set('pageSetup');
      await this.delay(300);
      await this.moveTo('pagesetup-ok-btn');
      this.replicaDialog.set(null);
      this.activeAction.set('');
    } else if (type === 'table') {
      this.activeAction.set('insert-table');
      await this.moveTo('insert-table');
      this.replicaDialog.set('insertTable');
      await this.delay(300);
      await this.moveTo('insert-table-ok-btn');
      this.replicaDialog.set(null);
      const rows = g.rows || 2;
      const cols = g.columns || 2;
      let tbl = '<table class="doc-table">';
      for (let r = 0; r < rows; r++) {
        tbl += '<tr>';
        for (let c = 0; c < cols; c++) {
          tbl += `<td>${g.cells?.[r]?.[c] || 'Cell'}</td>`;
        }
        tbl += '</tr>';
      }
      tbl += '</table>';
      this.replicaDocHtml.update((h) => h + tbl);
      this.activeAction.set('');
    } else if (type === 'element') {
      await this.moveTo('doc-canvas');
      const el = g.element;
      if (el === 'pageBreak') {
        this.replicaDocHtml.update((h) => h + '<hr data-element="pageBreak" />');
      } else if (el === 'tableOfContents') {
        this.replicaDocHtml.update((h) => `<div data-element="tableOfContents"><strong>Table of Contents</strong></div>` + h);
      } else if (el === 'footnote') {
        this.replicaDocHtml.update((h) => h + `<sup data-element="footnote">1</sup>`);
      } else {
        this.replicaDocHtml.update((h) => h + `<div data-element="${el}">${g.text || el}</div>`);
      }
    } else if (type === 'action') {
      await this.moveTo('doc-canvas');
      if (g.action === 'copy') await this.moveTo('copy');
      else if (g.action === 'paste') await this.moveTo('paste');
      else if (g.action === 'cut') await this.moveTo('cut');
      else if (g.action === 'find') await this.moveTo('find');
      else await this.moveTo('doc-canvas');
    }
  }

  private async executeExcelGoal(g: any): Promise<void> {
    if (g?.type === 'cellFormat') {
      await this.moveTo('cell-B2');
      this.replicaSelectedCell.set('B2');
      await this.moveTo('excel-currency');
      this.replicaCells.update((c) => ({ ...c, B2: '$0.50', B3: '$2.25', B4: '$24.00' }));
    } else if (g?.type === 'cellFormula') {
      await this.moveTo('cell-B5');
      this.replicaSelectedCell.set('B5');
      this.replicaFormulaCell.set('B5');
      this.replicaFormulaText.set(g.formula || '=SUM(B2:B4)');
      await this.moveTo('excel-sum');
      this.replicaCells.update((c) => ({ ...c, B5: '26.75' }));
    } else {
      await this.moveTo('excel-bold');
    }
  }

  private async executePptGoal(g: any): Promise<void> {
    if (g?.type === 'slideText') {
      await this.moveTo('ppt-title');
      this.activeAction.set('editTitle');
      const text = g.text || 'Mastering Presentation Skills';
      this.replicaSlideTitle.set(text);
      await this.delay(400);
      this.activeAction.set('');
    } else if (g?.type === 'slideCount') {
      await this.moveTo('ppt-newslide');
      this.replicaSlideCount.set(g.count || 2);
    } else if (g?.type === 'theme') {
      await this.moveTo('ppt-theme');
      this.replicaTheme.set(g.theme || 'streamline');
    } else {
      await this.moveTo('ppt-bullets');
      this.replicaSlideBulletList.set(true);
    }
  }
}
