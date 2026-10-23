import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  HostListener,
  OnDestroy,
  OnInit,
  ViewChild,
  computed,
  effect,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  ActivityPayload,
  COLORS,
  ConfettiService,
  FONTS,
  LessonTask,
  PluginBridgeService,
  ShowMeHowComponent,
  TaskCardComponent,
  ThemeService,
  clone,
  sanitizeHtml,
} from 'shared';
import { CustomStyleDefinition, PageSetup, WordComment, WordDocumentService } from './services/word-document.service';
import { WORD_LESSONS } from './word-lessons.data';
import { ChartDialogComponent, ChartDialogResult } from './chart-dialog.component';
import {
  DEFAULT_CHART_DATA,
  WordChartData,
  WordChartType,
  renderChartSvg,
} from './chart-svg.helper';

export interface TaskStartSnapshot {
  documentHtml: string;
  pageSetup: PageSetup;
  comments: WordComment[];
  header: string;
  footer: string;
  numbering: boolean;
  variables: Record<string, string>;
  images: any;
  assetBaseUrl: string;
}

@Component({
  selector: 'app-word-editor',
  standalone: true,
  imports: [CommonModule, FormsModule, TaskCardComponent, ShowMeHowComponent, ChartDialogComponent],
  templateUrl: './word-editor.component.html',
  styleUrls: ['./word-editor.component.css'],
  styles: [`
    .contextual-table-tab {
      color: #0f6cbd !important;
      font-weight: 600 !important;
    }
    .contextual-table-tab.active {
      border-bottom: 2px solid #0f6cbd !important;
    }
    .simulator.word.dark-theme .contextual-table-tab,
    .simulator.word[data-theme='DARK'] .contextual-table-tab {
      color: #4da3ff !important;
    }
    .simulator.word.dark-theme .contextual-table-tab.active,
    .simulator.word[data-theme='DARK'] .contextual-table-tab.active {
      border-bottom-color: #4da3ff !important;
    }

    /* Runtime Chart Element Styles */
    :host ::ng-deep figure.doc-chart {
      display: block;
      margin: 16px auto;
      padding: 8px;
      max-width: 100%;
      box-sizing: border-box;
      user-select: none;
      cursor: pointer;
      background: transparent;
      border: 1px solid transparent;
      position: relative;
      text-align: center;
    }
    :host ::ng-deep figure.doc-chart:hover {
      border: 1px dashed #c8c6c4;
    }
    :host ::ng-deep figure.doc-chart.selected {
      border: 1px solid #0f6cbd;
    }
    :host ::ng-deep figure.doc-chart svg {
      display: block;
      width: 100%;
      height: 100%;
    }
    :host ::ng-deep figure.doc-chart .chart-title-text {
      fill: #242424;
      cursor: text;
    }
    :host ::ng-deep figure.doc-chart .chart-axis-text {
      fill: #616161;
    }
    :host ::ng-deep figure.doc-chart .chart-legend-text {
      fill: #323130;
    }
    :host ::ng-deep figure.doc-chart .grid-line {
      stroke: #e1dfdd;
      stroke-width: 1;
    }

    /* Dark Mode Runtime Chart Styles - NO :is() selectors! */
    :host ::ng-deep .simulator.word.dark-theme figure.doc-chart .chart-title-text,
    :host ::ng-deep .simulator.word[data-theme='DARK'] figure.doc-chart .chart-title-text {
      fill: #ffffff;
    }
    :host ::ng-deep .simulator.word.dark-theme figure.doc-chart .chart-axis-text,
    :host ::ng-deep .simulator.word[data-theme='DARK'] figure.doc-chart .chart-axis-text {
      fill: #c8c6c4;
    }
    :host ::ng-deep .simulator.word.dark-theme figure.doc-chart .chart-legend-text,
    :host ::ng-deep .simulator.word[data-theme='DARK'] figure.doc-chart .chart-legend-text {
      fill: #f3f2f1;
    }
    :host ::ng-deep .simulator.word.dark-theme figure.doc-chart .grid-line,
    :host ::ng-deep .simulator.word[data-theme='DARK'] figure.doc-chart .grid-line {
      stroke: #3b3a39;
    }
    :host ::ng-deep .simulator.word.dark-theme figure.doc-chart:hover,
    :host ::ng-deep .simulator.word[data-theme='DARK'] figure.doc-chart:hover {
      border: 1px dashed #605e5c;
    }

    /* Chart Resize Box and Floating Side Buttons */
    .chart-resize-box {
      border-radius: 2px;
    }
    .chart-resize-box .resize-handle {
      border-radius: 50%;
    }
    .chart-side-buttons {
      position: absolute;
      top: 0;
      right: -36px;
      display: flex;
      flex-direction: column;
      gap: 4px;
      pointer-events: auto;
    }
    .cs-btn {
      width: 26px;
      height: 26px;
      background: #ffffff;
      border: 1px solid #d1d1d1;
      border-radius: 2px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.15);
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #323130;
      padding: 0;
    }
    .cs-btn:hover {
      background: #f3f2f1;
      border-color: #0f6cbd;
    }
    .simulator.word.dark-theme .cs-btn,
    .simulator.word[data-theme='DARK'] .cs-btn {
      background: #252423;
      border-color: #484644;
      color: #f3f2f1;
    }
    .simulator.word.dark-theme .cs-btn:hover,
    .simulator.word[data-theme='DARK'] .cs-btn:hover {
      background: #2d2c2b;
      border-color: #4da3ff;
    }

    /* Inline Title Edit Field */
    .chart-title-edit-overlay {
      position: absolute;
      z-index: 60;
      pointer-events: auto;
    }
    .chart-title-inline-field {
      width: 220px;
      padding: 3px 8px;
      border: 1px dashed #0f6cbd;
      border-radius: 2px;
      background: #ffffff;
      color: #242424;
      font-size: 14px;
      font-weight: 600;
      text-align: center;
      outline: none;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
    }
    .simulator.word.dark-theme .chart-title-inline-field,
    .simulator.word[data-theme='DARK'] .chart-title-inline-field {
      background: #201f1e;
      color: #ffffff;
      border-color: #4da3ff;
    }

    /* Ribbon Chart Styles strip */
    .chart-styles-strip {
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 2px 4px;
      border: 1px solid #e1dfdd;
      border-radius: 2px;
      background: #faf9f8;
    }
    .chart-style-tile {
      width: 38px;
      height: 38px;
      border: 1px solid #d1d1d1;
      border-radius: 2px;
      background: #ffffff;
      display: flex;
      align-items: flex-end;
      justify-content: center;
      padding: 4px;
      box-sizing: border-box;
    }
    .chart-style-tile.active {
      border: 2px solid #0f6cbd;
    }
    .chart-style-tile .cst-bars {
      display: flex;
      align-items: flex-end;
      gap: 2px;
      width: 100%;
      height: 100%;
    }
    .chart-style-tile .cst-bars span {
      flex: 1;
      height: 100%;
      border-radius: 1px 1px 0 0;
    }
    .chart-style-tile .cst-bars span:nth-child(1) { height: 60%; }
    .chart-style-tile .cst-bars span:nth-child(2) { height: 90%; }
    .chart-style-tile .cst-bars span:nth-child(3) { height: 40%; }

    /* Document Table Styles */
    :host ::ng-deep .document-editor table[data-table-style="Plain"] td,
    :host ::ng-deep .document-editor table[data-table-style="Plain"] th {
      border: 1px solid #d2d0ce;
      background-color: transparent;
    }
    :host ::ng-deep .document-editor table[data-table-style="Banded Rows"] tr:nth-child(even) td {
      background-color: #f3f2f1;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid"] td,
    :host ::ng-deep .document-editor table[data-table-style="Grid"] th {
      border: 1px solid #605e5c;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Blue"] tr:first-child td,
    :host ::ng-deep .document-editor table[data-table-style="Grid Blue"] tr:first-child th {
      background-color: #0f6cbd;
      color: #ffffff;
      font-weight: 600;
      border: 1px solid #0f6cbd;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Blue"] tr:not(:first-child):nth-child(even) td {
      background-color: #eff6fc;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Blue"] td {
      border: 1px solid #b4d6fa;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Green"] tr:first-child td,
    :host ::ng-deep .document-editor table[data-table-style="Grid Green"] tr:first-child th {
      background-color: #107c41;
      color: #ffffff;
      font-weight: 600;
      border: 1px solid #107c41;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Green"] tr:not(:first-child):nth-child(even) td {
      background-color: #f1f8f3;
    }
    :host ::ng-deep .document-editor table[data-table-style="Grid Green"] td {
      border: 1px solid #a2d6b3;
    }
    :host ::ng-deep .document-editor table[data-table-style="List Blue"] td,
    :host ::ng-deep .document-editor table[data-table-style="List Blue"] th {
      border-left: none;
      border-right: none;
      border-top: 1px solid #b4d6fa;
      border-bottom: 1px solid #b4d6fa;
    }
    :host ::ng-deep .document-editor table[data-table-style="List Blue"] tr:first-child td,
    :host ::ng-deep .document-editor table[data-table-style="List Blue"] tr:first-child th {
      background-color: #0f6cbd;
      color: #ffffff;
      font-weight: 600;
      border-top: 1px solid #0f6cbd;
      border-bottom: 2px solid #0f6cbd;
    }
    :host ::ng-deep .document-editor table[data-table-style="List Blue"] tr:not(:first-child):nth-child(even) td {
      background-color: #eff6fc;
    }

    /* Dark Mode Table Styles */
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Plain"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Plain"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Plain"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Plain"] th {
      border-color: #484644 !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Banded Rows"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Banded Rows"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Banded Rows"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Banded Rows"] th {
      border-color: #484644 !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Banded Rows"] tr:nth-child(even) td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Banded Rows"] tr:nth-child(even) td {
      background-color: #2b2b2b !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid"] th {
      border-color: #605e5c !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Blue"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Blue"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Blue"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Blue"] th {
      border-color: #2b496e !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Blue"] tr:first-child td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Blue"] tr:first-child th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Blue"] tr:first-child td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Blue"] tr:first-child th {
      background-color: #1a4f8b !important;
      color: #ffffff !important;
      border-color: #1a4f8b !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Blue"] tr:not(:first-child):nth-child(even) td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Blue"] tr:not(:first-child):nth-child(even) td {
      background-color: #1a2d42 !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Green"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Green"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Green"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Green"] th {
      border-color: #1f472e !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Green"] tr:first-child td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Green"] tr:first-child th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Green"] tr:first-child td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Green"] tr:first-child th {
      background-color: #0e5e32 !important;
      color: #ffffff !important;
      border-color: #0e5e32 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="Grid Green"] tr:not(:first-child):nth-child(even) td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="Grid Green"] tr:not(:first-child):nth-child(even) td {
      background-color: #162f20 !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="List Blue"] td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="List Blue"] th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="List Blue"] td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="List Blue"] th {
      border-color: #2b496e !important;
      color: #f3f2f1 !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="List Blue"] tr:first-child td,
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="List Blue"] tr:first-child th,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="List Blue"] tr:first-child td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="List Blue"] tr:first-child th {
      background-color: #1a4f8b !important;
      color: #ffffff !important;
      border-top-color: #1a4f8b !important;
      border-bottom-color: #1a4f8b !important;
    }
    :host ::ng-deep .simulator.word.dark-theme .document-editor table[data-table-style="List Blue"] tr:not(:first-child):nth-child(even) td,
    :host ::ng-deep .simulator.word[data-theme='DARK'] .document-editor table[data-table-style="List Blue"] tr:not(:first-child):nth-child(even) td {
      background-color: #1a2d42 !important;
      color: #f3f2f1 !important;
    }

    /* Ribbon UI Styling for Table Tabs */
    .table-styles-gallery-row {
      display: flex;
      align-items: center;
      gap: 3px;
    }
    .table-style-tile-btn {
      width: 44px;
      height: 44px;
      padding: 2px;
      border: 1px solid #d2d0ce;
      background: #ffffff;
      border-radius: 3px;
      cursor: pointer;
      display: flex;
      flex-direction: column;
      box-sizing: border-box;
      transition: all 0.15s ease;
    }
    .table-style-tile-btn:hover {
      border-color: #0f6cbd;
      box-shadow: 0 0 0 1px #0f6cbd;
    }
    .table-style-tile-btn.active {
      border-color: #0f6cbd;
      background: #eff6fc;
      outline: 2px solid #0f6cbd;
    }
    .mini-table-preview {
      width: 100%;
      height: 100%;
      border-collapse: collapse;
      pointer-events: none;
    }
    .mini-table-preview td {
      padding: 0;
      border: 1px solid #d2d0ce;
      font-size: 0;
    }
    .mini-table-preview.style-plain td {
      border: 1px solid #d2d0ce;
      background: #fff;
    }
    .mini-table-preview.style-banded tr:nth-child(even) td {
      background: #f3f2f1;
    }
    .mini-table-preview.style-grid td {
      border: 1px solid #605e5c;
    }
    .mini-table-preview.style-grid-blue tr:first-child td {
      background: #0f6cbd;
      border-color: #0f6cbd;
    }
    .mini-table-preview.style-grid-blue tr:not(:first-child):nth-child(even) td {
      background: #eff6fc;
    }
    .mini-table-preview.style-grid-blue td {
      border-color: #b4d6fa;
    }
    .mini-table-preview.style-grid-green tr:first-child td {
      background: #107c41;
      border-color: #107c41;
    }
    .mini-table-preview.style-grid-green tr:not(:first-child):nth-child(even) td {
      background: #f1f8f3;
    }
    .mini-table-preview.style-grid-green td {
      border-color: #a2d6b3;
    }
    .mini-table-preview.style-list-blue tr:first-child td {
      background: #0f6cbd;
    }
    .mini-table-preview.style-list-blue tr:not(:first-child):nth-child(even) td {
      background: #eff6fc;
    }
    .mini-table-preview.style-list-blue td {
      border-left: none;
      border-right: none;
      border-color: #b4d6fa;
    }

    .table-styles-more-btn {
      width: 18px;
      height: 44px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      border-radius: 3px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 11px;
      color: #605e5c;
      padding: 0;
    }
    .table-styles-more-btn:hover {
      background: #eff6fc;
      border-color: #0f6cbd;
      color: #0f6cbd;
    }

    /* Flyouts */
    .table-styles-flyout {
      position: fixed;
      min-width: 330px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      border-radius: 6px;
      box-shadow: 0 8px 24px rgba(0,0,0,0.18);
      z-index: 15000;
      padding: 10px 14px;
      box-sizing: border-box;
    }
    .table-styles-section-title {
      font-size: 12px;
      font-weight: 600;
      color: #605e5c;
      margin: 6px 0 4px 0;
    }
    .table-styles-grid {
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }
    .table-style-clear-btn {
      width: 100%;
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 6px 8px;
      background: none;
      border: none;
      border-top: 1px solid #edebe9;
      cursor: pointer;
      font-size: 13px;
      color: #201f1e;
      text-align: left;
    }
    .table-style-clear-btn:hover {
      background: #f3f2f1;
    }

    /* Cell Shading Dropdown */
    .cell-shading-flyout {
      position: fixed;
      width: 240px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      border-radius: 6px;
      box-shadow: 0 8px 24px rgba(0,0,0,0.18);
      z-index: 15000;
      padding: 10px;
      box-sizing: border-box;
    }
    .cell-shading-header {
      font-size: 13px;
      font-weight: 600;
      color: #201f1e;
      margin-bottom: 8px;
    }
    .theme-colors-matrix {
      display: grid;
      grid-template-columns: repeat(10, 1fr);
      gap: 3px;
      margin-bottom: 10px;
    }
    .shading-swatch {
      width: 17px;
      height: 17px;
      border-radius: 50%;
      border: 1px solid rgba(0,0,0,0.15);
      cursor: pointer;
      box-sizing: border-box;
      padding: 0;
      transition: transform 0.1s ease;
    }
    .shading-swatch:hover {
      transform: scale(1.2);
      outline: 1px solid #0f6cbd;
      z-index: 2;
    }
    .standard-colors-row {
      display: grid;
      grid-template-columns: repeat(10, 1fr);
      gap: 3px;
      margin-bottom: 10px;
    }
    .cell-shading-action-btn {
      width: 100%;
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 6px 4px;
      background: none;
      border: none;
      cursor: pointer;
      font-size: 12.5px;
      color: #201f1e;
      text-align: left;
    }
    .cell-shading-action-btn:hover:not(:disabled) {
      background: #f3f2f1;
    }
    .cell-shading-action-btn.disabled, .cell-shading-action-btn:disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }
    .no-color-icon {
      width: 16px;
      height: 16px;
      border: 1px solid #a19f9d;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 11px;
      color: #a80000;
    }

    /* Table Layout Options & Inputs */
    .table-options-grid {
      display: grid;
      grid-template-columns: auto auto;
      gap: 2px 8px;
      font-size: 11px;
      color: #797775;
    }
    .table-checkbox-label {
      display: flex;
      align-items: center;
      gap: 4px;
      cursor: not-allowed;
      user-select: none;
    }
    .table-checkbox-label input {
      margin: 0;
      cursor: not-allowed;
    }
    .align-grid-3x3 {
      display: grid;
      grid-template-columns: repeat(3, 18px);
      gap: 2px;
    }
    .align-cell-btn {
      width: 18px;
      height: 18px;
      padding: 0;
      background: #fff;
      border: 1px solid #d2d0ce;
      border-radius: 2px;
      cursor: not-allowed;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 9px;
      color: #a19f9d;
    }
    .cell-size-col {
      display: flex;
      flex-direction: column;
      gap: 3px;
      font-size: 11px;
    }
    .size-input-row {
      display: flex;
      align-items: center;
      gap: 4px;
      color: #605e5c;
    }
    .size-input-field {
      width: 50px;
      height: 20px;
      font-size: 11px;
      padding: 1px 4px;
      border: 1px solid #d2d0ce;
      background: #f3f2f1;
      border-radius: 2px;
      color: #797775;
      cursor: not-allowed;
    }

    /* Insert Menu Dropdown */
    .table-insert-menu {
      position: fixed;
      min-width: 160px;
      background: #ffffff;
      border: 1px solid #d2d0ce;
      border-radius: 6px;
      box-shadow: 0 8px 24px rgba(0,0,0,0.18);
      z-index: 15000;
      padding: 4px 0;
    }
    .table-insert-item {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 7px 14px;
      border: none;
      background: none;
      width: 100%;
      text-align: left;
      font-size: 13px;
      color: #201f1e;
      cursor: pointer;
      box-sizing: border-box;
    }
    .table-insert-item:hover {
      background: #f3f2f1;
    }

    /* Dark Mode Flyouts */
    .simulator.word.dark-theme .table-style-tile-btn,
    .simulator.word.dark-theme .table-styles-more-btn,
    .simulator.word[data-theme='DARK'] .table-style-tile-btn,
    .simulator.word[data-theme='DARK'] .table-styles-more-btn {
      background: #2b2b2b;
      border-color: #484644;
      color: #f3f2f1;
    }
    .simulator.word.dark-theme .table-style-tile-btn:hover,
    .simulator.word.dark-theme .table-styles-more-btn:hover,
    .simulator.word[data-theme='DARK'] .table-style-tile-btn:hover,
    .simulator.word[data-theme='DARK'] .table-styles-more-btn:hover {
      border-color: #2886de;
      background: #383838;
      color: #ffffff;
    }
    .simulator.word.dark-theme .table-styles-flyout,
    .simulator.word.dark-theme .cell-shading-flyout,
    .simulator.word.dark-theme .table-insert-menu,
    .simulator.word[data-theme='DARK'] .table-styles-flyout,
    .simulator.word[data-theme='DARK'] .cell-shading-flyout,
    .simulator.word[data-theme='DARK'] .table-insert-menu {
      background: #242424;
      border-color: #3b3a39;
      color: #f3f2f1;
      box-shadow: 0 8px 24px rgba(0,0,0,0.6);
    }
    .simulator.word.dark-theme .cell-shading-header,
    .simulator.word.dark-theme .table-styles-section-title,
    .simulator.word[data-theme='DARK'] .cell-shading-header,
    .simulator.word[data-theme='DARK'] .table-styles-section-title {
      color: #f3f2f1;
    }
    .simulator.word.dark-theme .table-style-clear-btn,
    .simulator.word.dark-theme .cell-shading-action-btn,
    .simulator.word.dark-theme .table-insert-item,
    .simulator.word[data-theme='DARK'] .table-style-clear-btn,
    .simulator.word[data-theme='DARK'] .cell-shading-action-btn,
    .simulator.word[data-theme='DARK'] .table-insert-item {
      color: #f3f2f1;
      border-top-color: #3b3a39;
    }
    .simulator.word.dark-theme .table-style-clear-btn:hover,
    .simulator.word.dark-theme .cell-shading-action-btn:hover:not(:disabled),
    .simulator.word.dark-theme .table-insert-item:hover,
    .simulator.word[data-theme='DARK'] .table-style-clear-btn:hover,
    .simulator.word[data-theme='DARK'] .cell-shading-action-btn:hover:not(:disabled),
    .simulator.word[data-theme='DARK'] .table-insert-item:hover {
      background: #333333;
    }
    .simulator.word.dark-theme .align-cell-btn,
    .simulator.word.dark-theme .size-input-field,
    .simulator.word[data-theme='DARK'] .align-cell-btn,
    .simulator.word[data-theme='DARK'] .size-input-field {
      background: #1b1b1b;
      border-color: #484644;
      color: #797775;
    }
  `],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class WordEditorComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('editorElement') editorRef?: ElementRef<HTMLElement>;
  @ViewChild('page2EditorElement') page2EditorRef?: ElementRef<HTMLElement>;
  @ViewChild('headerEl') headerRef?: ElementRef<HTMLElement>;
  @ViewChild('footerEl') footerRef?: ElementRef<HTMLElement>;
  @ViewChild('workLayout') workLayoutRef?: ElementRef<HTMLElement>;
  @ViewChild('wordPaper') paperRef?: ElementRef<HTMLElement>;
  @ViewChild('movableGroup') movableGroupRef?: ElementRef<HTMLElement>;

  readonly hasPageBreak = signal<boolean>(false);
  readonly activeRegion = signal<'body' | 'header' | 'footer'>('body');
  readonly activePageNum = signal<number>(1);
  readonly totalDocPages = computed<number>(() => this.hasPageBreak() ? 2 : 1);

  // Draggable Image Resizing State
  readonly imageBoxLeft = signal<number>(0);
  readonly imageBoxTop = signal<number>(0);
  readonly imageBoxWidth = signal<number>(0);
  readonly imageBoxHeight = signal<number>(0);
  private isResizing = false;
  private resizeHandle: string = '';
  private resizeStartX = 0;
  private resizeStartY = 0;
  private resizeStartWidth = 0;
  private resizeStartHeight = 0;
  private resizeAspectRatio = 1;

  readonly currentActivity = signal<ActivityPayload | null>(null);
  readonly taskIndex = signal<number>(0);
  readonly passedTasks = signal<Set<number>>(new Set());
  readonly isPassed = signal<boolean>(false);
  readonly isCompleted = signal<boolean>(false);
  readonly feedbackMsg = signal<string>('');
  readonly isDemoOpen = signal<boolean>(false);
  readonly currentTaskSnapshot = signal<TaskStartSnapshot | null>(null);
  private readonly taskStartSnapshots = new Map<number, TaskStartSnapshot>();
  readonly wordCount = signal<number>(0);

  get isFrozen(): boolean {
    return this.bridge.isPaused() || this.isCompleted();
  }

  // Movable Group Positioning State (initially bottom-left of work area: left: 14px, bottom: 24px)
  readonly groupX = signal<number | null>(null);
  readonly groupY = signal<number | null>(null);
  readonly hasUserMoved = signal<boolean>(false);
  readonly isDragging = signal<boolean>(false);
  private dragStartPointerX = 0;
  private dragStartPointerY = 0;
  private dragStartGroupX = 0;
  private dragStartGroupY = 0;

  readonly showPageSetup = signal<boolean>(false);
  readonly showFindReplace = signal<boolean>(false);
  readonly showCommentDialog = signal<boolean>(false);
  readonly commentSelectedText = signal<string>('');
  readonly newCommentText = signal<string>('Please review this section.');
  readonly activeRibbonTab = signal<'Home' | 'Insert' | 'Layout' | 'References' | 'Table Design' | 'Table Layout' | 'Chart Design' | 'Format'>('Home');
  readonly showOrientationPicker = signal<boolean>(false);
  readonly openMenuName = signal<string | null>(null);

  readonly selectedStyle = signal<string>('normal');
  readonly selectedFont = signal<string>('Calibri');
  readonly selectedFontSize = signal<number>(11);

  readonly standaloneLessons = WORD_LESSONS;
  readonly standaloneLessonIndex = signal<number>(0);

  private advanceTimer: any = null;

  readonly isAssessment = computed<boolean>(() => {
    return this.currentActivity()?.mode === 'assessment';
  });

  readonly totalTasks = computed<number>(() => {
    return this.currentActivity()?.tasks?.length || 1;
  });

  readonly currentTask = computed<LessonTask | null>(() => {
    const act = this.currentActivity();
    if (!act || !act.tasks) return null;
    return act.tasks[this.taskIndex()] || null;
  });

  readonly isDarkTheme = computed<boolean>(() => {
    return this.bridge.theme() === 'DARK';
  });

  findTextQuery: string = '';
  replaceTextQuery: string = '';
  readonly matchCase = signal<boolean>(false);
  readonly matchCount = signal<number>(0);
  readonly currentMatchIndex = signal<number>(1);

  readonly showImageDialog = signal<boolean>(false);
  readonly selectedImageFile = signal<string>('');
  readonly selectedImageEl = signal<HTMLImageElement | null>(null);
  readonly selectedTableEl = signal<HTMLTableElement | null>(null);
  readonly selectedTableCell = signal<HTMLTableCellElement | null>(null);

  // Charts State
  readonly showInsertChartDialog = signal<boolean>(false);
  readonly chartDialogMode = signal<'insert' | 'change'>('insert');
  readonly chartDialogInitialType = signal<WordChartType>('column');
  readonly chartDialogInitialTitle = signal<string>('Chart Title');
  readonly chartDialogInitialData = signal<WordChartData | null>(null);
  readonly selectedChartEl = signal<HTMLElement | null>(null);
  readonly chartBoxLeft = signal<number>(0);
  readonly chartBoxTop = signal<number>(0);
  readonly chartBoxWidth = signal<number>(480);
  readonly chartBoxHeight = signal<number>(300);
  readonly isEditingChartTitle = signal<boolean>(false);
  readonly chartTitleEditingText = signal<string>('Chart Title');

  readonly showTableStylesGallery = signal<boolean>(false);
  readonly showCellShadingDropdown = signal<boolean>(false);
  readonly showTableInsertMenu = signal<boolean>(false);
  readonly tableStylesPos = signal<{ top: number; left: number }>({ top: 120, left: 300 });
  readonly cellShadingPos = signal<{ top: number; left: number }>({ top: 120, left: 450 });
  readonly tableInsertPos = signal<{ top: number; left: number }>({ top: 120, left: 100 });
  private tableUndoStack: string[] = [];
  private lastTablePostHtml: string | null = null;

  readonly themeColorPalette = [
    [
      { name: 'White', color: '#ffffff', key: 'white' },
      { name: 'Black', color: '#000000', key: 'black' },
      { name: 'Dark Gray', color: '#434343', key: 'darkGray' },
      { name: 'Gray', color: '#999999', key: 'gray' },
      { name: 'Blue', color: '#3d85c6', key: 'blue' },
      { name: 'Orange', color: '#e69138', key: 'orange' },
      { name: 'Yellow', color: '#f1c232', key: 'yellow' },
      { name: 'Green', color: '#6aa84f', key: 'green' },
      { name: 'Cyan', color: '#45818e', key: 'cyan' },
      { name: 'Purple', color: '#674ea7', key: 'purple' },
    ],
    [
      { name: 'Lighter White', color: '#f3f2f1' },
      { name: 'Light Gray', color: '#d2d0ce' },
      { name: 'Light Red', color: '#f4cccc', key: 'lightRed' },
      { name: 'Magenta', color: '#a64d79', key: 'magenta' },
      { name: 'Light Blue', color: '#cfe2f3', key: 'lightBlue' },
      { name: 'Light Orange', color: '#fce5cd', key: 'lightOrange' },
      { name: 'Light Yellow', color: '#fff2cc', key: 'lightYellow' },
      { name: 'Light Green', color: '#d9ead3', key: 'lightGreen' },
      { name: 'Light Cyan', color: '#d0e0e3', key: 'lightCyan' },
      { name: 'Light Purple', color: '#d9d2e9', key: 'lightPurple' },
    ],
    [
      { name: 'Off White', color: '#faf9f8' },
      { name: 'Cool Gray', color: '#a19f9d' },
      { name: 'Light Magenta', color: '#ead1dc', key: 'lightMagenta' },
      { name: 'Soft Dark', color: '#323130' },
      { name: 'Soft Blue', color: '#b4d6fa' },
      { name: 'Soft Orange', color: '#f8d2aa' },
      { name: 'Soft Yellow', color: '#fae69e' },
      { name: 'Soft Green', color: '#b9dfaf' },
      { name: 'Soft Cyan', color: '#a2c8cf' },
      { name: 'Soft Purple', color: '#b7a8d4' },
    ],
    [
      { name: 'Border Gray', color: '#edebe9' },
      { name: 'Medium Gray', color: '#797775' },
      { name: 'Deep Gray', color: '#201f1e' },
      { name: 'Pure Red', color: '#cc0000', key: 'red' },
      { name: 'Medium Blue', color: '#2b73b5' },
      { name: 'Medium Orange', color: '#d17d26' },
      { name: 'Medium Yellow', color: '#dab028' },
      { name: 'Medium Green', color: '#57923d' },
      { name: 'Medium Cyan', color: '#386c77' },
      { name: 'Medium Purple', color: '#543e8c' },
    ],
    [
      { name: 'Neutral Gray', color: '#e1dfdd' },
      { name: 'Slate Gray', color: '#605e5c' },
      { name: 'Charcoal', color: '#1b1a19' },
      { name: 'Deep Red', color: '#a80000' },
      { name: 'Dark Blue', color: '#1e5a91' },
      { name: 'Dark Orange', color: '#b86718' },
      { name: 'Dark Yellow', color: '#ba941c' },
      { name: 'Dark Green', color: '#44782e' },
      { name: 'Dark Cyan', color: '#2c5861' },
      { name: 'Dark Purple', color: '#432f73' },
    ],
    [
      { name: 'Shadow Gray', color: '#c8c6c4' },
      { name: 'Dim Gray', color: '#484644' },
      { name: 'Night', color: '#11100f' },
      { name: 'Maroon', color: '#750b1c' },
      { name: 'Navy Blue', color: '#133e66' },
      { name: 'Rust Orange', color: '#85470d' },
      { name: 'Olive Yellow', color: '#876910' },
      { name: 'Forest Green', color: '#2f551f' },
      { name: 'Deep Cyan', color: '#1e3d43' },
      { name: 'Midnight Purple', color: '#2f2054' },
    ],
  ];

  readonly standardColorsList = [
    { name: 'Dark Red', color: '#c00000', key: 'darkRed' },
    { name: 'Red', color: '#cc0000', key: 'red' },
    { name: 'Orange', color: '#e69138', key: 'orange' },
    { name: 'Yellow', color: '#f1c232', key: 'yellow' },
    { name: 'Light Green', color: '#d9ead3', key: 'lightGreen' },
    { name: 'Green', color: '#6aa84f', key: 'green' },
    { name: 'Light Blue', color: '#cfe2f3', key: 'lightBlue' },
    { name: 'Blue', color: '#3d85c6', key: 'blue' },
    { name: 'Dark Blue', color: '#002060', key: 'darkBlue' },
    { name: 'Purple', color: '#674ea7', key: 'purple' },
  ];
  readonly selectedTocEl = signal<HTMLElement | null>(null);
  private lastTocNavigated = false;
  private taskStartHtml: string | null = null;
  readonly selectedPageNum = signal<boolean>(false);

  readonly showTablePicker = signal<boolean>(false);
  readonly tableHoverRows = signal<number>(2);
  readonly tableHoverCols = signal<number>(2);
  readonly gridRows = [1, 2, 3, 4, 5, 6];
  readonly gridCols = [1, 2, 3, 4, 5, 6, 7, 8];

  readonly showFontColorPicker = signal<boolean>(false);
  readonly showHighlightColorPicker = signal<boolean>(false);
  readonly activeFontColor = signal<string>(COLORS['black']);
  readonly activeHighlightColor = signal<string>(COLORS['yellow']);

  readonly showStylesMenu = signal<boolean>(false);
  readonly showMoreStylesDialog = signal<boolean>(false);
  readonly showCreateStyleDialog = signal<boolean>(false);
  readonly moreStylesSearchQuery = signal<string>('');
  readonly selectedMoreStyle = signal<string>('Normal');
  readonly newStyleName = signal<string>('');
  readonly styleNameError = signal<string>('');
  readonly capturedBold = signal<boolean>(false);
  readonly capturedItalic = signal<boolean>(false);
  readonly capturedUnderline = signal<boolean>(false);
  readonly capturedFont = signal<string>('Calibri');
  readonly capturedSize = signal<number>(11);
  readonly capturedColor = signal<string>('');
  readonly stylesMenuPos = signal<{ top: number; left: number }>({ top: 0, left: 0 });

  readonly filteredMoreStyles = computed<Array<{ name: string; key: string; isCustom?: boolean; customDef?: CustomStyleDefinition }>>(() => {
    const q = this.moreStylesSearchQuery().toLowerCase().trim();
    const list: Array<{ name: string; key: string; isCustom?: boolean; customDef?: CustomStyleDefinition }> = [
      { name: 'Normal', key: 'normal' },
      { name: 'Heading 1', key: 'h1' },
      { name: 'Heading 2', key: 'h2' },
      { name: 'Heading 3', key: 'h3' },
      { name: 'Title', key: 'title' },
      { name: 'Subtitle', key: 'subtitle' },
      ...this.wordDoc.customStyles.map((s) => ({ name: s.name, key: s.name, isCustom: true, customDef: s })),
    ];
    if (!q) return list;
    return list.filter((s) => s.name.toLowerCase().includes(q));
  });

  readonly baseColors: Array<{ key: string; name: string; hex: string }> = [
    { key: 'black', name: 'Black', hex: COLORS['black'] },
    { key: 'darkGray', name: 'Dark Gray', hex: COLORS['darkGray'] },
    { key: 'gray', name: 'Gray', hex: COLORS['gray'] },
    { key: 'white', name: 'White', hex: COLORS['white'] },
    { key: 'red', name: 'Red', hex: COLORS['red'] },
    { key: 'orange', name: 'Orange', hex: COLORS['orange'] },
    { key: 'yellow', name: 'Yellow', hex: COLORS['yellow'] },
    { key: 'green', name: 'Green', hex: COLORS['green'] },
    { key: 'cyan', name: 'Cyan', hex: COLORS['cyan'] },
    { key: 'blue', name: 'Blue', hex: COLORS['blue'] },
    { key: 'purple', name: 'Purple', hex: COLORS['purple'] },
    { key: 'magenta', name: 'Magenta', hex: COLORS['magenta'] },
  ];

  readonly lightColors: Array<{ key: string; name: string; hex: string }> = [
    { key: 'lightRed', name: 'Light Red', hex: COLORS['lightRed'] },
    { key: 'lightOrange', name: 'Light Orange', hex: COLORS['lightOrange'] },
    { key: 'lightYellow', name: 'Light Yellow', hex: COLORS['lightYellow'] },
    { key: 'lightGreen', name: 'Light Green', hex: COLORS['lightGreen'] },
    { key: 'lightCyan', name: 'Light Cyan', hex: COLORS['lightCyan'] },
    { key: 'lightBlue', name: 'Light Blue', hex: COLORS['lightBlue'] },
    { key: 'lightPurple', name: 'Light Purple', hex: COLORS['lightPurple'] },
    { key: 'lightMagenta', name: 'Light Magenta', hex: COLORS['lightMagenta'] },
  ];

  tempSetup: PageSetup = {
    orientation: 'portrait',
    paperSize: 'Letter',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
  };

  readonly fontList = FONTS;
  readonly sizeList = [8, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  readonly colorKeys = Object.keys(COLORS);
  readonly COLORS = COLORS;

  readonly docSetup = signal<PageSetup>({
    orientation: 'portrait',
    paperSize: 'Letter',
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
  });

  readonly paperPadding = computed<string>(() => {
    const m = this.docSetup().margins;
    return `${m.top * 48}px ${m.right * 48}px ${m.bottom * 48}px ${m.left * 48}px`;
  });

  readonly paperSize = computed<'Letter' | 'A4' | 'Legal'>(() => {
    return this.docSetup().paperSize || 'Letter';
  });

  readonly paperOrientation = computed<'portrait' | 'landscape'>(() => {
    return this.docSetup().orientation || 'portrait';
  });

  private bridgeUnsubscribes: Array<() => void> = [];

  constructor(
    public bridge: PluginBridgeService,
    public theme: ThemeService,
    private confetti: ConfettiService,
    public wordDoc: WordDocumentService,
    private cdr: ChangeDetectorRef
  ) {
    this.docSetup.set(clone(this.wordDoc.setup));
    effect(() => {
      if (!this.showFindReplace() || this.findTab() === 'headings') {
        this.clearFindHighlights();
      }
    });
  }

  private lastLoadedActivity: ActivityPayload | null = null;

  ngOnInit(): void {
    const initialAct = this.bridge.activity();
    if (initialAct) {
      this.lastLoadedActivity = initialAct;
      this.loadWordActivity(initialAct);
    }

    this.bridgeUnsubscribes.push(
      this.bridge.onActivity((act) => {
        if (act && act !== this.lastLoadedActivity) {
          this.lastLoadedActivity = act;
          this.loadWordActivity(act);
          this.cdr.markForCheck();
        }
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onForceFinish(() => {
        if (this.advanceTimer) {
          clearTimeout(this.advanceTimer);
          this.advanceTimer = null;
        }
        this.isDemoOpen.set(false);
        const total = this.totalTasks();
        const passedCount = this.passedTasks().size;
        const allDone = passedCount >= total;
        this.bridge.sendComplete(allDone, passedCount, total);
        this.isCompleted.set(true);
        this.cdr.markForCheck();
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onThemeChange(() => {
        this.cdr.detectChanges();
      })
    );

    this.bridgeUnsubscribes.push(
      this.bridge.onPauseChange(() => {
        this.cdr.detectChanges();
      })
    );

    window.addEventListener('scroll', this.onScrollOrResizeCloseMenu, true);
    window.addEventListener('resize', this.onScrollOrResizeCloseMenu);
  }

  private onScrollOrResizeCloseMenu = (): void => {
    if (this.showStylesMenu()) {
      this.showStylesMenu.set(false);
      this.cdr.markForCheck();
    }
  };

  ngAfterViewInit(): void {
    if (this.editorRef) {
      const act = this.currentActivity();
      const task = this.currentTask();
      if (act && task) {
        this.wordDoc.init(task, act, this.editorRef.nativeElement);
        this.docSetup.set(clone(this.wordDoc.setup));
        this.captureTaskStartSnapshot(this.taskIndex(), act, task);
        this.updateTaskStartSnapshot();
      }
    }

    setTimeout(() => {
      this.computeInitialPlacement();
      this.updateHeaderFooterDom();
    }, 50);
  }

  ngOnDestroy(): void {
    window.removeEventListener('scroll', this.onScrollOrResizeCloseMenu, true);
    window.removeEventListener('resize', this.onScrollOrResizeCloseMenu);
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
    }
    this.bridgeUnsubscribes.forEach((unsub) => unsub());
    this.bridgeUnsubscribes = [];
    this.clearFindHighlights();
  }

  toggleMenu(name: string, event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.openMenuName.update(current => (current === name ? null : name));
  }

  closeMenus(): void {
    this.openMenuName.set(null);
    this.showTablePicker.set(false);
    this.showOrientationPicker.set(false);
    this.showTableStylesGallery.set(false);
    this.showCellShadingDropdown.set(false);
    this.showTableInsertMenu.set(false);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event?: MouseEvent): void {
    const target = event?.target as HTMLElement | null;
    if (!target?.closest('.ribbon-tab-menu') && !target?.closest('.ribbon-popover-wrap')) {
      this.closeMenus();
    }
    if (!target?.closest('.color-picker-wrap')) {
      this.showFontColorPicker.set(false);
      this.showHighlightColorPicker.set(false);
    }
    if (!target?.closest('.orientation-dropdown-wrap')) {
      this.showOrientationPicker.set(false);
    }
    if (!target?.closest('.styles-dropdown-menu') && !target?.closest('[data-demo-target="styles-more-btn"]')) {
      this.showStylesMenu.set(false);
    }
    if (!target?.closest('.table-styles-flyout') && !target?.closest('[data-demo-target="table-styles-more"]')) {
      this.showTableStylesGallery.set(false);
    }
    if (!target?.closest('.cell-shading-flyout') && !target?.closest('[data-demo-target="cell-shading"]')) {
      this.showCellShadingDropdown.set(false);
    }
    if (!target?.closest('.table-insert-menu') && !target?.closest('[data-demo-target="table-insert"]')) {
      this.showTableInsertMenu.set(false);
    }
  }

  @HostListener('document:keydown.escape')
  onEscapeKey(): void {
    this.closeMenus();
    this.showFontColorPicker.set(false);
    this.showHighlightColorPicker.set(false);
    this.showCommentDialog.set(false);
    this.showOrientationPicker.set(false);
    this.showTableStylesGallery.set(false);
    this.showCellShadingDropdown.set(false);
    this.showTableInsertMenu.set(false);
    this.showStylesMenu.set(false);
    this.showMoreStylesDialog.set(false);
    this.showCreateStyleDialog.set(false);
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.activeRegion.set('body');
    this.clearFindHighlights();
    if (this.isEditingChartTitle()) {
      this.cancelChartTitleEdit();
    }
    this.showInsertChartDialog.set(false);
    if (this.selectedChartEl()) {
      const allCharts = Array.from(document.querySelectorAll('figure.doc-chart'));
      allCharts.forEach(c => c.classList.remove('selected'));
      this.selectedChartEl.set(null);
      if (this.activeRibbonTab() === 'Chart Design' || this.activeRibbonTab() === 'Format') {
        this.activeRibbonTab.set('Home');
      }
    }
  }

  @HostListener('document:keydown', ['$event'])
  onDocumentKeydown(e: KeyboardEvent): void {
    if (e.key === 'Backspace' || e.key === 'Delete') {
      if (this.selectedChartEl() && !this.isEditingChartTitle()) {
        e.preventDefault();
        this.deleteSelectedChart();
        return;
      }
      if (this.selectedImageEl()) {
        e.preventDefault();
        this.deleteSelectedImage();
        return;
      }
      if (this.selectedPageNum()) {
        e.preventDefault();
        this.deletePageNumbers();
        return;
      }
      if (this.selectedTocEl()) {
        e.preventDefault();
        this.deleteSelectedToc();
        return;
      }
      if (this.selectedTableEl()) {
        const sel = window.getSelection();
        const anchor = sel?.anchorNode;
        const el = anchor?.nodeType === Node.ELEMENT_NODE ? anchor as Element : anchor?.parentElement;
        const inCell = el?.closest('td, th');
        if (!inCell) {
          e.preventDefault();
          this.deleteSelectedTable();
          return;
        }
      }
    }
  }

  private captureTaskStartSnapshot(index: number, act: ActivityPayload, task: LessonTask): void {
    let startingHtml = '';
    if (task.startingDocument !== undefined && task.startingDocument !== null) {
      startingHtml = sanitizeHtml(task.startingDocument, act.assetBaseUrl);
    } else if (this.editorRef) {
      startingHtml = this.editorRef.nativeElement.innerHTML;
    } else {
      startingHtml = this.wordDoc.html || '';
    }

    const snapshot: TaskStartSnapshot = {
      documentHtml: startingHtml,
      pageSetup: clone(this.wordDoc.setup),
      comments: clone(this.wordDoc.comments),
      header: this.wordDoc.header,
      footer: this.wordDoc.footer,
      numbering: this.wordDoc.numbering,
      variables: clone(act.variables || {}),
      images: act.images,
      assetBaseUrl: act.assetBaseUrl || '',
    };
    this.taskStartSnapshots.set(index, snapshot);
    this.currentTaskSnapshot.set(snapshot);
  }

  loadWordActivity(act: ActivityPayload): void {

    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }

    let tasks: LessonTask[] = (act.tasks || []).map((t) => ({
      ...t,
      showMeHow: t.showMeHow === true,
      startingDocument: t.startingDocument !== undefined ? t.startingDocument : act.startingDocument,
      pageSetup: t.pageSetup || act.pageSetup,
    }));
    if (!tasks.length) {
      tasks = [
        {
          title: act.title || 'Apply formatting',
          explanation: 'Practise the tool, then check your work.',
          instructions: [act.instruction || 'Complete the lesson task.'],
          goals: act.goals || [],
          showMeHow: act.showMeHow === true,
          startingDocument: act.startingDocument,
          pageSetup: act.pageSetup,
        },
      ];
    }

    const normalizedAct: ActivityPayload = {
      ...act,
      tasks,
    };

    this.currentActivity.set(normalizedAct);
    this.taskIndex.set(0);
    this.passedTasks.set(new Set());
    this.isPassed.set(false);
    this.isCompleted.set(false);
    this.feedbackMsg.set('');
    this.isDemoOpen.set(false);
    this.showPageSetup.set(false);
    this.showFindReplace.set(false);
    this.clearFindHighlights();
    this.showCommentDialog.set(false);
    this.commentSelectedText.set('');
    this.newCommentText.set('Please review this section.');
    this.showOrientationPicker.set(false);
    this.wordDoc.customStyles = [];
    this.showStylesMenu.set(false);
    this.showMoreStylesDialog.set(false);
    this.showCreateStyleDialog.set(false);
    this.activeRibbonTab.set('Home');
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.findTextQuery = '';
    this.replaceTextQuery = '';
    this.taskStartSnapshots.clear();

    const firstTask = tasks[0];
    const initialDoc = firstTask.startingDocument !== undefined ? firstTask.startingDocument : (act.startingDocument || '');
    if (initialDoc !== undefined && initialDoc !== null) {
      this.wordDoc.html = initialDoc;
    }
    this.updateWordCount();

    const effectiveSetup = firstTask.pageSetup || act.pageSetup;
    this.wordDoc.setup = {
      orientation: 'portrait',
      paperSize: 'Letter',
      margins: { top: 1, bottom: 1, left: 1, right: 1 },
      ...clone(effectiveSetup || {}),
    };
    this.wordDoc.setup.margins = {
      top: 1,
      bottom: 1,
      left: 1,
      right: 1,
      ...effectiveSetup?.margins,
    };
    this.docSetup.set(clone(this.wordDoc.setup));

    if (this.editorRef) {
      this.setupDocumentForTask(firstTask, normalizedAct);
    }
    this.captureTaskStartSnapshot(0, normalizedAct, firstTask);
  }

  loadTaskAtIndex(index: number): void {
    if (!this.editorRef) return;
    const act = this.currentActivity();
    if (!act || !act.tasks) return;
    const task = act.tasks[index];
    if (!task) return;

    this.taskIndex.set(index);
    this.isPassed.set(false);
    this.feedbackMsg.set('');
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);

    this.setupDocumentForTask(task, act);
    this.captureTaskStartSnapshot(index, act, task);
    this.clearFindHighlights();
  }

  getCurrentDocumentHtml(): string {
    const p1 = this.editorRef?.nativeElement.innerHTML || '';
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      return p1 + '<hr data-element="pageBreak" style="display:none">' + this.page2EditorRef.nativeElement.innerHTML;
    }
    return p1;
  }

  private updateTaskStartSnapshot(): void {
    this.taskStartHtml = this.getCurrentDocumentHtml();
    if (this.wordDoc.rich) {
      this.wordDoc.rich.taskStartHtml = this.taskStartHtml;
      this.wordDoc.rich.getDocumentHtml = () => this.getCurrentDocumentHtml();
      (this.wordDoc.rich as any).customUndo = () => this.undoTableAction();
    }
  }

  private setupDocumentForTask(task: LessonTask, act: ActivityPayload): void {
    if (!this.editorRef) return;
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTableCell.set(null);
    this.selectedChartEl.set(null);
    this.isEditingChartTitle.set(false);
    this.showInsertChartDialog.set(false);
    if (this.activeRibbonTab() === 'Chart Design' || this.activeRibbonTab() === 'Format') {
      this.activeRibbonTab.set('Home');
    }
    this.tableUndoStack = [];
    this.lastTablePostHtml = null;
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.activeRegion.set('body');
    this.activePageNum.set(1);

    const rawHtml = task.startingDocument !== undefined && task.startingDocument !== null
      ? task.startingDocument
      : (act.startingDocument || '');

    // Check if document contains a page break element
    if (rawHtml && (rawHtml.includes('data-element="pageBreak"') || rawHtml.includes('pageBreak'))) {
      const parts = rawHtml.split(/<hr[^>]*data-element=["']pageBreak["'][^>]*>/i);
      if (parts.length > 1) {
        this.wordDoc.init({ ...task, startingDocument: parts[0] + '<hr data-element="pageBreak" style="display:none">' }, act, this.editorRef.nativeElement);
        this.docSetup.set(clone(this.wordDoc.setup));
        this.hasPageBreak.set(true);
        this.wordDoc.hasPageBreak = true;
        setTimeout(() => {
          if (this.page2EditorRef?.nativeElement) {
            this.page2EditorRef.nativeElement.innerHTML = parts.slice(1).join('') || '<p><br></p>';
            this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
          }
          this.updateHeaderFooterDom();
          this.syncDocumentContent();
          this.updateTaskStartSnapshot();
        });
        return;
      }
    }

    this.hasPageBreak.set(false);
    this.wordDoc.hasPageBreak = false;
    this.wordDoc.page2Element = null;
    if (task.startingDocument !== null) {
      this.wordDoc.init(task, act, this.editorRef.nativeElement);
      this.docSetup.set(clone(this.wordDoc.setup));
    } else {
      this.wordDoc.actions = [];
      const effectivePageSetup = task.pageSetup || act.pageSetup;
      if (effectivePageSetup) {
        this.wordDoc.setup = {
          orientation: 'portrait',
          paperSize: 'Letter',
          margins: { top: 1, bottom: 1, left: 1, right: 1 },
          ...clone(effectivePageSetup),
        };
        this.wordDoc.setup.margins = {
          top: 1,
          bottom: 1,
          left: 1,
          right: 1,
          ...effectivePageSetup.margins,
        };
      }
      this.docSetup.set(clone(this.wordDoc.setup));
    }
    this.updateHeaderFooterDom();
    this.updateWordCount();
    this.updateTaskStartSnapshot();
  }

  loadFallbackActivity(index: number = 0): void {
    const clamped = Math.max(0, Math.min(index, this.standaloneLessons.length - 1));
    this.standaloneLessonIndex.set(clamped);
    this.bridge.resetCompleted();
    const sample = this.standaloneLessons[clamped];
    if (sample) {
      this.loadWordActivity(sample);
    }
  }

  selectStandaloneLesson(idx: number): void {
    this.loadFallbackActivity(idx);
  }

  onEditorInput(): void {
    const editor = this.editorRef?.nativeElement;
    if (editor) {
      // 1. Clean up empty variable tokens
      editor.querySelectorAll<HTMLElement>('span[data-element="variable"]').forEach((v) => {
        if (!v.textContent?.trim()) {
          v.remove();
        }
      });

      // 2. Clean up orphaned footnote section if all footnote references are gone
      const fnRefs = editor.querySelectorAll('sup.footnote-ref');
      const fnSection = editor.querySelector('.footnote-section');
      if (fnSection && fnRefs.length === 0) {
        fnSection.remove();
      }

      // 3. Clean up completely empty list shells
      editor.querySelectorAll<HTMLElement>('ul, ol').forEach((list) => {
        if (!list.textContent?.trim() && !list.contains(document.activeElement)) {
          list.remove();
        }
      });

      // 4. Prune empty page break if removed from page 1 DOM
      if (this.hasPageBreak() && !editor.querySelector('[data-element="pageBreak"]')) {
        this.mergePage2IntoPage1();
      }
    }
    this.syncDocumentContent();
  }

  onPage2Input(): void {
    this.syncDocumentContent();
  }

  onEditorKeydown(e: KeyboardEvent): void {
    if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedChartEl() && !this.isEditingChartTitle()) {
      e.preventDefault();
      this.deleteSelectedChart();
      return;
    }

    if (e.key === 'Delete' && this.hasPageBreak()) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
        const range = sel.getRangeAt(0);
        const testRange = document.createRange();
        testRange.setStart(range.endContainer, range.endOffset);
        testRange.setEndAfter(this.editorRef.nativeElement.lastChild || this.editorRef.nativeElement);
        if (testRange.toString().trim() === '') {
          e.preventDefault();
          this.mergePage2IntoPage1();
          return;
        }
      }
    }

    if (e.key === 'Backspace') {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
        const anchor = sel.anchorNode;
        const el = anchor?.nodeType === Node.ELEMENT_NODE ? anchor as Element : anchor?.parentElement;

        // If inside an empty table
        const cell = el?.closest('td, th');
        const table = cell?.closest('table');
        if (cell && table) {
          const isEmpty = Array.from(table.querySelectorAll('td, th')).every(c => !c.textContent?.trim());
          if (isEmpty) {
            e.preventDefault();
            const p = document.createElement('p');
            p.innerHTML = '<br>';
            table.parentNode?.insertBefore(p, table);
            table.remove();
            this.selectedTableEl.set(null);
            this.onEditorInput();
            const r = document.createRange();
            r.setStart(p, 0);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
            return;
          }
        }

        // If in an empty li
        const li = el?.closest('li');
        if (li && !li.textContent?.trim()) {
          const list = li.closest('ul, ol');
          if (list) {
            e.preventDefault();
            li.remove();
            if (list.querySelectorAll('li').length === 0) {
              const p = document.createElement('p');
              p.innerHTML = '<br>';
              list.parentNode?.insertBefore(p, list);
              list.remove();
              const r = document.createRange();
              r.setStart(p, 0);
              r.collapse(true);
              sel.removeAllRanges();
              sel.addRange(r);
            }
            this.onEditorInput();
            return;
          }
        }
      }
    }
  }

  mergePage2IntoPage1(): void {
    const editor = this.editorRef?.nativeElement;
    const p2 = this.page2EditorRef?.nativeElement;
    if (editor) {
      const marker = editor.querySelector('[data-element="pageBreak"]');
      if (marker) marker.remove();
      if (p2) {
        const text = p2.innerText.trim();
        if (text) {
          editor.innerHTML += p2.innerHTML;
        }
      }
      editor.focus();
    }
    this.hasPageBreak.set(false);
    this.wordDoc.hasPageBreak = false;
    this.wordDoc.page2Element = null;
    this.activePageNum.set(1);
    this.syncDocumentContent();
  }

  onPage2Keydown(e: KeyboardEvent): void {
    if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedChartEl() && !this.isEditingChartTitle()) {
      e.preventDefault();
      this.deleteSelectedChart();
      return;
    }

    if (e.key === 'Backspace' && this.page2EditorRef) {
      const p2 = this.page2EditorRef.nativeElement;
      const text = p2.innerText.trim();
      const sel = window.getSelection();
      const isAtStart = sel && sel.rangeCount > 0 && sel.anchorOffset === 0;

      if (!text || isAtStart) {
        e.preventDefault();
        this.mergePage2IntoPage1();
      }
    }
  }

  syncDocumentContent(): void {
    const p1Html = this.editorRef?.nativeElement.innerHTML || '';
    const p2Html = this.hasPageBreak() && this.page2EditorRef?.nativeElement ? this.page2EditorRef.nativeElement.innerHTML : '';
    this.wordDoc.html = p1Html + (this.hasPageBreak() ? '<hr data-element="pageBreak" style="display:none">' + p2Html : '');
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
    } else {
      this.wordDoc.page2Element = null;
    }
    this.updateWordCount();
  }

  updateHeaderFooterDom(): void {
    if (this.headerRef?.nativeElement && document.activeElement !== this.headerRef.nativeElement) {
      this.headerRef.nativeElement.textContent = this.wordDoc.header || '';
    }
    if (this.footerRef?.nativeElement && document.activeElement !== this.footerRef.nativeElement) {
      this.footerRef.nativeElement.textContent = this.wordDoc.footer || '';
    }
  }

  onHeaderInput(e: Event): void {
    const text = (e.target as HTMLElement).textContent || '';
    this.wordDoc.header = text.trim() === '' ? '' : text;
  }

  onFooterInput(e: Event): void {
    const text = (e.target as HTMLElement).textContent || '';
    this.wordDoc.footer = text.trim() === '' ? '' : text;
  }

  removeHeader(): void {
    this.wordDoc.header = '';
    if (this.headerRef?.nativeElement) {
      this.headerRef.nativeElement.textContent = '';
    }
    this.wordDoc.logAction('header', '');
    this.closeHeaderFooterEdit();
  }

  removeFooter(): void {
    this.wordDoc.footer = '';
    if (this.footerRef?.nativeElement) {
      this.footerRef.nativeElement.textContent = '';
    }
    this.wordDoc.logAction('footer', '');
    this.closeHeaderFooterEdit();
  }

  onHeaderBlur(): void {
    // Keep active region or update if needed
  }

  onFooterBlur(): void {
    // Keep active region or update if needed
  }

  openHeaderEdit(e?: MouseEvent): void {
    if (e) e.stopPropagation();
    this.activeRegion.set('header');
    this.cdr.detectChanges();
    setTimeout(() => {
      this.headerRef?.nativeElement?.focus();
    });
  }

  openFooterEdit(e?: MouseEvent): void {
    if (e) e.stopPropagation();
    this.activeRegion.set('footer');
    this.cdr.detectChanges();
    setTimeout(() => {
      this.footerRef?.nativeElement?.focus();
    });
  }

  closeHeaderFooterEdit(): void {
    this.activeRegion.set('body');
  }

  private updateWordCount(): void {
    let text = this.editorRef?.nativeElement.innerText || '';
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      text += ' ' + this.page2EditorRef.nativeElement.innerText;
    }
    if (!text && this.wordDoc.html) {
      const temp = document.createElement('div');
      temp.innerHTML = this.wordDoc.html;
      text = temp.innerText || temp.textContent || '';
    }
    const words = text.trim().split(/\s+/).filter(Boolean).length;
    this.wordCount.set(words);
  }

  formatCmd(cmd: string, val?: string): void {
    if (cmd === 'undo') {
      const currentHtml = this.getCurrentDocumentHtml();
      if (this.taskStartHtml !== null && currentHtml === this.taskStartHtml) {
        return;
      }
      if (this.undoTableAction()) {
        return;
      }
    }
    this.wordDoc.rich?.command(cmd, val);
    if (cmd === 'undo' || cmd === 'redo') {
      this.wordDoc.logAction(cmd);
    }
    this.onEditorInput();
  }

  cutText(): void {
    this.wordDoc.rich?.copy(true);
    this.onEditorInput();
  }

  copyText(): void {
    this.wordDoc.rich?.copy(false);
  }

  pasteText(): void {
    this.wordDoc.rich?.paste();
    this.onEditorInput();
  }

  applyFont(font: string): void {
    this.wordDoc.rich?.apply({ fontFamily: font });
    this.onEditorInput();
  }

  applyFontSize(size: string): void {
    this.wordDoc.rich?.apply({ fontSize: size + 'pt' });
    this.onEditorInput();
  }

  applyColor(colorNameOrHex: string, colorKey?: string): void {
    const c = COLORS[colorNameOrHex] || colorNameOrHex;
    this.wordDoc.rich?.applyColor(c);
    this.activeFontColor.set(c);
    this.wordDoc.logAction('fontColor', colorKey || colorNameOrHex);
    this.onEditorInput();
  }

  applyHighlight(colorNameOrHex: string, colorKey?: string): void {
    const c = COLORS[colorNameOrHex] || colorNameOrHex;
    this.wordDoc.rich?.applyHighlight(c);
    this.activeHighlightColor.set(c && c !== 'transparent' ? c : COLORS['yellow']);
    this.wordDoc.logAction('highlight', colorKey || colorNameOrHex);
    this.onEditorInput();
  }

  toggleFontColorPicker(event?: MouseEvent): void {
    if (event) {
      event.stopPropagation();
    }
    this.showHighlightColorPicker.set(false);
    this.showFontColorPicker.update((v) => !v);
    this.wordDoc.rich?.capture();
  }

  toggleHighlightPicker(event?: MouseEvent): void {
    if (event) {
      event.stopPropagation();
    }
    this.showFontColorPicker.set(false);
    this.showHighlightColorPicker.update((v) => !v);
    this.wordDoc.rich?.capture();
  }

  selectFontColor(hex: string, colorKey?: string): void {
    this.applyColor(hex, colorKey);
    this.showFontColorPicker.set(false);
  }

  selectHighlightColor(hex: string, colorKey?: string): void {
    this.applyHighlight(hex, colorKey);
    this.showHighlightColorPicker.set(false);
  }

  applyAutomaticFontColor(): void {
    this.activeFontColor.set(COLORS['black']);
    this.showFontColorPicker.set(false);
    this.wordDoc.rich?.applyColor(COLORS['black']);
    this.wordDoc.logAction('fontColor', 'automatic');
    this.onEditorInput();
  }

  applyNoHighlight(): void {
    this.activeHighlightColor.set(COLORS['yellow']);
    this.showHighlightColorPicker.set(false);
    this.wordDoc.rich?.applyHighlight('transparent');
    this.wordDoc.logAction('highlight', 'none');
    this.onEditorInput();
  }

  applyLineSpacing(val: string): void {
    const b = this.wordDoc.rich?.block();
    if (b) {
      b.style.lineHeight = val;
      b.dataset['lineSpacing'] = val;
      this.wordDoc.logAction('lineSpacing', val);
      this.onEditorInput();
    }
  }

  matchColorKey(colorStr: string, el?: HTMLElement | null): string {
    if (!colorStr) return '';
    let curr: HTMLElement | null = el || null;
    while (curr && curr.tagName !== 'DIV' && curr.tagName !== 'BODY') {
      const inlineColor = curr.style?.color?.toLowerCase().trim();
      if (inlineColor) {
        if (COLORS[inlineColor]) return inlineColor;
        for (const [k, v] of Object.entries(COLORS)) {
          if (inlineColor === v.toLowerCase() || inlineColor === k.toLowerCase()) {
            return k;
          }
        }
      }
      curr = curr.parentElement;
    }

    const probe = document.createElement('span');
    document.body.appendChild(probe);
    let matched = '';
    for (const [k, v] of Object.entries(COLORS)) {
      probe.style.color = v;
      if (window.getComputedStyle(probe).color === colorStr) {
        matched = k;
        break;
      }
      probe.style.color = k;
      if (window.getComputedStyle(probe).color === colorStr) {
        matched = k;
        break;
      }
    }
    probe.remove();
    return matched;
  }

  toggleStylesMenu(event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    if (this.showStylesMenu()) {
      this.showStylesMenu.set(false);
      return;
    }
    this.wordDoc.rich?.capture();
    this.updateStylesMenuPos(event.currentTarget as HTMLElement);
    this.showStylesMenu.set(true);
  }

  updateStylesMenuPos(btnEl?: HTMLElement | null): void {
    const btn = btnEl || (document.querySelector('[data-demo-target="styles-more-btn"]') as HTMLElement | null);
    if (!btn) return;
    const rect = btn.getBoundingClientRect();
    const menuWidth = 240;
    let left = rect.right - menuWidth;
    let top = rect.bottom + 2;

    const vw = window.innerWidth;
    const vh = window.innerHeight;

    if (left + menuWidth > vw - 8) {
      left = vw - menuWidth - 8;
    }
    if (left < 8) {
      left = 8;
    }

    const estimatedMenuHeight = 320;
    if (top + estimatedMenuHeight > vh - 8 && rect.top - estimatedMenuHeight > 8) {
      top = rect.top - estimatedMenuHeight - 2;
    }

    this.stylesMenuPos.set({ top, left });
  }

  openMoreStylesDialog(): void {
    this.showStylesMenu.set(false);
    this.moreStylesSearchQuery.set('');
    this.selectedMoreStyle.set('Normal');
    this.showMoreStylesDialog.set(true);
  }

  confirmMoreStyles(): void {
    const sel = this.selectedMoreStyle();
    if (!sel) return;
    const custom = this.wordDoc.customStyles.find((cs) => cs.name === sel);
    if (custom) {
      this.applyCustomStyle(custom);
    } else {
      const mapping: Record<string, string> = {
        Normal: 'normal',
        'Heading 1': 'h1',
        'Heading 2': 'h2',
        'Heading 3': 'h3',
        Title: 'title',
        Subtitle: 'subtitle',
      };
      const key = mapping[sel] || sel.toLowerCase();
      this.applyParagraphStyle(key);
    }
    this.showMoreStylesDialog.set(false);
  }

  openCreateStyleDialog(): void {
    this.showStylesMenu.set(false);

    let targetEl: HTMLElement | null = null;
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      let node: Node | null = range.startContainer;
      if (node.nodeType === Node.TEXT_NODE) {
        targetEl = node.parentElement;
      } else if (node instanceof HTMLElement) {
        targetEl = node;
      }
    }
    if (!targetEl && this.wordDoc.rich?.saved) {
      const node = this.wordDoc.rich.saved.startContainer;
      targetEl = node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as HTMLElement);
    }
    if (!targetEl && this.editorRef?.nativeElement) {
      targetEl = (this.editorRef.nativeElement.querySelector('p, h1, h2, h3, div') as HTMLElement | null) || this.editorRef.nativeElement;
    }

    if (targetEl) {
      const cs = window.getComputedStyle(targetEl);
      this.capturedBold.set(Number(cs.fontWeight) >= 600 || cs.fontWeight === 'bold');
      this.capturedItalic.set(cs.fontStyle === 'italic');
      this.capturedUnderline.set(cs.textDecorationLine.includes('underline'));
      const font = cs.fontFamily.replaceAll('"', '').split(',')[0].trim();
      this.capturedFont.set(font || 'Calibri');
      const sizePt = Math.round(parseFloat(cs.fontSize) * 0.75);
      this.capturedSize.set(sizePt > 0 ? sizePt : 11);
      const colorKey = this.matchColorKey(cs.color, targetEl);
      this.capturedColor.set(colorKey);
    } else {
      this.capturedBold.set(false);
      this.capturedItalic.set(false);
      this.capturedUnderline.set(false);
      this.capturedFont.set('Calibri');
      this.capturedSize.set(11);
      this.capturedColor.set('');
    }

    this.newStyleName.set('');
    this.styleNameError.set('');
    this.showCreateStyleDialog.set(true);
  }

  confirmCreateStyle(): void {
    const name = this.newStyleName().trim();
    if (!name) {
      this.styleNameError.set('Please enter a style name.');
      return;
    }
    const builtInNames = ['normal', 'no spacing', 'heading 1', 'heading 2', 'heading 3', 'title', 'subtitle', 'h1', 'h2', 'h3'];
    if (
      builtInNames.includes(name.toLowerCase()) ||
      this.wordDoc.customStyles.some((s) => s.name.toLowerCase() === name.toLowerCase())
    ) {
      this.styleNameError.set('A style with this name already exists.');
      return;
    }

    const newStyle: CustomStyleDefinition = {
      name,
      bold: this.capturedBold(),
      italic: this.capturedItalic(),
      underline: this.capturedUnderline(),
      font: this.capturedFont(),
      size: this.capturedSize(),
      color: this.capturedColor(),
    };

    this.wordDoc.customStyles.push(newStyle);
    this.applyCustomStyle(newStyle);
    this.showCreateStyleDialog.set(false);
  }

  applyCustomStyle(style: CustomStyleDefinition): void {
    this.selectedStyle.set(style.name);
    this.wordDoc.rich?.command('formatBlock', 'p');
    let b = this.wordDoc.rich?.block();
    if (!b && this.editorRef?.nativeElement) {
      const sel = window.getSelection();
      if (sel && sel.anchorNode) {
        b = sel.anchorNode.parentElement?.closest('p, h1, h2, h3, li, div') as HTMLElement | null;
      }
      if (!b) {
        b = this.editorRef.nativeElement.querySelector('p, h1, h2, h3, div') as HTMLElement | null;
      }
    }
    if (b) {
      if (['H1', 'H2', 'H3'].includes(b.tagName)) {
        const p = document.createElement('p');
        p.innerHTML = b.innerHTML;
        b.replaceWith(p);
        b = p;
      }
      b.classList.remove('doc-title', 'doc-subtitle');
      b.dataset['paragraphStyle'] = style.name;
      b.style.fontWeight = style.bold ? 'bold' : 'normal';
      b.style.fontStyle = style.italic ? 'italic' : 'normal';
      b.style.textDecoration = style.underline ? 'underline' : 'none';
      if (style.font) b.style.fontFamily = style.font;
      if (style.size) b.style.fontSize = style.size + 'pt';
      if (style.color) b.style.color = COLORS[style.color] || style.color;

      b.querySelectorAll('*').forEach((el) => {
        const h = el as HTMLElement;
        if (style.bold !== undefined && h.style.fontWeight) h.style.fontWeight = '';
        if (style.italic !== undefined && h.style.fontStyle) h.style.fontStyle = '';
        if (style.underline !== undefined && h.style.textDecoration) h.style.textDecoration = '';
        if (style.color && h.style.color) h.style.color = '';
        if (style.font && h.style.fontFamily) h.style.fontFamily = '';
        if (style.size && h.style.fontSize) h.style.fontSize = '';
      });

      this.wordDoc.logAction('style', style.name);
      this.onEditorInput();
    }
  }

  applyParagraphStyle(style: string): void {
    this.selectedStyle.set(style);
    const tag = style === 'normal' ? 'p' : style === 'h1' || style === 'h2' || style === 'h3' ? style : 'p';
    this.wordDoc.rich?.command('formatBlock', tag);
    let b = this.wordDoc.rich?.block();
    if (!b && this.editorRef?.nativeElement) {
      const sel = window.getSelection();
      if (sel && sel.anchorNode) {
        b = sel.anchorNode.parentElement?.closest('p, h1, h2, h3, li, div') as HTMLElement | null;
      }
      if (!b) {
        b = this.editorRef.nativeElement.querySelector('p, h1, h2, h3, div') as HTMLElement | null;
      }
    }
    if (b) {
      b.dataset['paragraphStyle'] = style;
      b.classList.remove('doc-title', 'doc-subtitle');
      if (style === 'title') {
        b.classList.add('doc-title');
      } else if (style === 'subtitle') {
        b.classList.add('doc-subtitle');
      }
      b.style.fontWeight = '';
      b.style.fontStyle = '';
      b.style.textDecoration = '';
      b.style.color = '';
      b.style.fontFamily = '';
      b.style.fontSize = '';
      this.wordDoc.logAction('paragraphStyle', style);
      this.onEditorInput();
    }
  }

  openImageDialog(): void {
    const imgs = this.currentActivity()?.images;
    this.selectedImageFile.set(imgs?.[0]?.file || 'photo.png');
    this.showImageDialog.set(true);
    this.closeMenus();
  }

  insertChosenImage(): void {
    const file = this.selectedImageFile() || 'photo.png';
    const base = this.currentActivity()?.assetBaseUrl || './assets/';
    const img = document.createElement('img');
    img.src = base ? (base.endsWith('/') ? base + file : base + '/' + file) : file;
    img.dataset['file'] = file;
    img.dataset['wrap'] = 'inline';
    img.dataset['align'] = 'left';
    img.style.maxWidth = '300px';
    img.style.cursor = 'pointer';

    const sel = window.getSelection();
    const ed1 = this.editorRef?.nativeElement;
    const ed2 = this.page2EditorRef?.nativeElement;
    let inserted = false;

    if (sel && sel.rangeCount > 0 && sel.anchorNode && ed1) {
      const node = sel.anchorNode;
      const targetEditor = ed1.contains(node) ? ed1 : (ed2?.contains(node) ? ed2 : null);
      if (targetEditor) {
        const range = sel.getRangeAt(0);
        range.deleteContents();
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        inserted = true;
      }
    }

    if (!inserted) {
      ed1?.append(img);
    }

    this.wordDoc.logAction('insertImage', file);
    this.showImageDialog.set(false);
    this.onEditorInput();
  }

  insertSampleImage(): void {
    this.openImageDialog();
  }

  onEditorClick(event: MouseEvent): void {
    this.activeRegion.set('body');
    const target = event.target as HTMLElement;
    if (target && target.tagName === 'IMG') {
      const img = target as HTMLImageElement;
      this.selectedImageEl.set(img);
      this.wordDoc.logAction('selectImage', img.dataset['file'] || img.src);
      setTimeout(() => this.updateImageBoxRect());
    } else {
      this.selectedImageEl.set(null);
    }

    const chart = target?.closest('figure.doc-chart') as HTMLElement | null;
    if (chart) {
      const allCharts = Array.from(document.querySelectorAll('figure.doc-chart'));
      allCharts.forEach(c => c.classList.remove('selected'));
      chart.classList.add('selected');
      this.selectedChartEl.set(chart);
      this.activeRibbonTab.set('Chart Design');
      this.wordDoc.logAction('selectChart', chart.getAttribute('data-chart-type') || 'chart');
      setTimeout(() => this.updateChartBoxRect());

      if (target.classList?.contains('chart-title-text') || target.closest('.chart-title-text')) {
        this.startChartTitleEdit();
      }
    } else {
      if (this.selectedChartEl()) {
        const allCharts = Array.from(document.querySelectorAll('figure.doc-chart'));
        allCharts.forEach(c => c.classList.remove('selected'));
        this.selectedChartEl.set(null);
        if (this.isEditingChartTitle()) {
          this.finishChartTitleEdit();
        }
        if (this.activeRibbonTab() === 'Chart Design' || this.activeRibbonTab() === 'Format') {
          this.activeRibbonTab.set('Home');
        }
      }
    }

    const table = target?.closest('table') as HTMLTableElement | null;
    this.selectedTableEl.set(table);
    const cell = target?.closest('td, th') as HTMLTableCellElement | null;
    if (cell) this.selectedTableCell.set(cell);
    if (!table && (this.activeRibbonTab() === 'Table Design' || this.activeRibbonTab() === 'Table Layout')) {
      this.activeRibbonTab.set('Home');
    }

    const toc = target?.closest('.table-of-contents') as HTMLElement | null;
    this.selectedTocEl.set(toc);

    if (!target?.closest('.page-num')) {
      this.selectedPageNum.set(false);
    }
  }

  deleteSelectedImage(): void {
    const img = this.selectedImageEl();
    if (img) {
      img.remove();
      this.selectedImageEl.set(null);
      this.wordDoc.logAction('deleteImage');
      this.onEditorInput();
    }
  }

  deleteSelectedTable(): void {
    const table = this.selectedTableEl();
    if (table) {
      table.remove();
      this.selectedTableEl.set(null);
      this.wordDoc.logAction('deleteTable');
      this.onEditorInput();
    }
  }

  deleteSelectedToc(): void {
    const toc = this.selectedTocEl();
    if (toc) {
      toc.remove();
      this.selectedTocEl.set(null);
      this.wordDoc.logAction('deleteTableOfContents');
      this.onEditorInput();
    }
  }

  selectPageNumber(e: MouseEvent): void {
    e.stopPropagation();
    this.selectedPageNum.set(true);
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
  }

  deletePageNumbers(e?: Event): void {
    if (e) e.preventDefault();
    this.wordDoc.numbering = false;
    this.selectedPageNum.set(false);
    this.wordDoc.logAction('pageNumbers');
    this.onEditorInput();
  }

  onPage1Click(event: MouseEvent): void {
    this.activePageNum.set(1);
    this.onEditorClick(event);
    const target = event.target as HTMLElement | null;
    if (this.lastTocNavigated || target?.closest('.table-of-contents a, .doc-toc a')) {
      this.selectedTocEl.set(null);
    }
  }

  onPage2Click(event: MouseEvent): void {
    this.activePageNum.set(2);
    this.onEditorClick(event);
    const target = event.target as HTMLElement | null;
    if (this.lastTocNavigated || target?.closest('.table-of-contents a, .doc-toc a')) {
      this.selectedTocEl.set(null);
    }
  }

  onPaperClick(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    if (!target.closest('.page-header') && !target.closest('.page-footer')) {
      this.activeRegion.set('body');
    }
    if (!target.closest('figure.doc-chart') && !target.closest('.chart-resize-box') && !target.closest('.chart-title-edit-overlay')) {
      if (this.selectedChartEl()) {
        const allCharts = Array.from(document.querySelectorAll('figure.doc-chart'));
        allCharts.forEach(c => c.classList.remove('selected'));
        this.selectedChartEl.set(null);
        if (this.isEditingChartTitle()) {
          this.finishChartTitleEdit();
        }
        if (this.activeRibbonTab() === 'Chart Design' || this.activeRibbonTab() === 'Format') {
          this.activeRibbonTab.set('Home');
        }
      }
    }
  }

  alignSelectedImage(align: 'left' | 'center' | 'right'): void {
    const img = this.selectedImageEl();
    if (img) {
      img.dataset['align'] = align;
      if (align === 'center') {
        img.style.display = 'block';
        img.style.margin = '8px auto';
      } else if (align === 'right') {
        img.style.display = 'block';
        img.style.margin = '8px 0 8px auto';
      } else {
        img.style.display = 'inline-block';
        img.style.margin = '0';
      }
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  wrapSelectedImage(wrap: 'inline' | 'wrap' | 'break'): void {
    const img = this.selectedImageEl();
    if (img) {
      img.dataset['wrap'] = wrap;
      if (wrap === 'wrap') {
        img.style.float = 'left';
        img.style.margin = '0 12px 12px 0';
      } else if (wrap === 'break') {
        img.style.display = 'block';
        img.style.clear = 'both';
      } else {
        img.style.float = 'none';
        img.style.display = 'inline-block';
      }
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  resizeSelectedImage(width: number): void {
    const img = this.selectedImageEl();
    if (img) {
      img.width = width;
      img.style.width = width + 'px';
      img.style.height = 'auto';
      this.onEditorInput();
      setTimeout(() => this.updateImageBoxRect());
    }
  }

  updateImageBoxRect(): void {
    const img = this.selectedImageEl();
    if (!img) return;
    const canvas = document.querySelector('.word-canvas') as HTMLElement | null;
    if (!canvas) return;

    const cRect = canvas.getBoundingClientRect();
    const iRect = img.getBoundingClientRect();

    this.imageBoxLeft.set(Math.round(iRect.left - cRect.left + canvas.scrollLeft));
    this.imageBoxTop.set(Math.round(iRect.top - cRect.top + canvas.scrollTop));
    this.imageBoxWidth.set(Math.round(iRect.width));
    this.imageBoxHeight.set(Math.round(iRect.height));
  }

  onResizeStart(e: MouseEvent, handle: string): void {
    e.preventDefault();
    e.stopPropagation();
    const img = this.selectedImageEl();
    if (!img) return;

    this.isResizing = true;
    this.resizeHandle = handle;
    this.resizeStartX = e.clientX;
    this.resizeStartY = e.clientY;
    this.resizeStartWidth = img.clientWidth || img.width || 300;
    this.resizeStartHeight = img.clientHeight || img.height || 200;
    this.resizeAspectRatio = this.resizeStartWidth / (this.resizeStartHeight || 1);

    const onMouseMove = (moveEvt: MouseEvent) => {
      if (!this.isResizing || !this.selectedImageEl()) return;
      const dx = moveEvt.clientX - this.resizeStartX;
      const dy = moveEvt.clientY - this.resizeStartY;
      let newWidth = this.resizeStartWidth;

      if (handle === 'se') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'sw') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 'ne') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'nw') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 'e') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth + dx));
      } else if (handle === 'w') {
        newWidth = Math.max(50, Math.min(700, this.resizeStartWidth - dx));
      } else if (handle === 's' || handle === 'n') {
        const factor = handle === 's' ? 1 : -1;
        const newHeight = Math.max(40, this.resizeStartHeight + dy * factor);
        newWidth = Math.round(newHeight * this.resizeAspectRatio);
      }

      const targetImg = this.selectedImageEl()!;
      targetImg.width = Math.round(newWidth);
      targetImg.style.width = `${Math.round(newWidth)}px`;
      targetImg.style.height = 'auto';
      this.updateImageBoxRect();
    };

    const onMouseUp = () => {
      this.isResizing = false;
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      this.onEditorInput();
      this.updateImageBoxRect();
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  insertGridTable(rows: number, cols: number): void {
    const table = document.createElement('table');
    table.className = 'doc-table';
    for (let i = 0; i < rows; i++) {
      const r = table.insertRow();
      for (let j = 0; j < cols; j++) {
        const c = r.insertCell();
        c.innerHTML = '<br>';
      }
    }

    const sel = window.getSelection();
    const ed1 = this.editorRef?.nativeElement;
    const ed2 = this.page2EditorRef?.nativeElement;
    let targetBlock: HTMLElement | null = null;
    let targetEditor = ed1;

    if (sel && sel.rangeCount > 0 && sel.anchorNode && ed1) {
      const node = sel.anchorNode;
      const inEd = ed1.contains(node) ? ed1 : (ed2?.contains(node) ? ed2 : null);
      if (inEd) {
        targetEditor = inEd;
        let curr: HTMLElement | null = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
        while (curr && curr !== targetEditor) {
          if (curr.parentElement === targetEditor) {
            targetBlock = curr;
            break;
          }
          curr = curr.parentElement;
        }
      }
    }

    if (targetBlock) {
      targetBlock.after(table);
    } else {
      targetEditor?.append(table);
    }

    this.wordDoc.logAction('insertTable');
    this.showTablePicker.set(false);
    this.closeMenus();
    this.onEditorInput();
  }

  insertSampleTable(): void {
    this.insertGridTable(2, 2);
  }

  saveTableUndo(): void {
    const html = this.getCurrentDocumentHtml();
    this.tableUndoStack.push(html);
  }

  undoTableAction(): boolean {
    if (this.tableUndoStack.length > 0) {
      const currentHtml = this.getCurrentDocumentHtml();
      if (this.lastTablePostHtml !== null && currentHtml !== this.lastTablePostHtml) {
        this.tableUndoStack = [];
        this.lastTablePostHtml = null;
        return false;
      }
      const prev = this.tableUndoStack.pop()!;
      this.restoreTableUndo(prev);
      this.lastTablePostHtml = this.tableUndoStack.length > 0 ? prev : null;
      this.wordDoc.logAction('undo');
      return true;
    }
    return false;
  }

  restoreTableUndo(html: string): void {
    if (this.editorRef) {
      this.editorRef.nativeElement.innerHTML = html;
    }
    this.onEditorInput();
    const s = window.getSelection();
    if (s && s.anchorNode) {
      const el = s.anchorNode.nodeType === Node.ELEMENT_NODE ? (s.anchorNode as HTMLElement) : s.anchorNode.parentElement;
      const tbl = el?.closest('table') as HTMLTableElement | null;
      this.selectedTableEl.set(tbl);
      const cell = el?.closest('td, th') as HTMLTableCellElement | null;
      if (cell) this.selectedTableCell.set(cell);
      if (!tbl && (this.activeRibbonTab() === 'Table Design' || this.activeRibbonTab() === 'Table Layout')) {
        this.activeRibbonTab.set('Home');
      }
    } else {
      const tbl = this.editorRef?.nativeElement.querySelector('table') as HTMLTableElement | null;
      this.selectedTableEl.set(tbl);
    }
  }

  getActiveTableCell(): HTMLTableCellElement | null {
    const sel = window.getSelection();
    if (sel && sel.anchorNode) {
      const el = sel.anchorNode.nodeType === Node.ELEMENT_NODE
        ? (sel.anchorNode as HTMLElement)
        : sel.anchorNode.parentElement;
      const cell = el?.closest('td, th') as HTMLTableCellElement | null;
      if (cell) return cell;
    }
    return this.selectedTableCell() || null;
  }

  toggleTableStylesGallery(event?: MouseEvent): void {
    if (event) event.stopPropagation();
    const btn = (event?.currentTarget as HTMLElement) || (document.querySelector('[data-demo-target="table-styles-more"]') as HTMLElement | null);
    if (btn) {
      const rect = btn.getBoundingClientRect();
      const left = Math.max(10, Math.min(window.innerWidth - 350, rect.right - 330));
      this.tableStylesPos.set({ top: rect.bottom + 4, left });
    }
    this.showTableStylesGallery.update(v => !v);
    this.showCellShadingDropdown.set(false);
    this.showTableInsertMenu.set(false);
  }

  toggleCellShading(event?: MouseEvent): void {
    if (event) event.stopPropagation();
    const btn = (event?.currentTarget as HTMLElement) || (document.querySelector('[data-demo-target="cell-shading"]') as HTMLElement | null);
    if (btn) {
      const rect = btn.getBoundingClientRect();
      const left = Math.max(10, Math.min(window.innerWidth - 260, rect.left));
      this.cellShadingPos.set({ top: rect.bottom + 4, left });
    }
    this.showCellShadingDropdown.update(v => !v);
    this.showTableStylesGallery.set(false);
    this.showTableInsertMenu.set(false);
  }

  toggleTableInsertMenu(event?: MouseEvent): void {
    if (event) event.stopPropagation();
    const btn = (event?.currentTarget as HTMLElement) || (document.querySelector('[data-demo-target="table-insert"]') as HTMLElement | null);
    if (btn) {
      const rect = btn.getBoundingClientRect();
      const left = Math.max(10, Math.min(window.innerWidth - 180, rect.left));
      this.tableInsertPos.set({ top: rect.bottom + 4, left });
    }
    this.showTableInsertMenu.update(v => !v);
    this.showTableStylesGallery.set(false);
    this.showCellShadingDropdown.set(false);
  }

  applyTableStyle(styleName: string): void {
    const table = this.selectedTableEl();
    if (!table) return;
    this.saveTableUndo();
    table.setAttribute('data-table-style', styleName);
    this.wordDoc.logAction('tableStyle', styleName);
    this.showTableStylesGallery.set(false);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  clearTableStyle(): void {
    const table = this.selectedTableEl();
    if (!table) return;
    this.saveTableUndo();
    table.removeAttribute('data-table-style');
    this.wordDoc.logAction('tableStyleClear');
    this.showTableStylesGallery.set(false);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  applyCellShading(colorKey: string | null, fallbackHex?: string): void {
    const table = this.selectedTableEl();
    if (!table) return;
    this.saveTableUndo();
    let colorHex = '';
    if (colorKey) {
      if (COLORS[colorKey]) {
        colorHex = COLORS[colorKey];
      } else if (fallbackHex) {
        colorHex = fallbackHex;
      } else {
        const stdSwatch = this.standardColorsList.find(c => c.key === colorKey);
        if (stdSwatch) {
          colorHex = stdSwatch.color;
        } else {
          let foundInPalette: string | undefined;
          for (const row of this.themeColorPalette) {
            const sw = row.find(c => c.key === colorKey);
            if (sw) {
              foundInPalette = sw.color;
              break;
            }
          }
          if (foundInPalette) {
            colorHex = foundInPalette;
          } else {
            colorHex = colorKey;
          }
        }
      }
    }

    const sel = window.getSelection();
    const cellsToShade: HTMLTableCellElement[] = [];

    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      const allCells = Array.from(table.querySelectorAll('td, th')) as HTMLTableCellElement[];
      for (const cell of allCells) {
        if (range.intersectsNode(cell)) {
          cellsToShade.push(cell);
        }
      }
    }

    if (cellsToShade.length === 0) {
      const activeCell = this.getActiveTableCell();
      if (activeCell) {
        cellsToShade.push(activeCell);
      } else if (table.rows.length > 0 && table.rows[0].cells.length > 0) {
        cellsToShade.push(table.rows[0].cells[0]);
      }
    }

    for (const cell of cellsToShade) {
      if (colorHex) {
        cell.style.backgroundColor = colorHex;
      } else {
        cell.style.removeProperty('background-color');
      }
    }

    if (colorKey) {
      this.wordDoc.logAction('cellShading', colorKey);
    } else {
      this.wordDoc.logAction('cellShadingNoColor');
    }

    this.showCellShadingDropdown.set(false);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  insertTableRow(position: 'above' | 'below'): void {
    const table = this.selectedTableEl();
    if (!table || table.rows.length === 0) return;
    this.saveTableUndo();

    const activeCell = this.getActiveTableCell();
    let rowIndex = table.rows.length - 1;
    if (activeCell && activeCell.parentElement) {
      const tr = activeCell.parentElement as HTMLTableRowElement;
      if (tr.rowIndex >= 0) rowIndex = tr.rowIndex;
    }

    const colCount = table.rows[rowIndex]?.cells.length || table.rows[0].cells.length || 1;
    const targetIndex = position === 'above' ? rowIndex : rowIndex + 1;
    const newRow = table.insertRow(targetIndex);
    for (let j = 0; j < colCount; j++) {
      const newCell = newRow.insertCell();
      newCell.innerHTML = '<br>';
    }

    this.wordDoc.logAction(position === 'above' ? 'insertTableRowAbove' : 'insertTableRowBelow');
    this.showTableInsertMenu.set(false);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  insertTableColumn(position: 'left' | 'right'): void {
    const table = this.selectedTableEl();
    if (!table || table.rows.length === 0) return;
    this.saveTableUndo();

    const activeCell = this.getActiveTableCell();
    let colIndex = table.rows[0].cells.length - 1;
    if (activeCell && activeCell.parentElement) {
      const tr = activeCell.parentElement as HTMLTableRowElement;
      const idx = Array.from(tr.cells).indexOf(activeCell);
      if (idx >= 0) colIndex = idx;
    }

    const targetIndex = position === 'left' ? colIndex : colIndex + 1;
    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      const newCell = row.insertCell(Math.min(targetIndex, row.cells.length));
      newCell.innerHTML = '<br>';
    }

    this.wordDoc.logAction(position === 'left' ? 'insertTableColumnLeft' : 'insertTableColumnRight');
    this.showTableInsertMenu.set(false);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  insertPageBreak(): void {
    if (!this.editorRef) return;
    const editor = this.editorRef.nativeElement;

    let moveHtml = '';
    const sel = window.getSelection();

    if (sel && sel.rangeCount > 0 && editor.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      const afterRange = document.createRange();
      afterRange.setStart(range.endContainer, range.endOffset);
      afterRange.setEndAfter(editor.lastChild || editor);
      const frag = afterRange.extractContents();
      const div = document.createElement('div');
      div.appendChild(frag);
      moveHtml = div.innerHTML.trim();
    }

    const hr = document.createElement('hr');
    hr.dataset['element'] = 'pageBreak';
    hr.className = 'page-break-marker';
    hr.style.display = 'none';
    editor.appendChild(hr);

    this.hasPageBreak.set(true);
    this.wordDoc.hasPageBreak = true;
    this.wordDoc.logAction('pageBreak');
    this.closeMenus();

    setTimeout(() => {
      if (this.page2EditorRef?.nativeElement) {
        this.page2EditorRef.nativeElement.innerHTML = moveHtml || '<p><br></p>';
        this.wordDoc.page2Element = this.page2EditorRef.nativeElement;
        this.page2EditorRef.nativeElement.focus();
      }
      this.syncDocumentContent();
      this.activePageNum.set(2);
    });
  }

  toggleHeader(): void {
    const text = this.wordDoc.header || 'Document Header';
    this.wordDoc.header = text;
    this.activeRegion.set('header');
    this.cdr.detectChanges();
    this.wordDoc.logAction('header', text);
    this.closeMenus();
    setTimeout(() => {
      this.updateHeaderFooterDom();
      this.headerRef?.nativeElement?.focus();
    });
  }

  toggleFooter(): void {
    const text = this.wordDoc.footer || 'Document Footer';
    this.wordDoc.footer = text;
    this.activeRegion.set('footer');
    this.cdr.detectChanges();
    this.wordDoc.logAction('footer', text);
    this.closeMenus();
    setTimeout(() => {
      this.updateHeaderFooterDom();
      this.footerRef?.nativeElement?.focus();
    });
  }

  togglePageNumbers(): void {
    this.wordDoc.numbering = !this.wordDoc.numbering;
    this.wordDoc.logAction('pageNumbers');
    this.closeMenus();
  }

  insertTableOfContents(): void {
    if (!this.editorRef) return;
    const headings = Array.from(this.editorRef.nativeElement.querySelectorAll('h1, h2, h3, [data-paragraph-style="title"], [data-paragraph-style="h1"], [data-paragraph-style="h2"], [data-paragraph-style="h3"]'));
    const toc = document.createElement('div');
    toc.className = 'table-of-contents doc-toc';
    toc.dataset['element'] = 'tableOfContents';
    toc.tabIndex = 0;

    let listHtml = '';
    if (headings.length > 0) {
      headings.forEach((h: any, i: number) => {
        const id = h.id || `heading-anchor-${i}`;
        h.id = id;
        const level = h.tagName === 'H1' ? 1 : h.tagName === 'H2' ? 2 : 3;
        listHtml += `<li class="toc-item toc-level-${level}"><a href="#${id}">${h.textContent || 'Section ' + (i + 1)}</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li>`;
      });
    } else {
      listHtml = `<li class="toc-item"><a href="#doc-top">Document Overview</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li><li class="toc-item"><a href="#doc-details">Details</a><span class="toc-dots">.....</span><span class="toc-page">1</span></li>`;
    }

    toc.innerHTML = `<div class="toc-header-bar"><strong>Table of Contents</strong><button type="button" class="toc-remove-btn" title="Remove Table of Contents" aria-label="Remove Table of Contents">✕</button></div><ul class="toc-list">${listHtml}</ul>`;
    const btn = toc.querySelector('.toc-remove-btn');
    if (btn) {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        toc.remove();
        this.selectedTocEl.set(null);
        this.wordDoc.logAction('deleteTableOfContents');
        this.onEditorInput();
      });
    }

    const sel = window.getSelection();
    const ed1 = this.editorRef.nativeElement;
    const ed2 = this.page2EditorRef?.nativeElement;
    let targetBlock: HTMLElement | null = null;
    let targetEditor = ed1;

    if (sel && sel.rangeCount > 0 && sel.anchorNode) {
      const node = sel.anchorNode;
      const inEd = ed1.contains(node) ? ed1 : (ed2?.contains(node) ? ed2 : null);
      if (inEd) {
        targetEditor = inEd;
        let curr: HTMLElement | null = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
        while (curr && curr !== targetEditor) {
          if (curr.parentElement === targetEditor) {
            targetBlock = curr;
            break;
          }
          curr = curr.parentElement;
        }
      }
    }

    if (targetBlock) {
      targetBlock.after(toc);
    } else {
      targetEditor.prepend(toc);
    }

    this.wordDoc.logAction('tableOfContents');
    this.closeMenus();
    this.onEditorInput();
  }

  handleTocClick(event: MouseEvent): boolean {
    this.lastTocNavigated = false;
    if (this.bridge.isPaused() || this.isDemoOpen()) {
      return false;
    }

    const target = event.target as HTMLElement | null;
    const link = target?.closest('.table-of-contents a, .doc-toc a') as HTMLAnchorElement | null;
    if (!link) {
      return false;
    }

    event.preventDefault();

    const href = link.getAttribute('href') || '';
    const targetId = href.startsWith('#') ? href.slice(1) : href;
    if (!targetId) {
      return false;
    }

    const editor1 = this.editorRef?.nativeElement;
    const editor2 = this.page2EditorRef?.nativeElement;
    let targetHeading: HTMLElement | null = null;
    if (editor1) {
      targetHeading = editor1.querySelector('#' + CSS.escape(targetId)) as HTMLElement | null;
    }
    if (!targetHeading && editor2) {
      targetHeading = editor2.querySelector('#' + CSS.escape(targetId)) as HTMLElement | null;
    }
    if (!targetHeading) {
      targetHeading = document.getElementById(targetId);
    }

    if (!targetHeading) {
      return false;
    }

    targetHeading.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    const range = document.createRange();
    range.selectNodeContents(targetHeading);
    range.collapse(true);
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(range);
    }
    const headingEditor = targetHeading.closest('.document-editor') as HTMLElement | null;
    if (headingEditor) {
      headingEditor.focus();
      if (editor2 && headingEditor === editor2) {
        this.activePageNum.set(2);
      } else {
        this.activePageNum.set(1);
      }
    }

    this.wordDoc.logAction('navigateToc');
    this.updateActiveFormatting();
    this.selectedTocEl.set(null);
    this.lastTocNavigated = true;
    return true;
  }

  insertFootnote(): void {
    if (!this.editorRef) return;
    const currentNotes = this.editorRef.nativeElement.querySelectorAll('sup.footnote-ref').length;
    const noteNum = currentNotes + 1;
    const fnRef = document.createElement('sup');
    fnRef.className = 'footnote-ref';
    fnRef.dataset['element'] = 'footnote';
    fnRef.dataset['noteId'] = String(noteNum);
    fnRef.tabIndex = 0;
    fnRef.textContent = `[${noteNum}]`;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && this.editorRef.nativeElement.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(fnRef);
      range.setStartAfter(fnRef);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      this.editorRef.nativeElement.append(fnRef);
    }

    let fnSection = this.editorRef.nativeElement.querySelector<HTMLElement>('.footnote-section');
    if (!fnSection) {
      fnSection = document.createElement('div');
      fnSection.className = 'footnote-section';
      fnSection.innerHTML = '<hr class="footnote-divider"/>';
      this.editorRef.nativeElement.append(fnSection);
    }
    const fnItem = document.createElement('p');
    fnItem.className = 'footnote';
    fnItem.dataset['element'] = 'footnote';
    fnItem.dataset['noteId'] = String(noteNum);
    fnItem.innerHTML = `<sup>${noteNum}</sup> Footnote ${noteNum} details <button type="button" class="footnote-remove-btn" title="Remove footnote" style="background:none;border:none;cursor:pointer;color:#888;font-size:11px;margin-left:6px;">✕</button>`;
    fnSection.append(fnItem);

    const btn = fnItem.querySelector('.footnote-remove-btn');
    if (btn) {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        fnItem.remove();
        fnRef.remove();
        if (fnSection && fnSection.querySelectorAll('.footnote').length === 0) {
          fnSection.remove();
        }
        this.onEditorInput();
      });
    }

    this.wordDoc.logAction('footnote');
    this.closeMenus();
    this.onEditorInput();
  }

  insertComment(commentText?: string, targetText?: string): void {
    const sel = window.getSelection();
    const text = targetText !== undefined ? targetText : (sel ? sel.toString().trim() : '');
    if (commentText !== undefined) {
      this.wordDoc.comments.push({
        text: text || 'Selected text',
        comment: commentText || 'Please review this section.',
        resolved: false,
        replies: []
      });
      this.wordDoc.logAction('comment', text || commentText);
      this.closeMenus();
      this.onEditorInput();
      return;
    }
    this.commentSelectedText.set(text || 'Selected text');
    this.newCommentText.set('Please review this section.');
    this.showCommentDialog.set(true);
    this.closeMenus();
  }

  confirmInsertComment(): void {
    const comment = this.newCommentText().trim() || 'Please review this section.';
    const text = this.commentSelectedText();
    this.wordDoc.comments.push({
      text: text || 'Selected text',
      comment: comment,
      resolved: false,
      replies: []
    });
    this.wordDoc.logAction('comment', text || comment);
    this.showCommentDialog.set(false);
    this.newCommentText.set('Please review this section.');
    this.onEditorInput();
  }

  cancelInsertComment(): void {
    this.showCommentDialog.set(false);
    this.newCommentText.set('Please review this section.');
  }

  toggleResolveComment(index: number): void {
    if (this.wordDoc.comments[index]) {
      this.wordDoc.comments[index].resolved = !this.wordDoc.comments[index].resolved;
      this.wordDoc.logAction('resolveComment');
      this.onEditorInput();
    }
  }

  addCommentReply(index: number, text: string): void {
    const trimmed = (text || '').trim();
    if (!trimmed || !this.wordDoc.comments[index]) return;
    if (!this.wordDoc.comments[index].replies) {
      this.wordDoc.comments[index].replies = [];
    }
    this.wordDoc.comments[index].replies!.push(trimmed);
    this.wordDoc.logAction('replyComment', trimmed);
    this.onEditorInput();
  }

  deleteComment(index: number): void {
    if (this.wordDoc.comments && index >= 0 && index < this.wordDoc.comments.length) {
      this.wordDoc.comments.splice(index, 1);
      this.wordDoc.logAction('deleteComment');
      this.onEditorInput();
    }
  }

  showVariablesPanel = signal<boolean>(false);
  readonly selectedVariableKey = signal<string>('');
  readonly selectedVariableVal = signal<string>('');

  openVariablesPanel(): void {
    const list = this.getVariablesList();
    if (list.length > 0) {
      this.selectedVariableKey.set(list[0].key);
      this.selectedVariableVal.set(list[0].val);
    } else {
      this.selectedVariableKey.set('');
      this.selectedVariableVal.set('');
    }
    this.showVariablesPanel.set(true);
    this.wordDoc.logAction('openVariables');
    this.closeMenus();
  }

  confirmInsertVariable(): void {
    const key = this.selectedVariableKey();
    const val = this.selectedVariableVal();
    if (key) {
      this.insertVariable(key, val);
    }
  }

  getVariablesList(): { key: string; val: string }[] {
    const vars = this.currentActivity()?.variables;
    if (!vars || typeof vars !== 'object') return [];
    return Object.entries(vars).map(([key, val]) => ({ key, val: String(val) }));
  }

  insertVariable(key: string, val: string): void {
    const span = document.createElement('span');
    span.dataset['element'] = 'variable';
    span.dataset['variableKey'] = key;
    span.className = 'doc-variable-token';
    span.textContent = val;
    span.tabIndex = 0;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && this.editorRef?.nativeElement.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(span);
      range.setStartAfter(span);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      this.editorRef?.nativeElement.append(span);
    }

    this.wordDoc.logAction('variable', key);
    this.showVariablesPanel.set(false);
    this.onEditorInput();
  }

  selectAllText(): void {
    if (!this.editorRef) return;
    const r = document.createRange();
    r.selectNodeContents(this.editorRef.nativeElement);
    const s = window.getSelection();
    if (s) {
      s.removeAllRanges();
      s.addRange(r);
    }
    this.wordDoc.logAction('select', this.editorRef.nativeElement.innerText || '');
    this.closeMenus();
  }

  setOrientation(o: 'portrait' | 'landscape'): void {
    this.wordDoc.setup.orientation = o;
    this.docSetup.set(clone(this.wordDoc.setup));
    this.wordDoc.logAction('pageSetup');
    this.closeMenus();
  }

  openPageSetupDialog(): void {
    this.tempSetup = clone(this.wordDoc.setup);
    this.wordDoc.logAction('openPageSetup');
    this.showPageSetup.set(true);
    this.closeMenus();
  }

  savePageSetup(): void {
    this.wordDoc.setup = clone(this.tempSetup);
    this.docSetup.set(clone(this.wordDoc.setup));
    this.wordDoc.logAction('pageSetup');
    this.showPageSetup.set(false);
  }

  findTab = signal<'headings' | 'find' | 'replace'>('replace');
  openFindReplaceDialog(tab: 'find' | 'replace' = 'replace'): void {
    this.findTab.set(tab);
    this.updateMatchCount();
    this.wordDoc.logAction('openFindReplace');
    this.showFindReplace.set(true);
  }

  private getDocumentMatches(): Range[] {
    if (!this.findTextQuery || !this.editorRef) return [];

    const roots: HTMLElement[] = [this.editorRef.nativeElement];
    if (this.hasPageBreak() && this.page2EditorRef?.nativeElement) {
      roots.push(this.page2EditorRef.nativeElement);
    }

    const query = this.findTextQuery;
    const matchCase = this.matchCase();
    const allRanges: Range[] = [];

    for (const root of roots) {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const nodes: { n: Text; start: number }[] = [];
      let n: Text | null;
      let all = '';

      while ((n = walker.nextNode() as Text | null)) {
        nodes.push({ n, start: all.length });
        all += n.textContent || '';
      }

      const searchStr = matchCase ? all : all.toLowerCase();
      const target = matchCase ? query : query.toLowerCase();

      let pos = 0;
      while ((pos = searchStr.indexOf(target, pos)) !== -1) {
        const start = nodes.find((x) => x.start + x.n.length > pos);
        const end = nodes.find((x) => x.start + x.n.length >= pos + target.length);
        if (start && end) {
          const r = document.createRange();
          r.setStart(start.n, pos - start.start);
          r.setEnd(end.n, pos + target.length - end.start);
          allRanges.push(r);
        }
        pos += target.length;
      }
    }

    return allRanges;
  }

  private selectAndScrollMatch(r: Range): void {
    const el = r.startContainer.parentElement;
    if (el) {
      el.scrollIntoView({ block: 'nearest' });
    }
  }

  private applyFindHighlights(ranges: Range[], currentIndex: number): void {
    if (typeof CSS === 'undefined' || !(CSS as any).highlights || typeof (window as any).Highlight === 'undefined') {
      return;
    }
    try {
      if (ranges.length === 0) {
        this.clearFindHighlights();
        return;
      }
      const HighlightCtor = (window as any).Highlight;
      const allHl = new HighlightCtor(...ranges);
      allHl.priority = 1;
      (CSS as any).highlights.set('find-all', allHl);

      const curRange = (currentIndex >= 1 && currentIndex <= ranges.length)
        ? ranges[currentIndex - 1]
        : ranges[0];
      if (curRange) {
        const curHl = new HighlightCtor(curRange);
        curHl.priority = 2;
        (CSS as any).highlights.set('find-current', curHl);
      } else {
        (CSS as any).highlights.delete('find-current');
      }
    } catch {
      // CSS Highlight API failed, skip highlight
    }
  }

  private clearFindHighlights(): void {
    if (typeof CSS === 'undefined' || !(CSS as any).highlights) {
      return;
    }
    try {
      (CSS as any).highlights.delete('find-all');
      (CSS as any).highlights.delete('find-current');
    } catch {
      // ignore
    }
  }

  private refocusSearchBox(): void {
    const input = document.querySelector<HTMLInputElement>('.nav-search-field');
    if (input) {
      input.focus();
      setTimeout(() => input.focus(), 0);
    }
  }

  updateMatchCount(resetIndex: boolean = true): void {
    if (!this.findTextQuery || !this.editorRef) {
      this.clearFindHighlights();
      this.matchCount.set(0);
      this.currentMatchIndex.set(0);
      return;
    }
    const ranges = this.getDocumentMatches();
    const total = ranges.length;
    this.matchCount.set(total);
    if (total === 0) {
      this.clearFindHighlights();
      this.currentMatchIndex.set(0);
    } else {
      if (resetIndex || this.currentMatchIndex() < 1 || this.currentMatchIndex() > total) {
        this.currentMatchIndex.set(1);
      }
      const curIdx = this.currentMatchIndex();
      this.selectAndScrollMatch(ranges[curIdx - 1]);
      this.applyFindHighlights(ranges, curIdx);
    }
  }

  findNext(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    this.wordDoc.logAction('find', this.findTextQuery);
    const ranges = this.getDocumentMatches();
    this.matchCount.set(ranges.length);
    if (ranges.length === 0) {
      this.clearFindHighlights();
      this.currentMatchIndex.set(0);
      return;
    }
    const cur = this.currentMatchIndex();
    const next = (cur >= ranges.length || cur < 1) ? 1 : cur + 1;
    this.currentMatchIndex.set(next);
    this.selectAndScrollMatch(ranges[next - 1]);
    this.applyFindHighlights(ranges, next);
    this.refocusSearchBox();
  }

  findPrev(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    this.wordDoc.logAction('find', this.findTextQuery);
    const ranges = this.getDocumentMatches();
    this.matchCount.set(ranges.length);
    if (ranges.length === 0) {
      this.clearFindHighlights();
      this.currentMatchIndex.set(0);
      return;
    }
    const cur = this.currentMatchIndex();
    const prev = (cur <= 1 || cur > ranges.length) ? ranges.length : cur - 1;
    this.currentMatchIndex.set(prev);
    this.selectAndScrollMatch(ranges[prev - 1]);
    this.applyFindHighlights(ranges, prev);
    this.refocusSearchBox();
  }

  executeReplace(): void {
    if (!this.findTextQuery || !this.editorRef) return;

    const ranges = this.getDocumentMatches();
    if (ranges.length === 0) {
      this.clearFindHighlights();
      this.matchCount.set(0);
      this.currentMatchIndex.set(0);
      return;
    }

    const cur = Math.max(1, Math.min(this.currentMatchIndex(), ranges.length));
    const targetRange = ranges[cur - 1];

    targetRange.deleteContents();
    targetRange.insertNode(document.createTextNode(this.replaceTextQuery || ''));

    this.wordDoc.logAction('find', this.findTextQuery);
    this.wordDoc.logAction('findReplace', this.findTextQuery);

    this.onEditorInput();

    const newRanges = this.getDocumentMatches();
    this.matchCount.set(newRanges.length);
    if (newRanges.length === 0) {
      this.clearFindHighlights();
      this.currentMatchIndex.set(0);
    } else {
      const nextIndex = cur > newRanges.length ? 1 : cur;
      this.currentMatchIndex.set(nextIndex);
      this.selectAndScrollMatch(newRanges[nextIndex - 1]);
      this.applyFindHighlights(newRanges, nextIndex);
    }
  }

  executeReplaceAll(): void {
    if (!this.findTextQuery || !this.editorRef) return;
    const ranges = this.getDocumentMatches();
    if (ranges.length > 0) {
      const replacement = this.replaceTextQuery || '';
      for (let i = ranges.length - 1; i >= 0; i--) {
        ranges[i].deleteContents();
        ranges[i].insertNode(document.createTextNode(replacement));
      }
    }
    this.wordDoc.logAction('find', this.findTextQuery);
    this.wordDoc.logAction('findReplace', this.findTextQuery);
    this.updateMatchCount();
    this.showFindReplace.set(false);
    this.onEditorInput();
  }

  selectFontColorSwatch(colorKey: string): void {
    const hex = COLORS[colorKey] || colorKey;
    this.selectFontColor(hex, colorKey);
  }

  selectHighlightColorSwatch(colorKey: string): void {
    const hex = COLORS[colorKey] || colorKey;
    this.selectHighlightColor(hex, colorKey);
  }

  updateActiveFormatting(): void {
    this.wordDoc.rich?.capture();
    const s = window.getSelection();
    if (!s || !s.anchorNode) return;
    const el = (s.anchorNode.nodeType === Node.ELEMENT_NODE
      ? s.anchorNode
      : s.anchorNode.parentElement) as HTMLElement | null;
    if (el && this.editorRef?.nativeElement.contains(el)) {
      const cs = window.getComputedStyle(el);
      if (cs.color) {
        this.activeFontColor.set(cs.color);
      }
      if (cs.backgroundColor && cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent') {
        this.activeHighlightColor.set(cs.backgroundColor);
      }
      const table = el.closest('table') as HTMLTableElement | null;
      this.selectedTableEl.set(table);
      const cell = el.closest('td, th') as HTMLTableCellElement | null;
      if (cell) this.selectedTableCell.set(cell);
      if (!table && (this.activeRibbonTab() === 'Table Design' || this.activeRibbonTab() === 'Table Layout')) {
        this.activeRibbonTab.set('Home');
      }
    }
  }

  openDemoModal(): void {
    const curIdx = this.taskIndex();
    const existing = this.taskStartSnapshots.get(curIdx);
    if (existing) {
      this.currentTaskSnapshot.set(existing);
    }
    this.isDemoOpen.set(true);
  }

  onCheckWork(): void {
    if (this.bridge.isPaused() || this.isCompleted()) return;
    const task = this.currentTask();
    if (!task) return;

    let allPassed = false;
    try {
      allPassed = task.goals.length === 0 || task.goals.every((g) => this.wordDoc.check(g));
    } catch (err: any) {
      this.bridge.sendError('Goal check failed: ' + err?.message);
    }

    if (allPassed) {
      this.selectedImageEl.set(null);
      this.isPassed.set(true);
      this.feedbackMsg.set(task.goals.length === 0 ? '' : "Well done! You've got this!");
      this.confetti.celebrate();

      const passedSet = new Set(this.passedTasks());
      passedSet.add(this.taskIndex());
      this.passedTasks.set(passedSet);

      const total = this.totalTasks();
      const isLast = this.taskIndex() === total - 1;

      if (this.isAssessment()) {
        if (!isLast) {
          if (this.advanceTimer) clearTimeout(this.advanceTimer);
          this.advanceTimer = setTimeout(() => {
            this.loadTaskAtIndex(this.taskIndex() + 1);
          }, 1000);
        } else {
          const passedCount = passedSet.size;
          const allTasksPassed = passedCount === total;
          this.bridge.sendComplete(allTasksPassed, passedCount, total);
          this.isCompleted.set(true);
        }
      } else {
        if (total > 1 && !isLast) {
          if (this.advanceTimer) clearTimeout(this.advanceTimer);
          this.advanceTimer = setTimeout(() => {
            this.loadTaskAtIndex(this.taskIndex() + 1);
          }, 1500);
        } else {
          const passedCount = passedSet.size;
          const allTasksPassed = passedCount === total;
          this.bridge.sendComplete(allTasksPassed, passedCount, total);
          this.isCompleted.set(true);
        }
      }
    } else {
      this.isPassed.set(false);
      this.feedbackMsg.set(task.hint || 'Not quite right. Try again.');
    }
  }

  onStartAgain(): void {
    if (this.advanceTimer) {
      clearTimeout(this.advanceTimer);
      this.advanceTimer = null;
    }
    this.isDemoOpen.set(false);
    this.closeMenus();
    this.showPageSetup.set(false);
    this.showFindReplace.set(false);
    this.clearFindHighlights();
    this.showCommentDialog.set(false);
    this.showOrientationPicker.set(false);
    this.activeRibbonTab.set('Home');
    this.selectedImageEl.set(null);
    this.selectedTableEl.set(null);
    this.selectedTocEl.set(null);
    this.selectedPageNum.set(false);
    this.findTextQuery = '';
    this.replaceTextQuery = '';
    this.bridge.resetCompleted();

    window.getSelection()?.removeAllRanges();
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }

    const act = this.currentActivity();
    if (act) {
      this.loadWordActivity(act);
    }
    window.getSelection()?.removeAllRanges();
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    this.cdr.detectChanges();
  }

  // --- Movable Task Card & Mascot Group Positioning ---

  effectiveX(): number | null {
    return this.groupX();
  }

  effectiveY(): number | null {
    return this.groupY();
  }

  computeInitialPlacement(): void {
    if (!this.workLayoutRef) return;
    const layoutEl = this.workLayoutRef.nativeElement;
    const layoutRect = layoutEl.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupWidth = groupEl?.offsetWidth || 300;
    const groupHeight = groupEl?.offsetHeight || 441;

    // Check if assessment strip is visible on the left
    const stripEl = layoutEl.querySelector('.assessment-strip') as HTMLElement | null;
    const leftOffset = stripEl ? stripEl.offsetWidth : 0;

    // Find status bar to ensure mascot is at least 24px above it
    const statusBar = layoutEl.querySelector('.word-statusbar') as HTMLElement | null;
    const bottomLimit = statusBar
      ? (statusBar.getBoundingClientRect().top - layoutRect.top)
      : layoutRect.height;

    // Sit at the bottom-left of the work area: left: 14px, at least 24px above status bar
    const actualHeight = groupEl?.getBoundingClientRect().height || groupHeight;
    const targetX = Math.max(8, 14 + leftOffset);
    const targetY = Math.max(8, Math.floor(bottomLimit - actualHeight - 24));

    this.groupX.set(Math.round(this.clampX(targetX)));
    this.groupY.set(Math.round(this.clampY(targetY)));
  }

  clampX(x: number): number {
    if (!this.workLayoutRef) return Math.max(8, x);
    const layoutRect = this.workLayoutRef.nativeElement.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupWidth = groupEl?.offsetWidth || 300;
    const minX = 8;
    const maxX = Math.max(minX, layoutRect.width - groupWidth - 8);
    return Math.min(Math.max(x, minX), maxX);
  }

  clampY(y: number): number {
    if (!this.workLayoutRef) return Math.max(8, y);
    const layoutRect = this.workLayoutRef.nativeElement.getBoundingClientRect();
    const groupEl = this.movableGroupRef?.nativeElement;
    const groupHeight = groupEl?.offsetHeight || 441;
    const actualHeight = groupEl?.getBoundingClientRect().height || groupHeight;
    const statusBar = this.workLayoutRef.nativeElement.querySelector('.word-statusbar') as HTMLElement | null;
    const bottomLimit = statusBar
      ? (statusBar.getBoundingClientRect().top - layoutRect.top)
      : layoutRect.height;
    const minY = 8;
    const maxY = Math.max(minY, Math.floor(bottomLimit - actualHeight - 24));
    return Math.min(Math.max(y, minY), maxY);
  }

  onGroupPointerDown(event: PointerEvent): void {
    const target = event.target as HTMLElement | null;
    const isDragHandle = target?.closest('.card-drag-handle');
    if (!isDragHandle) {
      return; // Do not initiate drag from regular buttons or card body
    }

    this.isDragging.set(true);
    this.dragStartPointerX = event.clientX;
    this.dragStartPointerY = event.clientY;

    const groupEl = this.movableGroupRef?.nativeElement;
    this.dragStartGroupX = this.groupX() ?? groupEl?.offsetLeft ?? 0;
    this.dragStartGroupY = this.groupY() ?? groupEl?.offsetTop ?? 0;

    try {
      (event.currentTarget as HTMLElement)?.setPointerCapture(event.pointerId);
    } catch {}
    event.preventDefault();
  }

  onGroupPointerMove(event: PointerEvent): void {
    if (!this.isDragging()) return;
    const dx = event.clientX - this.dragStartPointerX;
    const dy = event.clientY - this.dragStartPointerY;
    const nextX = this.clampX(this.dragStartGroupX + dx);
    const nextY = this.clampY(this.dragStartGroupY + dy);
    this.groupX.set(Math.round(nextX));
    this.groupY.set(Math.round(nextY));
    this.hasUserMoved.set(true);
  }

  onGroupPointerUp(event: PointerEvent): void {
    if (this.isDragging()) {
      this.isDragging.set(false);
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  onGroupPointerCancel(event: PointerEvent): void {
    if (this.isDragging()) {
      this.isDragging.set(false);
      try {
        (event.currentTarget as HTMLElement)?.releasePointerCapture(event.pointerId);
      } catch {}
    }
  }

  onNudgePosition(delta: { dx: number; dy: number }): void {
    const groupEl = this.movableGroupRef?.nativeElement;
    const curX = this.groupX() ?? groupEl?.offsetLeft ?? 14;
    const curY = this.groupY() ?? groupEl?.offsetTop ?? 100;
    this.groupX.set(Math.round(this.clampX(curX + delta.dx)));
    this.groupY.set(Math.round(this.clampY(curY + delta.dy)));
    this.hasUserMoved.set(true);
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (!this.hasUserMoved()) {
      this.computeInitialPlacement();
    } else if (this.groupX() !== null && this.groupY() !== null) {
      this.groupX.set(Math.round(this.clampX(this.groupX()!)));
      this.groupY.set(Math.round(this.clampY(this.groupY()!)));
    }
    if (this.selectedChartEl()) {
      this.updateChartBoxRect();
    }
  }

  // --- Chart Management Methods ---

  openInsertChartDialog(): void {
    this.chartDialogMode.set('insert');
    this.chartDialogInitialType.set('column');
    this.chartDialogInitialTitle.set('Chart Title');
    this.chartDialogInitialData.set(null);
    this.showInsertChartDialog.set(true);
    this.wordDoc.logAction('openInsertChartDialog');
  }

  openChangeChartTypeDialog(): void {
    const chart = this.selectedChartEl();
    if (!chart) return;
    const type = (chart.getAttribute('data-chart-type') || 'column') as WordChartType;
    const title = chart.getAttribute('data-chart-title') || 'Chart Title';
    let data: WordChartData | null = null;
    const dataRaw = chart.getAttribute('data-chart-data');
    if (dataRaw) {
      try {
        data = JSON.parse(dataRaw);
      } catch {}
    }
    this.chartDialogMode.set('change');
    this.chartDialogInitialType.set(type);
    this.chartDialogInitialTitle.set(title);
    this.chartDialogInitialData.set(data);
    this.showInsertChartDialog.set(true);
    this.wordDoc.logAction('openChangeChartTypeDialog');
  }

  onChartDialogResult(result: ChartDialogResult): void {
    if (this.chartDialogMode() === 'insert') {
      this.insertChart(result.chartType, result.title, result.data);
    } else {
      this.changeSelectedChart(result.chartType, result.title, result.data);
    }
  }

  insertChart(type: WordChartType, title: string, data?: WordChartData): void {
    this.saveTableUndo();
    const chartData = data || DEFAULT_CHART_DATA;
    const figure = document.createElement('figure');
    figure.className = 'doc-chart selected';
    figure.setAttribute('data-chart-type', type);
    figure.setAttribute('data-chart-title', title);
    figure.setAttribute('data-chart-data', JSON.stringify(chartData));
    figure.setAttribute('contenteditable', 'false');
    figure.style.width = '480px';
    figure.style.height = '300px';

    figure.innerHTML = renderChartSvg(type, title, chartData, 480, 300);

    const allCharts = Array.from(document.querySelectorAll('figure.doc-chart'));
    allCharts.forEach(c => c.classList.remove('selected'));

    const targetEditor = this.activePageNum() === 2 && this.page2EditorRef
      ? this.page2EditorRef.nativeElement
      : this.editorRef?.nativeElement;

    const sel = window.getSelection();
    let inserted = false;
    if (sel && sel.rangeCount > 0 && targetEditor && targetEditor.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.collapse(false);
      range.insertNode(figure);
      inserted = true;
    }
    if (!inserted && targetEditor) {
      targetEditor.appendChild(figure);
    }

    this.selectedChartEl.set(figure);
    this.activeRibbonTab.set('Chart Design');
    this.wordDoc.logAction('insertChart', type);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
    setTimeout(() => this.updateChartBoxRect(), 10);
  }

  changeSelectedChart(type: WordChartType, title: string, data?: WordChartData): void {
    const chart = this.selectedChartEl();
    if (!chart) return;
    this.saveTableUndo();
    const chartData = data || DEFAULT_CHART_DATA;
    chart.setAttribute('data-chart-type', type);
    chart.setAttribute('data-chart-title', title);
    chart.setAttribute('data-chart-data', JSON.stringify(chartData));

    const w = chart.offsetWidth || 480;
    const h = chart.offsetHeight || 300;
    chart.innerHTML = renderChartSvg(type, title, chartData, w, h);

    this.wordDoc.logAction('changeChartType', type);
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
    setTimeout(() => this.updateChartBoxRect(), 10);
  }

  deleteSelectedChart(): void {
    const chart = this.selectedChartEl();
    if (!chart) return;
    this.saveTableUndo();
    chart.remove();
    this.selectedChartEl.set(null);
    this.isEditingChartTitle.set(false);
    if (this.activeRibbonTab() === 'Chart Design' || this.activeRibbonTab() === 'Format') {
      this.activeRibbonTab.set('Home');
    }
    this.wordDoc.logAction('deleteChart');
    this.onEditorInput();
    this.lastTablePostHtml = this.getCurrentDocumentHtml();
  }

  updateChartBoxRect(): void {
    const chart = this.selectedChartEl();
    if (!chart) return;
    const canvas = document.querySelector('.word-canvas');
    if (!canvas) return;
    const canvasRect = canvas.getBoundingClientRect();
    const chartRect = chart.getBoundingClientRect();

    const scrollLeft = canvas.scrollLeft || 0;
    const scrollTop = canvas.scrollTop || 0;

    const left = chartRect.left - canvasRect.left + scrollLeft;
    const top = chartRect.top - canvasRect.top + scrollTop;
    const width = chartRect.width;
    const height = chartRect.height;

    this.chartBoxLeft.set(left);
    this.chartBoxTop.set(top);
    this.chartBoxWidth.set(width);
    this.chartBoxHeight.set(height);
  }

  startChartTitleEdit(): void {
    const chart = this.selectedChartEl();
    if (!chart) return;
    const currentTitle = chart.getAttribute('data-chart-title') || 'Chart Title';
    this.chartTitleEditingText.set(currentTitle);
    this.isEditingChartTitle.set(true);
    setTimeout(() => {
      const input = document.querySelector('.chart-title-inline-field') as HTMLInputElement | null;
      if (input) {
        input.focus();
        input.select();
      }
    }, 20);
  }

  onInlineTitleInput(val: string): void {
    this.chartTitleEditingText.set(val);
    const chart = this.selectedChartEl();
    if (chart) {
      chart.setAttribute('data-chart-title', val);
      const titleEl = chart.querySelector('.chart-title-text');
      if (titleEl) {
        titleEl.textContent = val;
      }
    }
  }

  finishChartTitleEdit(): void {
    if (!this.isEditingChartTitle()) return;
    const chart = this.selectedChartEl();
    const newTitle = this.chartTitleEditingText().trim() || 'Chart Title';
    if (chart) {
      this.saveTableUndo();
      chart.setAttribute('data-chart-title', newTitle);
      const titleEl = chart.querySelector('.chart-title-text');
      if (titleEl) {
        titleEl.textContent = newTitle;
      }
      this.wordDoc.logAction('editChartTitle', newTitle);
      this.onEditorInput();
      this.lastTablePostHtml = this.getCurrentDocumentHtml();
    }
    this.isEditingChartTitle.set(false);
  }

  cancelChartTitleEdit(): void {
    if (!this.isEditingChartTitle()) return;
    const chart = this.selectedChartEl();
    if (chart) {
      const originalTitle = chart.getAttribute('data-chart-title') || 'Chart Title';
      const titleEl = chart.querySelector('.chart-title-text');
      if (titleEl) {
        titleEl.textContent = originalTitle;
      }
    }
    this.isEditingChartTitle.set(false);
  }

  onChartResizeStart(event: MouseEvent, handle: string): void {
    event.preventDefault();
    event.stopPropagation();
    const chart = this.selectedChartEl();
    if (!chart) return;
    this.saveTableUndo();

    const startX = event.clientX;
    const startY = event.clientY;
    const startWidth = chart.offsetWidth;
    const startHeight = chart.offsetHeight;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const dx = moveEvent.clientX - startX;
      const dy = moveEvent.clientY - startY;

      let newWidth = startWidth;
      let newHeight = startHeight;

      if (handle.includes('e')) newWidth = Math.max(200, startWidth + dx);
      if (handle.includes('w')) newWidth = Math.max(200, startWidth - dx);
      if (handle.includes('s')) newHeight = Math.max(150, startHeight + dy);
      if (handle.includes('n')) newHeight = Math.max(150, startHeight - dy);

      chart.style.width = `${newWidth}px`;
      chart.style.height = `${newHeight}px`;

      const type = (chart.getAttribute('data-chart-type') || 'column') as WordChartType;
      const title = chart.getAttribute('data-chart-title') || 'Chart Title';
      let data: WordChartData = DEFAULT_CHART_DATA;
      const dataRaw = chart.getAttribute('data-chart-data');
      if (dataRaw) {
        try {
          data = JSON.parse(dataRaw);
        } catch {}
      }
      chart.innerHTML = renderChartSvg(type, title, data, newWidth, newHeight);
      this.updateChartBoxRect();
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      this.wordDoc.logAction('resizeChart');
      this.onEditorInput();
      this.lastTablePostHtml = this.getCurrentDocumentHtml();
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }
}

