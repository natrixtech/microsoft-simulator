import {
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnDestroy,
  OnInit,
  Output,
  ViewChild,
  computed,
  signal,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  CellCoord,
  CellData,
  CellFormat,
  ChartItem,
  ChartType,
  DataValidationRule,
  PivotTableItem,
  SheetFilter,
  SheetSnapshot,
} from '../../models/excel.types';
import { ExcelStateService } from '../../services/excel-state.service';
import { ExcelClipboardService } from '../../services/excel-clipboard.service';
import { ExcelHistoryService } from '../../services/excel-history.service';
import { ExcelActionLoggerService } from '../../services/excel-action-logger.service';
import {
  colLetterToIndex,
  formatAddress,
  indexToColLetter,
  parseAddress,
  parseDateParts,
  parseRange,
} from '../../services/formula-engine.service';
import { COLORS, PluginBridgeService } from 'shared';

@Component({
  selector: 'app-excel-grid',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './grid.component.html',
  styleUrls: ['./grid.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class GridComponent implements OnInit, AfterViewInit, OnDestroy {
  @ViewChild('gridScroll') gridScrollRef?: ElementRef<HTMLDivElement>;
  @ViewChild('cellEditorInput') cellEditorInputRef?: ElementRef<HTMLInputElement>;

  @Input() isCompleted = false;

  @Output() readonly openFilterDropdown = new EventEmitter<{ colIndex: number; anchorX: number; anchorY: number }>();
  @Output() readonly openCommentPopup = new EventEmitter<string>();
  @Output() readonly validationError = new EventEmitter<{ allowed: string[] }>();

  readonly numCols = 26; // A to Z
  readonly numRows = 50; // 1 to 50
  readonly defaultColWidth = 84;
  readonly defaultRowHeight = 24;
  readonly rowHeight = 24;

  readonly cols = Array.from({ length: 26 }, (_, i) => ({
    index: i,
    letter: indexToColLetter(i),
  }));

  readonly rows = Array.from({ length: 50 }, (_, i) => ({
    index: i,
    num: i + 1,
  }));

  // Column resizing state
  private isResizingCol = false;
  private resizeColIndex = -1;
  private resizeStartX = 0;
  private resizeStartWidth = 0;
  private colResizeSnapshot: SheetSnapshot | null = null;

  // Row resizing state
  private isResizingRow = false;
  private resizeRowIndex = -1;
  private resizeStartY = 0;
  private resizeStartHeight = 0;
  private rowResizeSnapshot: SheetSnapshot | null = null;

  // Selection state
  isSelectingRange = false;
  private rangeAnchorCoord: CellCoord = { row: 0, col: 0 };
  private selectionAnchor: CellCoord = { row: 0, col: 0 };
  private selectionLead: CellCoord = { row: 0, col: 0 };

  // Edit mode: 'type' (arrow keys commit and move) vs 'full' (F2 or double-click, left/right move cursor)
  readonly editMode = signal<'type' | 'full'>('type');

  // Chart dragging state
  private isDraggingChart = false;
  private draggedChartId: string | null = null;
  private chartDragStartX = 0;
  private chartDragStartY = 0;
  private chartInitialX = 0;
  private chartInitialY = 0;

  // Active validation dropdown menu
  readonly activeValidationMenu = signal<{ cell: string; allowed: string[]; x: number; y: number } | null>(null);

  // Cell comment hover note state
  readonly hoveredCommentAddr = signal<string | null>(null);
  readonly hoveredComment = signal<HoveredCommentNote | null>(null);

  private wheelListener = (event: WheelEvent) => this.onGridWheel(event);

  constructor(
    public state: ExcelStateService,
    private clipboard: ExcelClipboardService,
    private history: ExcelHistoryService,
    private actionLogger: ExcelActionLoggerService,
    private bridge: PluginBridgeService,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    const cur = parseAddress(this.state.activeCell()) || { row: 0, col: 0 };
    this.selectionAnchor = { row: cur.row, col: cur.col };
    this.selectionLead = { row: cur.row, col: cur.col };
  }

  ngAfterViewInit(): void {
    this.gridScrollRef?.nativeElement.addEventListener('wheel', this.wheelListener, { passive: false });
  }

  ngOnDestroy(): void {
    this.gridScrollRef?.nativeElement.removeEventListener('wheel', this.wheelListener);
  }

  focus(): void {
    this.gridScrollRef?.nativeElement.focus({ preventScroll: true });
  }

  syncSelection(rangeOrCell: string): void {
    const clean = rangeOrCell.trim().toUpperCase();
    if (clean.includes(':')) {
      const parts = clean.split(':');
      const pStart = parseAddress(parts[0]);
      const pEnd = parseAddress(parts[1]);
      if (pStart) this.selectionAnchor = { row: pStart.row, col: pStart.col };
      if (pEnd) this.selectionLead = { row: pEnd.row, col: pEnd.col };
    } else {
      const p = parseAddress(clean);
      if (p) {
        this.selectionAnchor = { row: p.row, col: p.col };
        this.selectionLead = { row: p.row, col: p.col };
      }
    }
  }

  scrollToActiveCell(): void {
    this.cdr.detectChanges();
    const activeEl = this.gridScrollRef?.nativeElement.querySelector<HTMLElement>('.excel-cell.active-cell');
    if (activeEl) {
      this.scrollCellIntoView(activeEl);
    }
  }

  scrollToCell(row: number, col: number): void {
    this.cdr.detectChanges();
    const addr = formatAddress({ row, col });
    const cellEl = this.gridScrollRef?.nativeElement.querySelector<HTMLElement>(`[data-cell-address="${addr}"]`);
    if (cellEl) {
      this.scrollCellIntoView(cellEl);
    }
  }

  private scrollCellIntoView(cellEl: HTMLElement): void {
    const container = this.gridScrollRef?.nativeElement;
    if (!container) return;

    const containerRect = container.getBoundingClientRect();
    const cellRect = cellEl.getBoundingClientRect();

    const cellTop = (cellRect.top - containerRect.top) + container.scrollTop;
    const cellBottom = cellTop + cellRect.height;
    const cellLeft = (cellRect.left - containerRect.left) + container.scrollLeft;
    const cellRight = cellLeft + cellRect.width;

    const headerTop = 24;
    const headerLeft = 44;

    const scrollbarBottom = container.offsetHeight - container.clientHeight;
    const scrollbarRight = container.offsetWidth - container.clientWidth;

    const visibleTop = container.scrollTop + headerTop;
    const visibleBottom = container.scrollTop + (container.offsetHeight - scrollbarBottom);
    const visibleLeft = container.scrollLeft + headerLeft;
    const visibleRight = container.scrollLeft + (container.offsetWidth - scrollbarRight);

    if (cellTop < visibleTop) {
      container.scrollTop = Math.max(0, cellTop - headerTop);
    } else if (cellBottom > visibleBottom) {
      container.scrollTop = Math.max(0, cellBottom - (container.offsetHeight - scrollbarBottom));
    }

    if (cellLeft < visibleLeft) {
      container.scrollLeft = Math.max(0, cellLeft - headerLeft);
    } else if (cellRight > visibleRight) {
      container.scrollLeft = Math.max(0, cellRight - (container.offsetWidth - scrollbarRight));
    }
  }

  onGridWheel(event: WheelEvent): void {
    const el = this.gridScrollRef?.nativeElement;
    if (!el) return;
    let dy = event.deltaY;
    let dx = event.deltaX;
    if (event.deltaMode === 1) {
      dy *= 24;
      dx *= 24;
    } else if (event.deltaMode === 2) {
      dy *= el.clientHeight;
      dx *= el.clientWidth;
    }
    el.scrollTop += dy;
    el.scrollLeft += dx;
    event.preventDefault();
  }

  getColWidth(letter: string): number {
    return this.state.columnWidths().get(letter) || this.defaultColWidth;
  }

  getRowHeight(rowIdx: number): number {
    return this.state.rowHeights().get(rowIdx) || this.defaultRowHeight;
  }

  isRowHidden(rowIdx: number): boolean {
    return this.state.isRowFilteredOut(rowIdx);
  }

  // Address and Selection queries
  isCellSelected(row: number, col: number): boolean {
    const s = this.state.selectionCoords();
    return row >= s.startRow && row <= s.endRow && col >= s.startCol && col <= s.endCol;
  }

  isActiveCell(row: number, col: number): boolean {
    const active = parseAddress(this.state.activeCell());
    if (!active) return false;
    return active.row === row && active.col === col;
  }

  getCellData(row: number, col: number): CellData | undefined {
    const addr = formatAddress({ row, col });
    return this.state.getCell(addr);
  }

  getCellDisplay(row: number, col: number): string {
    const cell = this.getCellData(row, col);
    if (!cell) return '';
    return cell.displayValue;
  }

  getCellStyle(row: number, col: number): Record<string, string> {
    const addr = formatAddress({ row, col });
    const fmt = this.state.cellFormats().get(addr);
    const styles: Record<string, string> = {};

    if (fmt?.bold) styles['font-weight'] = 'bold';
    if (fmt?.italic) styles['font-style'] = 'italic';
    if (fmt?.underline) styles['text-decoration'] = 'underline';
    styles['font-size'] = `${fmt?.fontSize ?? 10}pt`;
    if (fmt?.fontFamily) styles['font-family'] = fmt.fontFamily;
    if (fmt?.align) {
      styles['text-align'] = fmt.align;
    } else {
      const cell = this.state.getCell(addr);
      const val = cell?.value;
      if (val !== undefined && val !== null && val !== '') {
        if (typeof val === 'number' || parseDateParts(val) !== null) {
          styles['text-align'] = 'right';
        } else {
          styles['text-align'] = 'left';
        }
      } else {
        styles['text-align'] = 'left';
      }
    }
    if (fmt?.textColor) {
      styles['color'] = COLORS[fmt.textColor] || fmt.textColor;
    }

    // Fill Color or Conditional Format Color
    const condColor = this.state.getConditionalFillColor(addr);
    if (condColor) {
      styles['background-color'] = condColor;
    } else if (fmt?.fillColor) {
      styles['background-color'] = COLORS[fmt.fillColor] || fmt.fillColor;
    }
    if (fmt?.border) {
      if (fmt.border.top) {
        styles['border-top'] = `${fmt.border.top === 'thick' ? 2 : 1}px solid var(--excel-user-border, #000000)`;
      }
      if (fmt.border.bottom) {
        styles['border-bottom'] = `${fmt.border.bottom === 'thick' ? 2 : 1}px solid var(--excel-user-border, #000000)`;
      }
      if (fmt.border.left) {
        styles['border-left'] = `${fmt.border.left === 'thick' ? 2 : 1}px solid var(--excel-user-border, #000000)`;
      }
      if (fmt.border.right) {
        styles['border-right'] = `${fmt.border.right === 'thick' ? 2 : 1}px solid var(--excel-user-border, #000000)`;
      }
    }

    return styles;
  }

  hasComment(row: number, col: number): boolean {
    const addr = formatAddress({ row, col });
    return this.state.comments().has(addr);
  }

  hasValidation(row: number, col: number): boolean {
    const addr = formatAddress({ row, col });
    return !!this.state.getDataValidation(addr);
  }

  isHeaderCellWithFilter(row: number, col: number): boolean {
    const f = this.state.filter();
    if (!f) return false;
    return row === f.headerRow && col >= f.startCol && col <= f.endCol;
  }

  // Mouse selection handlers
  onCellMouseDown(row: number, col: number, event: MouseEvent): void {
    if (event.button !== 0) return; // only left click

    if (this.state.isEditing()) {
      this.commitEdit(0, 0);
    }

    const addr = formatAddress({ row, col });

    if (event.shiftKey) {
      // Expand range from anchor
      this.extendSelectionTo(row, col);
    } else {
      this.state.selectCell(addr);
      this.selectionAnchor = { row, col };
      this.selectionLead = { row, col };
      this.rangeAnchorCoord = { row, col };
      this.isSelectingRange = true;
      this.actionLogger.log('selectRange', addr);
    }

    this.activeValidationMenu.set(null);
  }

  onCellMouseEnter(row: number, col: number, event?: MouseEvent): void {
    if (this.isSelectingRange) {
      this.selectionLead = { row, col };
      const range = this.formatBoundingRange(this.rangeAnchorCoord, { row, col });
      this.state.selectRange(range);
      this.actionLogger.log('selectRange', range);
      this.hoveredCommentAddr.set(null);
      this.hoveredComment.set(null);
      return;
    }

    if (this.state.isEditing()) {
      this.hoveredCommentAddr.set(null);
      this.hoveredComment.set(null);
      return;
    }

    const addr = formatAddress({ row, col });
    if (this.state.comments().has(addr)) {
      this.hoveredCommentAddr.set(addr);
      const targetEl = event?.currentTarget as HTMLElement | undefined;
      this.updateCommentNotePosition(addr, targetEl);
    } else {
      this.hoveredCommentAddr.set(null);
      this.hoveredComment.set(null);
    }
  }

  onCellMouseLeave(row: number, col: number): void {
    const addr = formatAddress({ row, col });
    if (this.hoveredCommentAddr() === addr) {
      this.hoveredCommentAddr.set(null);
      this.hoveredComment.set(null);
    }
  }

  onCommentMarkerClick(addr: string, event: MouseEvent): void {
    event.stopPropagation();
    this.hoveredCommentAddr.set(null);
    this.hoveredComment.set(null);
    this.openCommentPopup.emit(addr);
  }

  onGridScroll(): void {
    if (this.hoveredCommentAddr()) {
      this.updateCommentNotePosition(this.hoveredCommentAddr()!);
    }
  }

  @HostListener('document:mouseup')
  onDocumentMouseUp(): void {
    if (this.isResizingCol) {
      this.isResizingCol = false;
      if (this.resizeColIndex >= 0) {
        const letter = indexToColLetter(this.resizeColIndex);
        const curWidth = this.getColWidth(letter);
        if (curWidth !== this.resizeStartWidth && this.colResizeSnapshot) {
          this.history.push(this.colResizeSnapshot);
        }
      }
      this.resizeColIndex = -1;
      this.colResizeSnapshot = null;
    }

    if (this.isResizingRow) {
      this.isResizingRow = false;
      if (this.resizeRowIndex >= 0) {
        const curHeight = this.getRowHeight(this.resizeRowIndex);
        if (curHeight !== this.resizeStartHeight && this.rowResizeSnapshot) {
          this.history.push(this.rowResizeSnapshot);
        }
      }
      this.resizeRowIndex = -1;
      this.rowResizeSnapshot = null;
    }

    this.isSelectingRange = false;
    this.isDraggingChart = false;
    this.draggedChartId = null;
    if (this.hoveredCommentAddr()) {
      this.updateCommentNotePosition(this.hoveredCommentAddr()!);
    }
  }

  private formatBoundingRange(p1: CellCoord, p2: CellCoord): string {
    const r1 = Math.min(p1.row, p2.row);
    const c1 = Math.min(p1.col, p2.col);
    const r2 = Math.max(p1.row, p2.row);
    const c2 = Math.max(p1.col, p2.col);

    const a1 = formatAddress({ row: r1, col: c1 });
    const a2 = formatAddress({ row: r2, col: c2 });
    if (a1 === a2) return a1;
    return `${a1}:${a2}`;
  }

  private extendSelectionTo(targetRow: number, targetCol: number): void {
    this.selectionLead = { row: targetRow, col: targetCol };
    const range = this.formatBoundingRange(this.selectionAnchor, this.selectionLead);
    this.state.selectRange(range);
    this.actionLogger.log('selectRange', range);
    this.scrollToCell(targetRow, targetCol);
  }

  // Double click to edit cell
  onCellDblClick(row: number, col: number): void {
    const addr = formatAddress({ row, col });
    this.startEditing(addr, undefined, 'full');
  }

  startEditing(addr: string, initialChar?: string, mode: 'type' | 'full' = 'type'): void {
    this.hoveredCommentAddr.set(null);
    this.hoveredComment.set(null);
    this.editMode.set(mode);
    const cell = this.state.getCell(addr);
    const rawVal = initialChar !== undefined ? initialChar : (cell?.formula || (cell?.raw !== undefined ? String(cell.raw) : ''));

    this.state.editingAddress.set(addr);
    this.state.editInput.set(rawVal);
    this.state.isEditing.set(true);

    // Synchronously render and focus the inline editor so keystrokes are never lost
    this.cdr.detectChanges();
    const inputEl = this.cellEditorInputRef?.nativeElement;
    if (inputEl) {
      inputEl.focus({ preventScroll: true });
      const len = inputEl.value.length;
      inputEl.setSelectionRange(len, len);
    }
  }

  commitEdit(dRow: number = 0, dCol: number = 0): void {
    if (!this.state.isEditing()) return;

    const snap = this.state.getSnapshot();
    this.history.push(snap);

    const addr = this.state.editingAddress();
    const val = this.state.editInput();

    // Data validation reject check
    const dv = this.state.getDataValidation(addr);
    if (dv && dv.rejectInvalid && val !== '') {
      const isAllowed = dv.allowed.some((a) => a.trim().toLowerCase() === val.trim().toLowerCase());
      if (!isAllowed) {
        this.cancelEdit();
        this.validationError.emit({ allowed: dv.allowed });
        return;
      }
    }

    this.state.setCellValue(addr, val);
    this.state.isEditing.set(false);

    // Move to next cell if dRow != 0 or dCol != 0
    if (dRow !== 0 || dCol !== 0) {
      const p = parseAddress(addr);
      if (p) {
        const nextR = Math.max(0, Math.min(this.numRows - 1, p.row + dRow));
        const nextC = Math.max(0, Math.min(this.numCols - 1, p.col + dCol));
        const nextAddr = formatAddress({ row: nextR, col: nextC });
        this.state.selectCell(nextAddr);
        this.selectionAnchor = { row: nextR, col: nextC };
        this.selectionLead = { row: nextR, col: nextC };
      }
    } else {
      const p = parseAddress(addr);
      if (p) {
        this.selectionAnchor = { row: p.row, col: p.col };
        this.selectionLead = { row: p.row, col: p.col };
      }
    }

    this.focus();
    this.scrollToActiveCell();
  }

  cancelEdit(): void {
    this.state.isEditing.set(false);
    const cur = parseAddress(this.state.activeCell());
    if (cur) {
      this.selectionAnchor = { row: cur.row, col: cur.col };
      this.selectionLead = { row: cur.row, col: cur.col };
    }
    this.focus();
    this.scrollToActiveCell();
  }

  onCellEditorKeyDown(event: KeyboardEvent): void {
    if (this.bridge.isPaused() || this.isCompleted) return;

    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      this.commitEdit(event.shiftKey ? -1 : 1, 0);
      return;
    }

    if (event.key === 'Tab') {
      event.preventDefault();
      event.stopPropagation();
      this.commitEdit(0, event.shiftKey ? -1 : 1);
      return;
    }

    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      this.cancelEdit();
      return;
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault();
      event.stopPropagation();
      this.commitEdit(-1, 0);
      return;
    }

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      event.stopPropagation();
      this.commitEdit(1, 0);
      return;
    }

    if (event.key === 'ArrowLeft') {
      if (this.editMode() === 'type') {
        event.preventDefault();
        event.stopPropagation();
        this.commitEdit(0, -1);
      }
      return;
    }

    if (event.key === 'ArrowRight') {
      if (this.editMode() === 'type') {
        event.preventDefault();
        event.stopPropagation();
        this.commitEdit(0, 1);
      }
      return;
    }
  }

  // Keyboard navigation
  @HostListener('keydown', ['$event'])
  onGridKeyDown(event: KeyboardEvent): void {
    if (this.bridge.isPaused() || this.isCompleted) return;

    // If editing, allow normal input keys or handle commit/cancel
    if (this.state.isEditing()) {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.commitEdit(event.shiftKey ? -1 : 1, 0);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        this.cancelEdit();
      } else if (event.key === 'Tab') {
        event.preventDefault();
        this.commitEdit(0, event.shiftKey ? -1 : 1);
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        this.commitEdit(-1, 0);
      } else if (event.key === 'ArrowDown') {
        event.preventDefault();
        this.commitEdit(1, 0);
      } else if (event.key === 'ArrowLeft' && this.editMode() === 'type') {
        event.preventDefault();
        this.commitEdit(0, -1);
      } else if (event.key === 'ArrowRight' && this.editMode() === 'type') {
        event.preventDefault();
        this.commitEdit(0, 1);
      } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
        if (event.target !== this.cellEditorInputRef?.nativeElement) {
          // Fast typing buffer: append key directly to edit input and sync editor
          event.preventDefault();
          this.state.editInput.update((cur) => cur + event.key);
          const inputEl = this.cellEditorInputRef?.nativeElement;
          if (inputEl) {
            inputEl.value = this.state.editInput();
            inputEl.focus({ preventScroll: true });
            const len = inputEl.value.length;
            inputEl.setSelectionRange(len, len);
          }
        }
      }
      return;
    }

    const cur = parseAddress(this.state.activeCell());
    if (!cur) return;

    // Shortcuts
    if (event.ctrlKey || event.metaKey) {
      const key = event.key.toLowerCase();
      if (key === 'a') {
        event.preventDefault();
        event.stopPropagation();
        window.getSelection()?.removeAllRanges();
        this.state.selectAll();
        this.actionLogger.log('selectRange', this.state.selectedRange());
        return;
      }
      if (key === 'z') {
        event.preventDefault();
        const prev = this.history.undo(this.state.getSnapshot());
        if (prev) this.state.restoreSnapshot(prev);
        return;
      }
      if (key === 'y') {
        event.preventDefault();
        const next = this.history.redo(this.state.getSnapshot());
        if (next) this.state.restoreSnapshot(next);
        return;
      }
      if (key === 'c') {
        event.preventDefault();
        this.clipboard.copy(this.state.selectedRange());
        return;
      }
      if (key === 'x') {
        event.preventDefault();
        this.clipboard.cut(this.state.selectedRange());
        return;
      }
      if (key === 'v') {
        event.preventDefault();
        this.history.push(this.state.getSnapshot());
        this.clipboard.paste(this.state.activeCell());
        return;
      }
      if (key === 'b') {
        event.preventDefault();
        const fmt = this.state.cellFormats().get(this.state.activeCell());
        this.state.formatSelection({ bold: !fmt?.bold });
        return;
      }
      if (key === 'i') {
        event.preventDefault();
        const fmt = this.state.cellFormats().get(this.state.activeCell());
        this.state.formatSelection({ italic: !fmt?.italic });
        return;
      }
      if (key === 'u') {
        event.preventDefault();
        const fmt = this.state.cellFormats().get(this.state.activeCell());
        this.state.formatSelection({ underline: !fmt?.underline });
        return;
      }
      if (key === 'home') {
        event.preventDefault();
        this.state.selectCell('A1');
        this.selectionAnchor = { row: 0, col: 0 };
        this.selectionLead = { row: 0, col: 0 };
        this.actionLogger.log('goToCell', 'A1');
        this.scrollToActiveCell();
        return;
      }
      if (key === 'end') {
        event.preventDefault();
        const lastCell = this.getLastUsedCell();
        this.state.selectCell(lastCell);
        const p = parseAddress(lastCell) || { row: 0, col: 0 };
        this.selectionAnchor = { row: p.row, col: p.col };
        this.selectionLead = { row: p.row, col: p.col };
        this.actionLogger.log('goToCell', lastCell);
        this.scrollToActiveCell();
        return;
      }
      // Ctrl+Arrows or Ctrl+Shift+Arrows jump to edge of data block
      if (
        event.key === 'ArrowUp' ||
        event.key === 'ArrowDown' ||
        event.key === 'ArrowLeft' ||
        event.key === 'ArrowRight'
      ) {
        event.preventDefault();
        let dRow = 0;
        let dCol = 0;
        if (event.key === 'ArrowUp') dRow = -1;
        else if (event.key === 'ArrowDown') dRow = 1;
        else if (event.key === 'ArrowLeft') dCol = -1;
        else if (event.key === 'ArrowRight') dCol = 1;

        const baseRow = event.shiftKey ? this.selectionLead.row : cur.row;
        const baseCol = event.shiftKey ? this.selectionLead.col : cur.col;
        const target = this.findDataBlockEdge(baseRow, baseCol, dRow, dCol);

        if (event.shiftKey) {
          this.extendSelectionTo(target.row, target.col);
        } else {
          const nextAddr = formatAddress(target);
          this.state.selectCell(nextAddr);
          this.selectionAnchor = { ...target };
          this.selectionLead = { ...target };
          this.scrollToActiveCell();
        }
        return;
      }
    }

    if (event.key === 'Home') {
      event.preventDefault();
      if (event.shiftKey) {
        this.extendSelectionTo(cur.row, 0);
      } else {
        const nextAddr = formatAddress({ row: cur.row, col: 0 });
        this.state.selectCell(nextAddr);
        this.selectionAnchor = { row: cur.row, col: 0 };
        this.selectionLead = { row: cur.row, col: 0 };
        this.scrollToActiveCell();
      }
      return;
    }

    if (event.key === 'PageDown') {
      event.preventDefault();
      const vRows = this.getVisibleRowCount();
      const targetRow = Math.min(this.numRows - 1, cur.row + vRows);
      if (event.shiftKey) {
        this.extendSelectionTo(targetRow, cur.col);
      } else {
        const nextAddr = formatAddress({ row: targetRow, col: cur.col });
        this.state.selectCell(nextAddr);
        this.selectionAnchor = { row: targetRow, col: cur.col };
        this.selectionLead = { row: targetRow, col: cur.col };
        this.scrollToActiveCell();
      }
      return;
    }

    if (event.key === 'PageUp') {
      event.preventDefault();
      const vRows = this.getVisibleRowCount();
      const targetRow = Math.max(0, cur.row - vRows);
      if (event.shiftKey) {
        this.extendSelectionTo(targetRow, cur.col);
      } else {
        const nextAddr = formatAddress({ row: targetRow, col: cur.col });
        this.state.selectCell(nextAddr);
        this.selectionAnchor = { row: targetRow, col: cur.col };
        this.selectionLead = { row: targetRow, col: cur.col };
        this.scrollToActiveCell();
      }
      return;
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault();
      this.handleArrowKey(-1, 0, event.shiftKey);
    } else if (event.key === 'ArrowDown') {
      event.preventDefault();
      this.handleArrowKey(1, 0, event.shiftKey);
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      this.handleArrowKey(0, -1, event.shiftKey);
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      this.handleArrowKey(0, 1, event.shiftKey);
    } else if (event.key === 'Enter') {
      event.preventDefault();
      this.moveActiveCell(event.shiftKey ? -1 : 1, 0);
    } else if (event.key === 'Tab') {
      event.preventDefault();
      this.moveActiveCell(0, event.shiftKey ? -1 : 1);
    } else if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      if (this.state.selectedChartId()) {
        this.state.deleteChart(this.state.selectedChartId()!);
        return;
      }
      this.history.push(this.state.getSnapshot());
      this.state.clearSelection();
    } else if (event.key === 'F2') {
      event.preventDefault();
      this.startEditing(this.state.activeCell(), undefined, 'full');
    } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // Start typing directly into active cell
      event.preventDefault();
      this.startEditing(this.state.activeCell(), event.key, 'type');
    }
  }

  private handleArrowKey(dRow: number, dCol: number, isShift: boolean): void {
    const cur = parseAddress(this.state.activeCell());
    if (!cur) return;

    if (!isShift) {
      this.moveActiveCell(dRow, dCol);
    } else {
      const nextR = Math.max(0, Math.min(this.numRows - 1, this.selectionLead.row + dRow));
      const nextC = Math.max(0, Math.min(this.numCols - 1, this.selectionLead.col + dCol));
      this.extendSelectionTo(nextR, nextC);
    }
  }

  private moveActiveCell(dRow: number, dCol: number): void {
    const cur = parseAddress(this.state.activeCell());
    if (!cur) return;
    const r = Math.max(0, Math.min(this.numRows - 1, cur.row + dRow));
    const c = Math.max(0, Math.min(this.numCols - 1, cur.col + dCol));
    const nextAddr = formatAddress({ row: r, col: c });
    this.state.selectCell(nextAddr);
    this.selectionAnchor = { row: r, col: c };
    this.selectionLead = { row: r, col: c };
    this.scrollToActiveCell();
  }

  private getLastUsedCell(): string {
    let maxRow = 0;
    let maxCol = 0;

    for (const [addr, cell] of this.state.cells().entries()) {
      if (cell.raw !== undefined && cell.raw !== '' && cell.raw !== null) {
        const p = parseAddress(addr);
        if (p) {
          if (p.row > maxRow) maxRow = p.row;
          if (p.col > maxCol) maxCol = p.col;
        }
      }
    }

    for (const [addr, fmt] of this.state.cellFormats().entries()) {
      if (Object.keys(fmt).length > 0) {
        const p = parseAddress(addr);
        if (p) {
          if (p.row > maxRow) maxRow = p.row;
          if (p.col > maxCol) maxCol = p.col;
        }
      }
    }

    maxRow = Math.max(0, Math.min(this.numRows - 1, maxRow));
    maxCol = Math.max(0, Math.min(this.numCols - 1, maxCol));
    return formatAddress({ row: maxRow, col: maxCol });
  }

  private getVisibleRowCount(): number {
    const container = this.gridScrollRef?.nativeElement;
    if (!container || container.clientHeight <= 0) return 15;
    const availableHeight = container.clientHeight - this.rowHeight;
    return Math.max(1, Math.floor(availableHeight / this.rowHeight));
  }

  private findDataBlockEdge(startRow: number, startCol: number, dRow: number, dCol: number): CellCoord {
    const hasCellData = (r: number, c: number): boolean => {
      const addr = formatAddress({ row: r, col: c });
      const cell = this.state.getCell(addr);
      if (!cell) return false;
      const v = cell.value;
      const raw = cell.raw;
      return (v !== undefined && v !== null && v !== '') || (raw !== undefined && raw !== null && raw !== '');
    };

    const isCurrentCellFilled = hasCellData(startRow, startCol);
    let r = startRow;
    let c = startCol;

    const nextR = r + dRow;
    const nextC = c + dCol;

    if (nextR < 0 || nextR >= this.numRows || nextC < 0 || nextC >= this.numCols) {
      return { row: r, col: c };
    }

    const isNextCellFilled = hasCellData(nextR, nextC);

    if (isCurrentCellFilled) {
      if (isNextCellFilled) {
        while (
          r + dRow >= 0 &&
          r + dRow < this.numRows &&
          c + dCol >= 0 &&
          c + dCol < this.numCols &&
          hasCellData(r + dRow, c + dCol)
        ) {
          r += dRow;
          c += dCol;
        }
        return { row: r, col: c };
      } else {
        r += dRow;
        c += dCol;
        while (
          r >= 0 &&
          r < this.numRows &&
          c >= 0 &&
          c < this.numCols &&
          !hasCellData(r, c)
        ) {
          if (
            r + dRow < 0 ||
            r + dRow >= this.numRows ||
            c + dCol < 0 ||
            c + dCol >= this.numCols
          ) {
            return { row: r, col: c };
          }
          r += dRow;
          c += dCol;
        }
        return { row: r, col: c };
      }
    } else {
      while (
        r + dRow >= 0 &&
        r + dRow < this.numRows &&
        c + dCol >= 0 &&
        c + dCol < this.numCols
      ) {
        r += dRow;
        c += dCol;
        if (hasCellData(r, c)) {
          return { row: r, col: c };
        }
      }
      return { row: r, col: c };
    }
  }

  // Column header resize
  onResizeHandleMouseDown(colIndex: number, event: MouseEvent): void {
    event.stopPropagation();
    event.preventDefault();
    this.isResizingCol = true;
    this.resizeColIndex = colIndex;
    this.resizeStartX = event.clientX;
    const letter = indexToColLetter(colIndex);
    this.resizeStartWidth = this.getColWidth(letter);
    this.colResizeSnapshot = this.state.getSnapshot();
  }

  onResizeHandleDblClick(colIndex: number, event: MouseEvent): void {
    event.stopPropagation();
    event.preventDefault();
    this.autoFitColumn(colIndex);
  }

  autoFitColumn(colIndex: number): void {
    const letter = indexToColLetter(colIndex);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    let maxContentWidth = 40; // baseline minimum width

    // Check header width
    if (ctx) {
      ctx.font = '500 11px "Segoe UI", Arial, sans-serif';
      let headerWidth = ctx.measureText(letter).width + 16;
      if (this.isHeaderCellWithFilter(0, colIndex)) {
        headerWidth += 26; // filter button width + spacing
      }
      maxContentWidth = Math.max(maxContentWidth, headerWidth);
    }

    // Check all data rows in this column
    for (let r = 0; r < this.numRows; r++) {
      const cell = this.getCellData(r, colIndex);
      const text = cell?.displayValue !== undefined && cell?.displayValue !== null
        ? String(cell.displayValue)
        : (cell?.value !== undefined && cell?.value !== null && cell?.value !== '' ? String(cell.value) : '');
      if (!text) continue;

      const addr = formatAddress({ row: r, col: colIndex });
      const fmt = this.state.cellFormats().get(addr);
      const isBold = fmt?.bold ? 'bold ' : '';
      const isItalic = fmt?.italic ? 'italic ' : '';
      const fontSize = fmt?.fontSize ? `${fmt.fontSize}pt` : '10pt';
      const fontFamily = fmt?.fontFamily || '"Segoe UI", Arial, sans-serif';

      if (ctx) {
        ctx.font = `${isItalic}${isBold}${fontSize} ${fontFamily}`;
        let cellWidth = ctx.measureText(text).width + 18;
        if (cell?.comment) cellWidth += 6;
        if (this.hasValidation(r, colIndex)) cellWidth += 18;
        maxContentWidth = Math.max(maxContentWidth, cellWidth);
      }
    }

    // Also check DOM scrollWidth if elements are rendered
    const tableEl = this.gridScrollRef?.nativeElement;
    if (tableEl) {
      for (let r = 0; r < this.numRows; r++) {
        const addr = `${letter}${r + 1}`;
        const span = tableEl.querySelector<HTMLElement>(`td[data-cell-address="${addr}"] .cell-text`);
        if (span && span.scrollWidth > 0) {
          let cellWidth = span.scrollWidth + 16;
          if (this.hasValidation(r, colIndex)) cellWidth += 18;
          maxContentWidth = Math.max(maxContentWidth, cellWidth);
        }
      }
    }

    const newWidth = Math.max(20, Math.ceil(maxContentWidth));
    const oldWidth = this.getColWidth(letter);
    if (newWidth !== oldWidth) {
      this.history.push(this.state.getSnapshot());
      const map = new Map(this.state.columnWidths());
      map.set(letter, newWidth);
      this.state.columnWidths.set(map);
      if (this.hoveredCommentAddr()) {
        this.updateCommentNotePosition(this.hoveredCommentAddr()!);
      }
      this.cdr.markForCheck();
    }
  }

  // Row header resize
  onRowResizeHandleMouseDown(rowIndex: number, event: MouseEvent): void {
    event.stopPropagation();
    event.preventDefault();
    this.isResizingRow = true;
    this.resizeRowIndex = rowIndex;
    this.resizeStartY = event.clientY;
    this.resizeStartHeight = this.getRowHeight(rowIndex);
    this.rowResizeSnapshot = this.state.getSnapshot();
  }

  onRowResizeHandleDblClick(rowIndex: number, event: MouseEvent): void {
    event.stopPropagation();
    event.preventDefault();
    const curHeight = this.getRowHeight(rowIndex);
    if (curHeight !== this.defaultRowHeight) {
      this.history.push(this.state.getSnapshot());
      const map = new Map(this.state.rowHeights());
      map.delete(rowIndex);
      this.state.rowHeights.set(map);
      this.cdr.markForCheck();
    }
  }

  @HostListener('document:mousemove', ['$event'])
  onDocumentMouseMove(event: MouseEvent): void {
    if (this.isResizingCol && this.resizeColIndex >= 0) {
      event.preventDefault();
      const deltaX = event.clientX - this.resizeStartX;
      const newWidth = Math.max(20, this.resizeStartWidth + deltaX);
      const letter = indexToColLetter(this.resizeColIndex);
      const map = new Map(this.state.columnWidths());
      map.set(letter, newWidth);
      this.state.columnWidths.set(map);
      if (this.hoveredCommentAddr()) {
        this.updateCommentNotePosition(this.hoveredCommentAddr()!);
      }
      this.cdr.markForCheck();
    }

    if (this.isResizingRow && this.resizeRowIndex >= 0) {
      event.preventDefault();
      const deltaY = event.clientY - this.resizeStartY;
      const newHeight = Math.max(12, this.resizeStartHeight + deltaY);
      const map = new Map(this.state.rowHeights());
      map.set(this.resizeRowIndex, newHeight);
      this.state.rowHeights.set(map);
      this.cdr.markForCheck();
    }

    if (this.isDraggingChart && this.draggedChartId) {
      const deltaX = event.clientX - this.chartDragStartX;
      const deltaY = event.clientY - this.chartDragStartY;
      const chart = this.state.charts().find((c) => c.id === this.draggedChartId);
      if (chart) {
        chart.x = Math.max(10, this.chartInitialX + deltaX);
        chart.y = Math.max(10, this.chartInitialY + deltaY);
        this.cdr.markForCheck();
      }
    }
  }

  // Filter button click
  onFilterButtonClick(colIndex: number, event: MouseEvent): void {
    event.stopPropagation();
    const btn = event.currentTarget as HTMLElement;
    const rect = btn.getBoundingClientRect();
    this.openFilterDropdown.emit({
      colIndex,
      anchorX: rect.left,
      anchorY: rect.bottom + 2,
    });
  }

  // Validation dropdown arrow
  onValidationArrowClick(row: number, col: number, event: MouseEvent): void {
    event.stopPropagation();
    const addr = formatAddress({ row, col });
    const dv = this.state.getDataValidation(addr);
    if (!dv) return;

    const btn = event.currentTarget as HTMLElement;
    const rect = btn.getBoundingClientRect();
    this.activeValidationMenu.set({
      cell: addr,
      allowed: dv.allowed,
      x: rect.left,
      y: rect.bottom + 2,
    });
  }

  selectValidationOption(val: string): void {
    const cur = this.activeValidationMenu();
    if (!cur) return;
    this.history.push(this.state.getSnapshot());
    this.state.setCellValue(cur.cell, val);
    this.activeValidationMenu.set(null);
  }

  // Charts interaction
  onChartMouseDown(chart: ChartItem, event: MouseEvent): void {
    event.stopPropagation();
    this.state.selectedChartId.set(chart.id);
    this.isDraggingChart = true;
    this.draggedChartId = chart.id;
    this.chartDragStartX = event.clientX;
    this.chartDragStartY = event.clientY;
    this.chartInitialX = chart.x;
    this.chartInitialY = chart.y;
  }

  readonly chartColors: string[] = [
    '#2b579a', // Office blue
    '#ed7d31', // Office orange
    '#107c41', // Excel green
    '#ffc000', // Excel yellow
    '#5c2d91', // Office purple
    '#0078d4', // Office cyan/blue
    '#b4009e', // Office magenta
    '#70ad47', // Olive green
  ];

  truncateText(text: string, maxLen: number): string {
    if (!text) return '';
    return text.length > maxLen ? text.slice(0, maxLen - 1) + '…' : text;
  }

  private parseCellNumber(val: any): number | null {
    if (val === null || val === undefined || val === '') return null;
    if (typeof val === 'number') return isNaN(val) ? null : val;
    const s = String(val).trim();
    if (s === '') return null;
    if (['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A'].includes(s)) return null;
    const n = Number(s);
    return isNaN(n) ? null : n;
  }

  getChartRenderData(chart: ChartItem, _cells: Map<string, CellData>): ChartRenderModel {
    const parsed = parseRange(chart.range);
    if (!parsed) {
      return this.createEmptyChartModel(chart.chartType);
    }

    const { startRow, startCol, endRow, endCol } = parsed;

    const getCellInfo = (r: number, c: number) => {
      const addr = formatAddress({ row: r, col: c });
      const cell = this.state.getCell(addr);
      const rawVal = cell?.value ?? '';
      const str = cell?.displayValue !== undefined ? String(cell.displayValue) : String(rawVal);
      const num = this.parseCellNumber(rawVal);
      return { str, num, val: rawVal };
    };

    const isSingleCol = startCol === endCol;
    const catCol = startCol;
    const seriesCols: number[] = [];
    if (isSingleCol) {
      seriesCols.push(startCol);
    } else {
      for (let c = startCol + 1; c <= endCol; c++) {
        seriesCols.push(c);
      }
    }

    // Header detection:
    // "if the first row is text above numbers, use it as the series names (header)."
    let hasHeader = false;
    if (startRow < endRow) {
      for (const sc of seriesCols) {
        const topInfo = getCellInfo(startRow, sc);
        if (topInfo.num === null) {
          for (let r = startRow + 1; r <= endRow; r++) {
            if (getCellInfo(r, sc).num !== null) {
              hasHeader = true;
              break;
            }
          }
        }
        if (hasHeader) break;
      }
    }

    const dataStartRow = hasHeader ? startRow + 1 : startRow;

    // Categories: one per data row (do not merge duplicate labels)
    const categories: string[] = [];
    for (let r = dataStartRow; r <= endRow; r++) {
      if (isSingleCol) {
        categories.push(String(r - dataStartRow + 1));
      } else {
        const cInfo = getCellInfo(r, catCol);
        categories.push(cInfo.str.trim() !== '' ? cInfo.str : `Row ${r + 1}`);
      }
    }

    // Series:
    const seriesList: ChartSeriesData[] = [];
    for (let sIdx = 0; sIdx < seriesCols.length; sIdx++) {
      const sc = seriesCols[sIdx];
      let seriesName = '';
      if (hasHeader) {
        const headerInfo = getCellInfo(startRow, sc);
        seriesName = headerInfo.str.trim() || `Series ${sIdx + 1}`;
      } else {
        seriesName = `Series ${sIdx + 1}`;
      }

      const values: number[] = [];
      for (let r = dataStartRow; r <= endRow; r++) {
        const rowInfo = getCellInfo(r, sc);
        values.push(rowInfo.num !== null ? rowInfo.num : 0);
      }
      seriesList.push({ name: seriesName, values });
    }

    const hasMultipleSeries = seriesList.length > 1;

    if (chart.chartType === 'pie') {
      return this.buildPieChartModel(chart.chartType, categories, seriesList);
    } else if (chart.chartType === 'bar') {
      return this.buildBarChartModel(chart.chartType, categories, seriesList, hasMultipleSeries);
    } else if (chart.chartType === 'line') {
      return this.buildLineChartModel(chart.chartType, categories, seriesList, hasMultipleSeries);
    } else {
      return this.buildColumnChartModel(chart.chartType, categories, seriesList, hasMultipleSeries);
    }
  }

  private computeNiceScale(allValues: number[]): { axisMin: number; axisMax: number; ticks: Array<{ val: number; label: string }> } {
    let minVal = 0;
    let maxVal = 0;
    if (allValues.length > 0) {
      minVal = Math.min(...allValues);
      maxVal = Math.max(...allValues);
    }

    const rangeMin = Math.min(0, minVal);
    const rangeMax = Math.max(0, maxVal);

    let rawSpan = rangeMax - rangeMin;
    if (rawSpan === 0) rawSpan = 10;

    const rawStep = rawSpan / 4;
    const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
    const frac = rawStep / mag;
    let niceFrac = 1;
    if (frac <= 1) niceFrac = 1;
    else if (frac <= 2) niceFrac = 2;
    else if (frac <= 2.5) niceFrac = 2.5;
    else if (frac <= 5) niceFrac = 5;
    else niceFrac = 10;

    const niceStep = niceFrac * mag;
    const axisMin = Math.floor(rangeMin / niceStep) * niceStep;
    const axisMax = Math.ceil(rangeMax / niceStep) * niceStep;

    const count = Math.max(1, Math.round((axisMax - axisMin) / niceStep));
    const ticks: Array<{ val: number; label: string }> = [];

    for (let i = 0; i <= count; i++) {
      const val = Math.round((axisMin + i * niceStep) * 1e6) / 1e6;
      let label: string;
      if (Math.abs(val) >= 1000000) {
        label = val.toLocaleString();
      } else if (Math.abs(val) >= 10000) {
        label = String(val);
      } else if (Number.isInteger(val)) {
        label = String(val);
      } else {
        label = String(Number(val.toFixed(2)));
      }
      ticks.push({ val, label });
    }

    return { axisMin, axisMax, ticks };
  }

  private buildLegendItems(seriesList: ChartSeriesData[], plotLeft: number, plotRight: number): ChartLegendItem[] {
    const items: ChartLegendItem[] = [];
    const S = seriesList.length;
    if (S <= 1) return items;

    const totalWidth = plotRight - plotLeft;
    const itemWidth = Math.min(80, totalWidth / S);
    for (let s = 0; s < S; s++) {
      items.push({
        name: this.truncateText(seriesList[s].name, 9),
        color: this.chartColors[s % this.chartColors.length],
        x: plotLeft + s * itemWidth,
      });
    }
    return items;
  }

  private buildColumnChartModel(
    chartType: 'column',
    categories: string[],
    seriesList: ChartSeriesData[],
    hasMultipleSeries: boolean
  ): ChartRenderModel {
    const allVals = seriesList.flatMap((s) => s.values);
    const { axisMin, axisMax, ticks } = this.computeNiceScale(allVals);

    const maxLabelLen = Math.max(0, ...ticks.map((t) => t.label.length));
    const plotLeft = maxLabelLen >= 5 ? 46 : 38;
    const plotRight = 290;
    const plotTop = hasMultipleSeries ? 22 : 12;
    const plotBottom = 142;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const span = axisMax - axisMin || 1;
    const getY = (val: number) => plotBottom - ((val - axisMin) / span) * plotHeight;
    const zeroY = getY(0);

    const gridlines: ChartGridLine[] = [];
    const chartTicks: ChartTick[] = [];
    for (const t of ticks) {
      const y = getY(t.val);
      gridlines.push({ x1: plotLeft, y1: y, x2: plotRight, y2: y });
      chartTicks.push({ val: t.val, label: t.label, pos: y, x: plotLeft - 4, y: y + 3 });
    }

    let zeroLine: ChartGridLine | undefined;
    if (axisMin < 0 && axisMax > 0) {
      zeroLine = { x1: plotLeft, y1: zeroY, x2: plotRight, y2: zeroY };
    }

    const N = categories.length;
    const S = seriesList.length;
    const catWidth = plotWidth / Math.max(1, N);
    const categoryLabels: ChartCategoryLabel[] = [];
    const bars: ChartBarRect[] = [];

    for (let i = 0; i < N; i++) {
      const centerX = plotLeft + (i + 0.5) * catWidth;
      categoryLabels.push({
        text: this.truncateText(categories[i], 8),
        x: centerX,
        y: 156,
        anchor: 'middle',
      });

      const usableW = catWidth * 0.75;
      const barW = Math.max(2, Math.min(28, usableW / Math.max(1, S)));
      const groupW = barW * S;
      const startX = centerX - groupW / 2;

      for (let s = 0; s < S; s++) {
        const val = seriesList[s].values[i] ?? 0;
        const valY = getY(val);
        let topY = valY;
        let height = 0;

        if (val >= 0) {
          topY = valY;
          height = zeroY - valY;
        } else {
          topY = zeroY;
          height = valY - zeroY;
        }

        if (val !== 0 && height < 1) height = 1;

        const rectW = Math.max(1, barW > 3 ? barW - 1 : barW);
        bars.push({
          x: startX + s * barW,
          y: topY,
          width: rectW,
          height,
          fill: this.chartColors[s % this.chartColors.length],
        });
      }
    }

    return {
      chartType,
      hasMultipleSeries,
      legendItems: this.buildLegendItems(seriesList, plotLeft, plotRight),
      ticks: chartTicks,
      gridlines,
      zeroLine,
      categoryLabels,
      bars,
      lineSeries: [],
      pieSlices: [],
    };
  }

  private buildBarChartModel(
    chartType: 'bar',
    categories: string[],
    seriesList: ChartSeriesData[],
    hasMultipleSeries: boolean
  ): ChartRenderModel {
    const allVals = seriesList.flatMap((s) => s.values);
    const { axisMin, axisMax, ticks } = this.computeNiceScale(allVals);

    const plotLeft = 60;
    const plotRight = 290;
    const plotTop = hasMultipleSeries ? 22 : 12;
    const plotBottom = 142;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const span = axisMax - axisMin || 1;
    const getX = (val: number) => plotLeft + ((val - axisMin) / span) * plotWidth;
    const zeroX = getX(0);

    const gridlines: ChartGridLine[] = [];
    const chartTicks: ChartTick[] = [];
    for (const t of ticks) {
      const x = getX(t.val);
      gridlines.push({ x1: x, y1: plotTop, x2: x, y2: plotBottom });
      chartTicks.push({ val: t.val, label: t.label, pos: x, x, y: 156 });
    }

    let zeroLine: ChartGridLine | undefined;
    if (axisMin < 0 && axisMax > 0) {
      zeroLine = { x1: zeroX, y1: plotTop, x2: zeroX, y2: plotBottom };
    }

    const N = categories.length;
    const S = seriesList.length;
    const catHeight = plotHeight / Math.max(1, N);
    const categoryLabels: ChartCategoryLabel[] = [];
    const bars: ChartBarRect[] = [];

    for (let i = 0; i < N; i++) {
      const centerY = plotTop + (i + 0.5) * catHeight;
      categoryLabels.push({
        text: this.truncateText(categories[i], 8),
        x: plotLeft - 5,
        y: centerY + 3,
        anchor: 'end',
      });

      const usableH = catHeight * 0.75;
      const barH = Math.max(2, Math.min(22, usableH / Math.max(1, S)));
      const groupH = barH * S;
      const startY = centerY - groupH / 2;

      for (let s = 0; s < S; s++) {
        const val = seriesList[s].values[i] ?? 0;
        const valX = getX(val);
        let startBarX = zeroX;
        let width = 0;

        if (val >= 0) {
          startBarX = zeroX;
          width = valX - zeroX;
        } else {
          startBarX = valX;
          width = zeroX - valX;
        }

        if (val !== 0 && width < 1) width = 1;

        const rectH = Math.max(1, barH > 3 ? barH - 1 : barH);
        bars.push({
          x: startBarX,
          y: startY + s * barH,
          width,
          height: rectH,
          fill: this.chartColors[s % this.chartColors.length],
        });
      }
    }

    return {
      chartType,
      hasMultipleSeries,
      legendItems: this.buildLegendItems(seriesList, plotLeft, plotRight),
      ticks: chartTicks,
      gridlines,
      zeroLine,
      categoryLabels,
      bars,
      lineSeries: [],
      pieSlices: [],
    };
  }

  private buildLineChartModel(
    chartType: 'line',
    categories: string[],
    seriesList: ChartSeriesData[],
    hasMultipleSeries: boolean
  ): ChartRenderModel {
    const allVals = seriesList.flatMap((s) => s.values);
    const { axisMin, axisMax, ticks } = this.computeNiceScale(allVals);

    const maxLabelLen = Math.max(0, ...ticks.map((t) => t.label.length));
    const plotLeft = maxLabelLen >= 5 ? 46 : 38;
    const plotRight = 290;
    const plotTop = hasMultipleSeries ? 22 : 12;
    const plotBottom = 142;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const span = axisMax - axisMin || 1;
    const getY = (val: number) => plotBottom - ((val - axisMin) / span) * plotHeight;
    const zeroY = getY(0);

    const gridlines: ChartGridLine[] = [];
    const chartTicks: ChartTick[] = [];
    for (const t of ticks) {
      const y = getY(t.val);
      gridlines.push({ x1: plotLeft, y1: y, x2: plotRight, y2: y });
      chartTicks.push({ val: t.val, label: t.label, pos: y, x: plotLeft - 4, y: y + 3 });
    }

    let zeroLine: ChartGridLine | undefined;
    if (axisMin < 0 && axisMax > 0) {
      zeroLine = { x1: plotLeft, y1: zeroY, x2: plotRight, y2: zeroY };
    }

    const N = categories.length;
    const catWidth = plotWidth / Math.max(1, N);
    const categoryLabels: ChartCategoryLabel[] = [];

    for (let i = 0; i < N; i++) {
      const centerX = plotLeft + (i + 0.5) * catWidth;
      categoryLabels.push({
        text: this.truncateText(categories[i], 8),
        x: centerX,
        y: 156,
        anchor: 'middle',
      });
    }

    const lineSeries: ChartLineSeries[] = [];
    for (let s = 0; s < seriesList.length; s++) {
      const points: Array<{ x: number; y: number }> = [];
      for (let i = 0; i < N; i++) {
        const centerX = plotLeft + (i + 0.5) * catWidth;
        const val = seriesList[s].values[i] ?? 0;
        const y = getY(val);
        points.push({ x: centerX, y });
      }
      const polylinePoints = points.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
      lineSeries.push({
        name: seriesList[s].name,
        color: this.chartColors[s % this.chartColors.length],
        polylinePoints,
        points,
      });
    }

    return {
      chartType,
      hasMultipleSeries,
      legendItems: this.buildLegendItems(seriesList, plotLeft, plotRight),
      ticks: chartTicks,
      gridlines,
      zeroLine,
      categoryLabels,
      bars: [],
      lineSeries,
      pieSlices: [],
    };
  }

  private buildPieChartModel(
    chartType: 'pie',
    categories: string[],
    seriesList: ChartSeriesData[]
  ): ChartRenderModel {
    const firstSeries = seriesList[0];
    const items: Array<{ label: string; val: number }> = [];

    if (firstSeries) {
      for (let i = 0; i < categories.length; i++) {
        const val = firstSeries.values[i] ?? 0;
        if (val > 0) {
          items.push({ label: categories[i], val });
        }
      }
    }

    if (items.length === 0) {
      return {
        chartType,
        hasMultipleSeries: false,
        legendItems: [],
        ticks: [],
        gridlines: [],
        categoryLabels: [],
        bars: [],
        lineSeries: [],
        pieSlices: [],
        isPieEmpty: true,
      };
    }

    const total = items.reduce((acc, it) => acc + it.val, 0);
    const cx = 82;
    const cy = 85;
    const r = 54;
    const pieSlices: ChartPieSlice[] = [];

    if (items.length === 1) {
      pieSlices.push({
        label: items[0].label,
        val: items[0].val,
        percentage: 100,
        percentageLabel: '100%',
        isFullCircle: true,
        color: this.chartColors[0],
      });
    } else {
      let curAngle = -Math.PI / 2;
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const share = it.val / total;
        const angle = share * 2 * Math.PI;
        const startAngle = curAngle;
        const endAngle = curAngle + angle;
        curAngle = endAngle;

        const x1 = cx + r * Math.cos(startAngle);
        const y1 = cy + r * Math.sin(startAngle);
        const x2 = cx + r * Math.cos(endAngle);
        const y2 = cy + r * Math.sin(endAngle);
        const largeArc = angle > Math.PI ? 1 : 0;

        const pathD = `M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
        const pct = Math.round(share * 100);

        pieSlices.push({
          label: it.label,
          val: it.val,
          percentage: pct,
          percentageLabel: `${pct}%`,
          pathD,
          isFullCircle: false,
          color: this.chartColors[i % this.chartColors.length],
        });
      }
    }

    return {
      chartType,
      hasMultipleSeries: false,
      legendItems: [],
      ticks: [],
      gridlines: [],
      categoryLabels: [],
      bars: [],
      lineSeries: [],
      pieSlices,
      pieCenter: { cx, cy, r },
      isPieEmpty: false,
    };
  }

  private createEmptyChartModel(chartType: 'column' | 'bar' | 'line' | 'pie'): ChartRenderModel {
    return {
      chartType,
      hasMultipleSeries: false,
      legendItems: [],
      ticks: [],
      gridlines: [],
      categoryLabels: [],
      bars: [],
      lineSeries: [],
      pieSlices: [],
      isPieEmpty: chartType === 'pie',
    };
  }

  updateCommentNotePosition(addr: string, targetCellEl?: HTMLElement): void {
    const text = this.state.comments().get(addr);
    if (!text) {
      this.hoveredComment.set(null);
      return;
    }

    const container = this.gridScrollRef?.nativeElement;
    if (!container) return;

    let cellEl = targetCellEl;
    if (!cellEl) {
      cellEl = container.querySelector<HTMLElement>(`[data-cell-address="${addr}"]`) ?? undefined;
    }
    if (!cellEl) return;

    const containerRect = container.getBoundingClientRect();
    const cellRect = cellEl.getBoundingClientRect();

    const headerTop = 24;
    const headerLeft = 44;

    // Hide if cell is scrolled completely out of view or under sticky headers
    if (
      cellRect.bottom < containerRect.top + headerTop ||
      cellRect.right < containerRect.left + headerLeft ||
      cellRect.top > containerRect.bottom ||
      cellRect.left > containerRect.right
    ) {
      this.hoveredComment.set(null);
      return;
    }

    const cellTop = (cellRect.top - containerRect.top) + container.scrollTop;
    const cellLeft = (cellRect.left - containerRect.left) + container.scrollLeft;
    const cellWidth = cellRect.width;
    const cellRight = cellLeft + cellWidth;

    const noteWidth = 180;
    const gap = 14;

    const wouldOverflowRight = (cellRect.right + gap + noteWidth) > (containerRect.right - 4);

    let boxLeft: number;
    let fromX: number;
    let fromY: number;
    const boxTop = Math.max(container.scrollTop + headerTop + 2, cellTop - 4);
    const toX = cellRight;
    const toY = cellTop;

    if (wouldOverflowRight) {
      boxLeft = Math.max(container.scrollLeft + headerLeft + 4, cellLeft - gap - noteWidth);
      fromX = boxLeft + noteWidth;
      fromY = boxTop;
    } else {
      boxLeft = cellRight + gap;
      fromX = boxLeft;
      fromY = boxTop;
    }

    const minX = Math.min(fromX, toX) - 4;
    const minY = Math.min(fromY, toY) - 4;
    const maxX = Math.max(fromX, toX) + 4;
    const maxY = Math.max(fromY, toY) + 4;

    const svgLeft = minX;
    const svgTop = minY;
    const svgWidth = Math.max(maxX - minX, 10);
    const svgHeight = Math.max(maxY - minY, 10);

    const relFromX = fromX - svgLeft;
    const relFromY = fromY - svgTop;
    const relToX = toX - svgLeft;
    const relToY = toY - svgTop;

    const arrowPoints = this.calculateArrowheadPoints(relFromX, relFromY, relToX, relToY);

    this.hoveredComment.set({
      address: addr,
      text,
      boxLeft,
      boxTop,
      svgLeft,
      svgTop,
      svgWidth,
      svgHeight,
      lineFromX: relFromX,
      lineFromY: relFromY,
      lineToX: relToX,
      lineToY: relToY,
      arrowPoints,
    });
    this.cdr.markForCheck();
  }

  private calculateArrowheadPoints(fromX: number, fromY: number, toX: number, toY: number): string {
    const angle = Math.atan2(toY - fromY, toX - fromX);
    const length = 5;
    const width = 2.5;

    const baseX = toX - length * Math.cos(angle);
    const baseY = toY - length * Math.sin(angle);

    const p1X = baseX + width * Math.sin(angle);
    const p1Y = baseY - width * Math.cos(angle);

    const p2X = baseX - width * Math.sin(angle);
    const p2Y = baseY + width * Math.cos(angle);

    return `${toX.toFixed(1)},${toY.toFixed(1)} ${p1X.toFixed(1)},${p1Y.toFixed(1)} ${p2X.toFixed(1)},${p2Y.toFixed(1)}`;
  }
}

export interface HoveredCommentNote {
  address: string;
  text: string;
  boxLeft: number;
  boxTop: number;
  svgLeft: number;
  svgTop: number;
  svgWidth: number;
  svgHeight: number;
  lineFromX: number;
  lineFromY: number;
  lineToX: number;
  lineToY: number;
  arrowPoints: string;
}

export interface ChartSeriesData {
  name: string;
  values: number[];
}

export interface ChartTick {
  val: number;
  label: string;
  pos: number;
  x: number;
  y: number;
}

export interface ChartBarRect {
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}

export interface ChartLineSeries {
  name: string;
  color: string;
  polylinePoints: string;
  points: Array<{ x: number; y: number }>;
}

export interface ChartPieSlice {
  label: string;
  val: number;
  percentage: number;
  percentageLabel: string;
  pathD?: string;
  isFullCircle?: boolean;
  color: string;
}

export interface ChartLegendItem {
  name: string;
  color: string;
  x: number;
}

export interface ChartCategoryLabel {
  text: string;
  x: number;
  y: number;
  anchor: string;
}

export interface ChartGridLine {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface ChartRenderModel {
  chartType: 'column' | 'bar' | 'line' | 'pie';
  hasMultipleSeries: boolean;
  legendItems: ChartLegendItem[];
  ticks: ChartTick[];
  gridlines: ChartGridLine[];
  zeroLine?: ChartGridLine;
  categoryLabels: ChartCategoryLabel[];
  bars: ChartBarRect[];
  lineSeries: ChartLineSeries[];
  pieSlices: ChartPieSlice[];
  pieCenter?: { cx: number; cy: number; r: number };
  isPieEmpty?: boolean;
}
