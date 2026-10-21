import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-excel-icon',
  standalone: true,
  imports: [CommonModule],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <span class="excel-icon-wrapper" [style.width.px]="size" [style.height.px]="size" [style.color]="color" aria-hidden="true">
      <svg
        [attr.width]="size"
        [attr.height]="size"
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        class="excel-svg"
      >
        <!-- SEARCH -->
        <g *ngIf="name === 'search'">
          <circle cx="10" cy="10" r="6" stroke="currentColor" stroke-width="2" fill="none"/>
          <line x1="14.5" y1="14.5" x2="20" y2="20" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        </g>

        <!-- DATE -->
        <g *ngIf="name === 'date'">
          <rect x="3" y="4" width="18" height="17" rx="2" stroke="currentColor" stroke-width="1.8" fill="none"/>
          <line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" stroke-width="1.6"/>
          <line x1="8" y1="2" x2="8" y2="5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
          <line x1="16" y1="2" x2="16" y2="5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        </g>

        <!-- UNDO / REDO -->
        <g *ngIf="name === 'undo'">
          <path d="M10 4L6 8l4 4" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M6 8h6a7 7 0 0 1 7 7v4" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'redo'">
          <path d="M14 4l4 4-4 4" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M18 8h-6a7 7 0 0 0-7 7v4" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <!-- CLIPBOARD -->
        <g *ngIf="name === 'paste'">
          <rect x="4" y="2" width="13" height="19" rx="1.5" fill="#f59e0b" stroke="#d97706" stroke-width="1.2"/>
          <rect x="7" y="0.5" width="7" height="3.5" rx="0.8" fill="#e2e8f0" stroke="#64748b" stroke-width="1"/>
          <rect x="8" y="5.5" width="13" height="16.5" rx="1.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.5"/>
          <line x1="11" y1="10" x2="18" y2="10" stroke="#3b82f6" stroke-width="1.6" stroke-linecap="round"/>
          <line x1="11" y1="13.5" x2="18" y2="13.5" stroke="#3b82f6" stroke-width="1.6" stroke-linecap="round"/>
          <line x1="11" y1="17" x2="16" y2="17" stroke="#3b82f6" stroke-width="1.6" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'cut'">
          <circle cx="5.5" cy="5.5" r="3" stroke="#475569" stroke-width="1.6" fill="none"/>
          <circle cx="5.5" cy="18.5" r="3" stroke="#475569" stroke-width="1.6" fill="none"/>
          <path d="M7.8 7.8L20 18.5M7.8 16.2L20 5.5" stroke="#334155" stroke-width="1.8" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'copy'">
          <rect x="2" y="5" width="13" height="16" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.5"/>
          <path d="M7 5V2.5A1.5 1.5 0 0 1 8.5 1h10A1.5 1.5 0 0 1 20 2.5v13a1.5 1.5 0 0 1-1.5 1.5H15" stroke="#2563eb" stroke-width="1.6"/>
          <line x1="5" y1="9.5" x2="12" y2="9.5" stroke="#94a3b8" stroke-width="1.3"/>
          <line x1="5" y1="13" x2="12" y2="13" stroke="#94a3b8" stroke-width="1.3"/>
          <line x1="5" y1="16.5" x2="10" y2="16.5" stroke="#94a3b8" stroke-width="1.3"/>
        </g>

        <g *ngIf="name === 'formatPainter'">
          <path d="M3 3a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v4H3V3z" fill="#f59e0b"/>
          <rect x="2" y="7" width="16" height="3.5" rx="0.5" fill="#3b82f6"/>
          <path d="M8 10.5v8a2 2 0 0 0 2 2h0a2 2 0 0 0 2-2v-8H8z" fill="#94a3b8"/>
        </g>

        <g *ngIf="name === 'bold'">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M5 2.5h8a4.5 4.5 0 0 1 3.2 7.7A5 5 0 0 1 13.5 21.5H5V2.5zm3.5 3.5h4a2 2 0 0 1 0 4h-4V6zm0 6h4.5a2.2 2.2 0 0 1 0 4.4H8.5V12z"
            fill="currentColor"
          />
        </g>

        <g *ngIf="name === 'italic'">
          <path d="M9 2.5h9M6 21.5h9M15 2.5l-6 19" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'underline'">
          <path d="M6 2.5v9a6 6 0 0 0 12 0v-9" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="4" y1="21.5" x2="20" y2="21.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'borders'">
          <rect x="2.5" y="2.5" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.5"/>
          <line x1="12" y1="2.5" x2="12" y2="21.5" stroke="currentColor" stroke-width="1.2"/>
          <line x1="2.5" y1="12" x2="21.5" y2="12" stroke="currentColor" stroke-width="1.2"/>
          <line x1="2.5" y1="21.5" x2="21.5" y2="21.5" stroke="#2563eb" stroke-width="2.8"/>
        </g>

        <g *ngIf="name === 'borderBottom'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <line x1="3" y1="21" x2="21" y2="21" stroke="currentColor" stroke-width="2.6"/>
        </g>

        <g *ngIf="name === 'borderTop'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <line x1="3" y1="3" x2="21" y2="3" stroke="currentColor" stroke-width="2.6"/>
        </g>

        <g *ngIf="name === 'borderLeft'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <line x1="3" y1="3" x2="3" y2="21" stroke="currentColor" stroke-width="2.6"/>
        </g>

        <g *ngIf="name === 'borderRight'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <line x1="21" y1="3" x2="21" y2="21" stroke="currentColor" stroke-width="2.6"/>
        </g>

        <g *ngIf="name === 'borderNone'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
          <line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
        </g>

        <g *ngIf="name === 'borderAll'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.8" fill="none"/>
          <line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" stroke-width="1.8"/>
          <line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.8"/>
        </g>

        <g *ngIf="name === 'borderOutside'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="1.8" fill="none"/>
          <line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
          <line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
        </g>

        <g *ngIf="name === 'borderThickOutside'">
          <rect x="3" y="3" width="18" height="18" stroke="currentColor" stroke-width="3" fill="none"/>
          <line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
          <line x1="3" y1="12" x2="21" y2="12" stroke="currentColor" stroke-width="1.2" stroke-dasharray="2 2"/>
        </g>

        <g *ngIf="name === 'fillColor'">
          <path d="M16 2.5l4 4-10 10-5-5 11-9z" fill="#f59e0b" stroke="#475569" stroke-width="1.2"/>
          <path d="M4.5 12l-2 2a2 2 0 0 0 0 2.8l2.5 2.5a2 2 0 0 0 2.8 0l2-2-5.3-5.3z" fill="#94a3b8"/>
          <path d="M21 14.5a2.2 2.2 0 1 1-4.4 0c0-1.6 2.2-4.5 2.2-4.5s2.2 2.9 2.2 4.5z" fill="#eab308"/>
          <rect x="1" y="20.5" width="22" height="3.5" fill="#eab308"/>
        </g>

        <g *ngIf="name === 'fontColor'">
          <path d="M5 18L12 2l7 16M7.8 13h8.4" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
          <rect x="1" y="20.5" width="22" height="3.5" fill="#dc2626"/>
        </g>

        <g *ngIf="name === 'increaseFont'">
          <path d="M3 19L9 3l6 16M5.5 14h7" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M16 9l3.5-5 3.5 5M19.5 5v10" stroke="#16a34a" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'decreaseFont'">
          <path d="M3 19L8 6l5 13M5 15h6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M16 10l3.5 5 3.5-5M19.5 14V4" stroke="#dc2626" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <!-- ALIGNMENT -->
        <g *ngIf="name === 'alignTop'">
          <line x1="2" y1="3" x2="22" y2="3" stroke="#2563eb" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="4" y1="8" x2="20" y2="8" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="4" y1="13" x2="16" y2="13" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'alignMiddle'">
          <line x1="2" y1="12" x2="22" y2="12" stroke="#2563eb" stroke-width="1.8" stroke-dasharray="2 2"/>
          <line x1="4" y1="7" x2="20" y2="7" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="4" y1="17" x2="16" y2="17" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'alignBottom'">
          <line x1="2" y1="21" x2="22" y2="21" stroke="#2563eb" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="4" y1="11" x2="16" y2="11" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="4" y1="16" x2="20" y2="16" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'orientation'">
          <text x="1" y="19" font-family="Segoe UI, sans-serif" font-size="14" font-weight="bold" fill="currentColor">ab</text>
          <path d="M14 16l6-6M15 10h5v5" stroke="#2563eb" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'wrapText'">
          <text x="1" y="10" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">ab</text>
          <path d="M11 6h7a3 3 0 0 1 3 3v0a3 3 0 0 1-3 3h-7m0 0l2.5-2.5m-2.5 2.5l2.5 2.5" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
          <text x="1" y="21" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">c</text>
        </g>

        <g *ngIf="name === 'alignLeft'">
          <line x1="3" y1="5" x2="21" y2="5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="3" y1="11" x2="14" y2="11" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="3" y1="17" x2="19" y2="17" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'alignCenter'">
          <line x1="3" y1="5" x2="21" y2="5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="6.5" y1="11" x2="17.5" y2="11" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="4" y1="17" x2="20" y2="17" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'alignRight'">
          <line x1="3" y1="5" x2="21" y2="5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="10" y1="11" x2="21" y2="11" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
          <line x1="5" y1="17" x2="21" y2="17" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'decreaseIndent'">
          <path d="M8 7l-5 5 5 5" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <line x1="10" y1="6" x2="21" y2="6" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="10" y1="12" x2="18" y2="12" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="10" y1="18" x2="21" y2="18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'increaseIndent'">
          <path d="M3 7l5 5-5 5" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <line x1="10" y1="6" x2="21" y2="6" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="10" y1="12" x2="18" y2="12" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="10" y1="18" x2="21" y2="18" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'mergeCenter'">
          <rect x="2" y="3.5" width="20" height="17" rx="2" stroke="currentColor" stroke-width="1.6" fill="none"/>
          <path d="M7 12h10m-6-3.5L7 12l4 3.5m2-7l4 3.5-4 3.5" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <!-- NUMBER -->
        <g *ngIf="name === 'accounting'">
          <text x="6" y="18" font-family="Segoe UI, sans-serif" font-size="19" font-weight="bold" fill="currentColor">$</text>
        </g>

        <g *ngIf="name === 'percent'">
          <text x="4" y="18" font-family="Segoe UI, sans-serif" font-size="19" font-weight="bold" fill="currentColor">%</text>
        </g>

        <g *ngIf="name === 'comma'">
          <text x="8" y="19" font-family="Segoe UI, sans-serif" font-size="24" font-weight="bold" fill="currentColor">,</text>
        </g>

        <g *ngIf="name === 'increaseDecimal'">
          <text x="1" y="16" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">.0</text>
          <path d="M12 12l-3 3 3 3m-2-3h7" stroke="#2563eb" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
          <text x="15" y="16" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">00</text>
        </g>

        <g *ngIf="name === 'decreaseDecimal'">
          <text x="1" y="16" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">.00</text>
          <path d="M14 12l3 3-3 3m2-3H8" stroke="#2563eb" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
          <text x="18" y="16" font-family="Segoe UI, sans-serif" font-size="10" font-weight="bold" fill="currentColor">0</text>
        </g>

        <!-- STYLES -->
        <g *ngIf="name === 'conditionalFormatting'">
          <rect x="2" y="2" width="20" height="20" rx="2" fill="#ffffff" stroke="#94a3b8" stroke-width="1.4"/>
          <rect x="4" y="4.5" width="16" height="3.5" rx="1" fill="#3b82f6"/>
          <rect x="4" y="10.5" width="11" height="3.5" rx="1" fill="#ef4444"/>
          <rect x="4" y="16.5" width="7" height="3.5" rx="1" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'formatAsTable'">
          <rect x="2" y="3" width="20" height="18" rx="2" fill="#ffffff" stroke="#2563eb" stroke-width="1.6"/>
          <rect x="2" y="3" width="20" height="6" fill="#2563eb"/>
          <line x1="2" y1="13.5" x2="22" y2="13.5" stroke="#cbd5e1" stroke-width="1.2"/>
          <line x1="8.5" y1="9" x2="8.5" y2="21" stroke="#cbd5e1" stroke-width="1.2"/>
          <line x1="15.5" y1="9" x2="15.5" y2="21" stroke="#cbd5e1" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'cellStyles'">
          <rect x="2" y="3" width="9" height="8" rx="1.5" fill="#fed7aa" stroke="#f97316" stroke-width="1.2"/>
          <rect x="13" y="3" width="9" height="8" rx="1.5" fill="#bfdbfe" stroke="#3b82f6" stroke-width="1.2"/>
          <rect x="2" y="13" width="9" height="8" rx="1.5" fill="#bbf7d0" stroke="#22c55e" stroke-width="1.2"/>
          <rect x="13" y="13" width="9" height="8" rx="1.5" fill="#fef08a" stroke="#eab308" stroke-width="1.2"/>
        </g>

        <!-- CELLS -->
        <g *ngIf="name === 'insertCells'">
          <rect x="2" y="3" width="20" height="18" rx="2" fill="#ffffff" stroke="#94a3b8" stroke-width="1.4"/>
          <circle cx="16" cy="15" r="6" fill="#16a34a"/>
          <path d="M16 11.5v7M12.5 15h7" stroke="#ffffff" stroke-width="1.8" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'deleteCells'">
          <rect x="2" y="3" width="20" height="18" rx="2" fill="#ffffff" stroke="#94a3b8" stroke-width="1.4"/>
          <circle cx="16" cy="15" r="6" fill="#dc2626"/>
          <line x1="12.5" y1="15" x2="19.5" y2="15" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'formatCells'">
          <rect x="2" y="3" width="20" height="18" rx="2" fill="#ffffff" stroke="#94a3b8" stroke-width="1.4"/>
          <circle cx="16" cy="15" r="6" fill="#f59e0b"/>
          <path d="M13.5 15a2.5 2.5 0 1 0 5 0 2.5 2.5 0 0 0-5 0z" stroke="#ffffff" stroke-width="1.4"/>
        </g>

        <!-- EDITING -->
        <g *ngIf="name === 'autoSum'">
          <path d="M4 3h16v3.5L13 12l7 5.5V21H4v-4h9L7.5 12 13 7H4V3z" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'fill'">
          <rect x="3.5" y="2.5" width="17" height="19" rx="1.5" fill="#ffffff" stroke="#94a3b8" stroke-width="1.4"/>
          <path d="M12 5.5v12m-4.5-4.5l4.5 4.5 4.5-4.5" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'clear'">
          <path d="M19 6l-6-5a2 2 0 0 0-2.8 0L2.5 9a2 2 0 0 0 0 2.8l5.5 5.5H21v-3l-7-7z" fill="#f43f5e" stroke="#e11d48" stroke-width="1.4"/>
          <path d="M10.5 17.5l6.5-6.5" stroke="#ffffff" stroke-width="1.8"/>
        </g>

        <g *ngIf="name === 'sortFilter'">
          <text x="1" y="10" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">A</text>
          <text x="1" y="21" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">Z</text>
          <path d="M7 6v12m-2.5-2.5L7 18l2.5-2.5" stroke="#16a34a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M12 4h10l-4 5.5v6.5l-2 1.5v-8L12 4z" fill="#2563eb" stroke="#1d4ed8" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'findSelect'">
          <circle cx="10" cy="10" r="7" stroke="#2563eb" stroke-width="2.2" fill="none"/>
          <line x1="15" y1="15" x2="21" y2="21" stroke="#2563eb" stroke-width="2.6" stroke-linecap="round"/>
          <rect x="7" y="7" width="6" height="2" fill="#93c5fd"/>
        </g>

        <g *ngIf="name === 'addIns'">
          <rect x="2" y="2" width="9" height="9" rx="1.5" fill="#ef4444"/>
          <rect x="13" y="2" width="9" height="9" rx="1.5" fill="#10b981"/>
          <rect x="2" y="13" width="9" height="9" rx="1.5" fill="#3b82f6"/>
          <rect x="13" y="13" width="9" height="9" rx="1.5" fill="#f59e0b"/>
        </g>

        <!-- INSERT TAB -->
        <g *ngIf="name === 'pivotTable'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.4"/>
          <path d="M3 4h18v5H3z" fill="#3b82f6"/>
          <path d="M3 9h6v11H3z" fill="#93c5fd"/>
          <path d="M12 12a3 3 0 1 1 3 3" stroke="#d97706" stroke-width="1.5" stroke-linecap="round"/>
          <path d="M15 15l2-1-1-2" fill="#d97706"/>
        </g>

        <g *ngIf="name === 'recommendedPivotTables'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="3" y="4" width="7" height="16" fill="#cbd5e1"/>
          <circle cx="15" cy="13" r="4.5" fill="#ffffff" stroke="#16a34a" stroke-width="1.5"/>
          <path d="M13 13l1.5 1.5 3-3" stroke="#16a34a" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'table'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#1e40af" stroke-width="1.4"/>
          <rect x="3" y="4" width="18" height="4" fill="#1e40af"/>
          <line x1="3" y1="12" x2="21" y2="12" stroke="#93c5fd" stroke-width="1"/>
          <line x1="3" y1="16" x2="21" y2="16" stroke="#93c5fd" stroke-width="1"/>
          <line x1="9" y1="8" x2="9" y2="20" stroke="#93c5fd" stroke-width="1"/>
          <line x1="15" y1="8" x2="15" y2="20" stroke="#93c5fd" stroke-width="1"/>
        </g>

        <g *ngIf="name === 'forms'">
          <circle cx="12" cy="12" r="9" fill="#107c41"/>
          <rect x="8" y="7" width="8" height="10" rx="1" fill="#ffffff"/>
          <line x1="10" y1="10" x2="14" y2="10" stroke="#107c41" stroke-width="1.2"/>
          <line x1="10" y1="13" x2="14" y2="13" stroke="#107c41" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'pictures'">
          <rect x="3" y="4" width="18" height="16" rx="2" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="8" cy="8.5" r="2" fill="#eab308"/>
          <path d="M4 17l5-5 4 4 3-3 4 4v2H4v-2z" fill="#3b82f6"/>
        </g>

        <g *ngIf="name === 'shapes'">
          <circle cx="7" cy="9" r="4" fill="#3b82f6"/>
          <rect x="11" y="9" width="9" height="9" rx="1" fill="#ef4444"/>
          <path d="M12 4l3 5H9z" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'icons'">
          <circle cx="12" cy="8" r="3.5" stroke="#475569" stroke-width="1.4" fill="none"/>
          <path d="M6 19c0-3.3 2.7-6 6-6s6 2.7 6 6" stroke="#475569" stroke-width="1.4" fill="none"/>
        </g>

        <g *ngIf="name === 'models3d'">
          <path d="M12 2l8 4.5v9l-8 4.5-8-4.5v-9L12 2z" fill="#bfdbfe" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M12 2v9m0 0l8-4.5m-8 4.5L4 6.5" stroke="#2563eb" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'smartArt'">
          <path d="M3 6h6l3 4-3 4H3l3-4-3-4z" fill="#3b82f6"/>
          <path d="M10 6h6l3 4-3 4h-6l3-4-3-4z" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'screenshot'">
          <rect x="3" y="4" width="18" height="16" rx="2" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="12" cy="12" r="3.5" stroke="#3b82f6" stroke-width="1.5" fill="none"/>
          <rect x="8" y="2" width="4" height="2" fill="#64748b"/>
        </g>

        <g *ngIf="name === 'checkbox'">
          <rect x="4" y="4" width="16" height="16" rx="2" fill="#ffffff" stroke="#107c41" stroke-width="1.6"/>
          <path d="M8 12l3 3 5-6" stroke="#107c41" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'recommendedCharts'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="5.5" y="11" width="3" height="7" fill="#3b82f6"/>
          <rect x="10.5" y="7" width="3" height="11" fill="#f59e0b"/>
          <rect x="15.5" y="9" width="3" height="9" fill="#10b981"/>
          <circle cx="16" cy="7" r="3" fill="#ffffff" stroke="#dc2626" stroke-width="1.2"/>
          <line x1="18" y1="9" x2="20" y2="11" stroke="#dc2626" stroke-width="1.4"/>
        </g>

        <!-- CHARTS GRID -->
        <g *ngIf="name === 'chartColumn'">
          <rect x="4" y="8" width="3.5" height="12" fill="#2563eb"/>
          <rect x="10" y="4" width="3.5" height="16" fill="#3b82f6"/>
          <rect x="16" y="11" width="3.5" height="9" fill="#60a5fa"/>
        </g>

        <g *ngIf="name === 'chartLine'">
          <path d="M4 16l5-8 5 4 6-8" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="4" cy="16" r="1.5" fill="#2563eb"/>
          <circle cx="9" cy="8" r="1.5" fill="#2563eb"/>
          <circle cx="14" cy="12" r="1.5" fill="#2563eb"/>
          <circle cx="20" cy="4" r="1.5" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'chartPie'">
          <circle cx="12" cy="12" r="9" fill="#3b82f6"/>
          <path d="M12 12V3a9 9 0 0 1 8.5 6L12 12z" fill="#f59e0b"/>
          <path d="M12 12l8.5-3A9 9 0 0 1 12 21V12z" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'chartBar'">
          <rect x="4" y="4" width="14" height="3.5" fill="#2563eb"/>
          <rect x="4" y="10" width="10" height="3.5" fill="#3b82f6"/>
          <rect x="4" y="16" width="16" height="3.5" fill="#60a5fa"/>
        </g>

        <g *ngIf="name === 'chartHierarchy'">
          <rect x="3" y="4" width="18" height="16" fill="none" stroke="#2563eb" stroke-width="1.2"/>
          <rect x="5" y="6" width="6" height="12" fill="#93c5fd"/>
          <rect x="13" y="6" width="6" height="5" fill="#fde68a"/>
          <rect x="13" y="13" width="6" height="5" fill="#a7f3d0"/>
        </g>

        <g *ngIf="name === 'chartStatistic'">
          <rect x="4" y="10" width="4" height="10" fill="#2563eb"/>
          <rect x="9" y="5" width="4" height="15" fill="#3b82f6"/>
          <rect x="14" y="8" width="4" height="12" fill="#60a5fa"/>
          <path d="M2 18c4-1 6-12 10-12s6 11 10 12" stroke="#dc2626" stroke-width="1.2" fill="none"/>
        </g>

        <g *ngIf="name === 'chartWaterfall'">
          <rect x="4" y="12" width="3" height="8" fill="#10b981"/>
          <rect x="9" y="8" width="3" height="4" fill="#10b981"/>
          <rect x="14" y="12" width="3" height="5" fill="#ef4444"/>
          <rect x="19" y="7" width="3" height="13" fill="#3b82f6"/>
        </g>

        <g *ngIf="name === 'chartCombo'">
          <rect x="4" y="10" width="3" height="10" fill="#3b82f6"/>
          <rect x="10" y="6" width="3" height="14" fill="#3b82f6"/>
          <rect x="16" y="12" width="3" height="8" fill="#3b82f6"/>
          <path d="M4 14l6-6 6 5 4-7" stroke="#f59e0b" stroke-width="1.8" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'chartScatter'">
          <circle cx="5" cy="18" r="2" fill="#2563eb"/>
          <circle cx="9" cy="12" r="2" fill="#2563eb"/>
          <circle cx="13" cy="14" r="2" fill="#2563eb"/>
          <circle cx="16" cy="8" r="2" fill="#2563eb"/>
          <circle cx="20" cy="5" r="2" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'maps'">
          <circle cx="12" cy="12" r="9" fill="#bfdbfe" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M7 10c2-1 4 2 6 1s3-3 4-1c-1 3-3 5-5 5s-3-2-5-5z" fill="#22c55e"/>
          <circle cx="15" cy="8" r="2" fill="#dc2626"/>
        </g>

        <g *ngIf="name === 'pivotChart'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.2"/>
          <rect x="5" y="10" width="3" height="8" fill="#3b82f6"/>
          <rect x="10" y="6" width="3" height="12" fill="#f59e0b"/>
          <circle cx="16" cy="14" r="3.5" fill="#10b981"/>
        </g>

        <!-- SPARKLINES -->
        <g *ngIf="name === 'sparklineLine'">
          <path d="M3 15l5-7 5 5 7-9" stroke="#2563eb" stroke-width="2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'sparklineColumn'">
          <rect x="3" y="10" width="3" height="10" fill="#2563eb"/>
          <rect x="8" y="5" width="3" height="15" fill="#2563eb"/>
          <rect x="13" y="12" width="3" height="8" fill="#2563eb"/>
          <rect x="18" y="8" width="3" height="12" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'sparklineWinLoss'">
          <rect x="3" y="5" width="3" height="7" fill="#16a34a"/>
          <rect x="8" y="12" width="3" height="7" fill="#dc2626"/>
          <rect x="13" y="5" width="3" height="7" fill="#16a34a"/>
          <rect x="18" y="5" width="3" height="7" fill="#16a34a"/>
        </g>

        <!-- FILTERS & LINKS -->
        <g *ngIf="name === 'slicer'">
          <rect x="4" y="4" width="16" height="16" rx="2" fill="#ffffff" stroke="#2563eb" stroke-width="1.4"/>
          <rect x="6" y="6" width="12" height="3" rx="0.5" fill="#3b82f6"/>
          <rect x="6" y="11" width="12" height="3" rx="0.5" fill="#93c5fd"/>
          <rect x="6" y="15" width="12" height="3" rx="0.5" fill="#cbd5e1"/>
        </g>

        <g *ngIf="name === 'timeline'">
          <rect x="4" y="4" width="16" height="16" rx="2" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="12" cy="12" r="5" stroke="#f59e0b" stroke-width="1.5" fill="none"/>
          <path d="M12 9v3l2 2" stroke="#f59e0b" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'link'">
          <path d="M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.5 1.5" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round"/>
          <path d="M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.5-1.5" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'comment'">
          <path d="M4 5a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-5l-4 4v-4H6a2 2 0 0 1-2-2V5z" fill="#8b5cf6" stroke="#7c3aed" stroke-width="1.2"/>
          <line x1="8" y1="8" x2="16" y2="8" stroke="#ffffff" stroke-width="1.5" stroke-linecap="round"/>
          <line x1="8" y1="11" x2="13" y2="11" stroke="#ffffff" stroke-width="1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'textBox'">
          <rect x="4" y="4" width="16" height="16" rx="1.5" stroke="#475569" stroke-width="1.2" stroke-dasharray="2 2" fill="none"/>
          <text x="8.5" y="16" font-family="Segoe UI, sans-serif" font-size="12" font-weight="bold" fill="currentColor">A</text>
        </g>

        <g *ngIf="name === 'headerFooter'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="4" y="3" width="16" height="4" fill="#3b82f6"/>
          <rect x="4" y="17" width="16" height="4" fill="#3b82f6"/>
        </g>

        <g *ngIf="name === 'wordArt'">
          <text x="4" y="18" font-family="Segoe UI, sans-serif" font-size="17" font-weight="900" fill="#3b82f6" transform="rotate(-15 12 12)">A</text>
        </g>

        <g *ngIf="name === 'signatureLine'">
          <path d="M4 18h16" stroke="#475569" stroke-width="1.4"/>
          <text x="5" y="16" font-family="Segoe UI, sans-serif" font-size="11" font-weight="bold" fill="#dc2626">X</text>
          <path d="M12 5l5 7h-4v4h-2v-4H7l5-7z" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'insertObject'">
          <rect x="4" y="5" width="16" height="14" rx="2" fill="#fed7aa" stroke="#ea580c" stroke-width="1.2"/>
          <rect x="7" y="8" width="10" height="8" rx="1" fill="#ffffff"/>
        </g>

        <g *ngIf="name === 'equation'">
          <text x="5" y="18" font-family="Times New Roman, serif" font-size="18" font-style="italic" fill="currentColor">π</text>
        </g>

        <g *ngIf="name === 'symbol'">
          <text x="5" y="18" font-family="Times New Roman, serif" font-size="18" font-weight="bold" fill="currentColor">Ω</text>
        </g>

        <!-- FORMULAS TAB -->
        <g *ngIf="name === 'insertFunction'">
          <text x="2" y="17" font-family="Times New Roman, Georgia, serif" font-size="18" font-style="italic" font-weight="bold" fill="#2563eb">fx</text>
          <circle cx="18" cy="15" r="4.5" fill="#16a34a"/>
          <path d="M18 12.5v5M15.5 15h5" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'recentlyUsed'">
          <circle cx="12" cy="12" r="8.5" stroke="#2563eb" stroke-width="1.6" fill="none"/>
          <path d="M12 7v5l3.5 2" stroke="#2563eb" stroke-width="1.6" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'financial'">
          <circle cx="12" cy="12" r="8.5" fill="#fef08a" stroke="#ca8a04" stroke-width="1.2"/>
          <text x="8.5" y="16.5" font-family="Segoe UI, sans-serif" font-size="13" font-weight="bold" fill="#a16207">$</text>
        </g>

        <g *ngIf="name === 'logical'">
          <path d="M12 4v14m-7-9h14M5 9l-2 5h4l-2-5zm14 0l-2 5h4l-2-5z" stroke="#2563eb" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'textFunction'">
          <path d="M4 5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v14H5a1 1 0 0 1-1-1V5z" fill="#bfdbfe" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M12 5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1h-7V5z" fill="#ffffff" stroke="#2563eb" stroke-width="1.2"/>
          <text x="14" y="15" font-family="Segoe UI, sans-serif" font-size="11" font-weight="bold" fill="#2563eb">A</text>
        </g>

        <g *ngIf="name === 'dateTime'">
          <rect x="4" y="5" width="16" height="15" rx="2" fill="#ffffff" stroke="#dc2626" stroke-width="1.2"/>
          <rect x="4" y="5" width="16" height="4" fill="#dc2626"/>
          <circle cx="15" cy="14" r="4.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M15 11.5v2.5l1.5 1" stroke="#2563eb" stroke-width="1.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'lookupRef'">
          <path d="M4 5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v14H5a1 1 0 0 1-1-1V5z" fill="#fed7aa" stroke="#ea580c" stroke-width="1.2"/>
          <path d="M12 5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1h-7V5z" fill="#ffffff" stroke="#ea580c" stroke-width="1.2"/>
          <circle cx="15" cy="12" r="3" stroke="#2563eb" stroke-width="1.4" fill="none"/>
          <line x1="17" y1="14" x2="20" y2="17" stroke="#2563eb" stroke-width="1.6"/>
        </g>

        <g *ngIf="name === 'mathTrig'">
          <text x="3" y="11" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="#2563eb">+</text>
          <text x="14" y="11" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="#2563eb">-</text>
          <text x="3" y="20" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="#2563eb">×</text>
          <text x="14" y="20" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="#2563eb">÷</text>
        </g>

        <g *ngIf="name === 'moreFunctions'">
          <rect x="4" y="6" width="16" height="3" rx="0.5" fill="#3b82f6"/>
          <rect x="4" y="11" width="16" height="3" rx="0.5" fill="#f59e0b"/>
          <rect x="4" y="16" width="16" height="3" rx="0.5" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'nameManager'">
          <rect x="3" y="4" width="18" height="15" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <rect x="3" y="4" width="18" height="4" fill="#3b82f6"/>
          <path d="M13 11h5l-2 3 2 3h-5v-6z" fill="#10b981"/>
        </g>

        <g *ngIf="name === 'defineName'">
          <path d="M4 12l6-6h7v7l-6 6-7-7z" fill="#ffffff" stroke="#16a34a" stroke-width="1.4"/>
          <circle cx="14" cy="9" r="1.5" fill="#16a34a"/>
        </g>

        <g *ngIf="name === 'useInFormula'">
          <rect x="4" y="4" width="16" height="16" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <path d="M8 8h8M8 12h5" stroke="#2563eb" stroke-width="1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'createFromSelection'">
          <rect x="4" y="4" width="16" height="16" rx="1.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.4"/>
          <rect x="4" y="4" width="16" height="5" fill="#bfdbfe"/>
          <rect x="4" y="4" width="5" height="16" fill="#bfdbfe"/>
        </g>

        <g *ngIf="name === 'tracePrecedents'">
          <circle cx="6" cy="12" r="3" fill="#2563eb"/>
          <path d="M9 12h8m-3-3l3 3-3 3" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
          <rect x="17" y="9" width="5" height="6" rx="1" fill="#93c5fd" stroke="#2563eb" stroke-width="1"/>
        </g>

        <g *ngIf="name === 'traceDependents'">
          <rect x="3" y="9" width="5" height="6" rx="1" fill="#93c5fd" stroke="#2563eb" stroke-width="1"/>
          <path d="M8 12h8m-3-3l3 3-3 3" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="19" cy="12" r="3" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'removeArrows'">
          <path d="M4 12h12m-3-3l3 3-3 3" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
          <line x1="16" y1="5" x2="22" y2="11" stroke="#dc2626" stroke-width="2.2" stroke-linecap="round"/>
          <line x1="22" y1="5" x2="16" y2="11" stroke="#dc2626" stroke-width="2.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'showFormulas'">
          <rect x="3" y="5" width="18" height="14" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <text x="6" y="15" font-family="Times New Roman, serif" font-size="11" font-style="italic" font-weight="bold" fill="#2563eb">fx</text>
        </g>

        <g *ngIf="name === 'errorChecking'">
          <path d="M12 3l9 16H3l9-16z" fill="#fef08a" stroke="#ca8a04" stroke-width="1.4"/>
          <line x1="12" y1="9" x2="12" y2="13" stroke="#a16207" stroke-width="2" stroke-linecap="round"/>
          <circle cx="12" cy="16" r="1" fill="#a16207"/>
        </g>

        <g *ngIf="name === 'evaluateFormula'">
          <rect x="3" y="5" width="18" height="14" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <circle cx="12" cy="12" r="4.5" stroke="#2563eb" stroke-width="1.6" fill="none"/>
          <line x1="15.5" y1="15.5" x2="19" y2="19" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'watchWindow'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <circle cx="8" cy="12" r="3.5" stroke="#2563eb" stroke-width="1.5" fill="none"/>
          <circle cx="16" cy="12" r="3.5" stroke="#2563eb" stroke-width="1.5" fill="none"/>
          <path d="M11.5 12h1" stroke="#2563eb" stroke-width="1.5"/>
        </g>

        <g *ngIf="name === 'calcOptions'">
          <rect x="4" y="3" width="16" height="18" rx="2" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <rect x="6" y="5" width="12" height="4" fill="#94a3b8"/>
          <circle cx="8" cy="12" r="1.2" fill="#2563eb"/>
          <circle cx="12" cy="12" r="1.2" fill="#2563eb"/>
          <circle cx="16" cy="12" r="1.2" fill="#2563eb"/>
        </g>

        <g *ngIf="name === 'calculateNow'">
          <rect x="4" y="3" width="16" height="18" rx="2" fill="#ffffff" stroke="#2563eb" stroke-width="1.4"/>
          <path d="M7 6h10M7 10h10M7 14h5M7 17h10" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M15 13l3 3-3 3" stroke="#16a34a" stroke-width="1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'calculateSheet'">
          <rect x="3" y="3" width="18" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="11" y="9" width="9" height="11" rx="1" fill="#bfdbfe" stroke="#2563eb" stroke-width="1"/>
          <line x1="5" y1="7" x2="10" y2="7" stroke="#64748b" stroke-width="1.2"/>
          <line x1="5" y1="11" x2="9" y2="11" stroke="#64748b" stroke-width="1.2"/>
        </g>

        <!-- DATA TAB -->
        <g *ngIf="name === 'getData'">
          <ellipse cx="12" cy="6" rx="8" ry="3.5" fill="#3b82f6" stroke="#1d4ed8" stroke-width="1.2"/>
          <path d="M4 6v10c0 1.9 3.6 3.5 8 3.5s8-1.6 8-3.5V6" stroke="#1d4ed8" stroke-width="1.2" fill="none"/>
          <path d="M4 11c0 1.9 3.6 3.5 8 3.5s8-1.6 8-3.5" stroke="#1d4ed8" stroke-width="1.2" fill="none"/>
          <path d="M12 10v6m-2-2l2 2 2-2" stroke="#ffffff" stroke-width="1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'fromTextCsv'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="6" y="11" width="12" height="6" fill="#16a34a"/>
          <text x="7" y="16" font-family="Segoe UI, sans-serif" font-size="5" font-weight="bold" fill="#ffffff">CSV</text>
        </g>

        <g *ngIf="name === 'fromWeb'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="15" cy="15" r="4.5" fill="#bfdbfe" stroke="#2563eb" stroke-width="1"/>
          <line x1="10.5" y1="15" x2="19.5" y2="15" stroke="#2563eb" stroke-width="0.8"/>
          <ellipse cx="15" cy="15" rx="2" ry="4.5" stroke="#2563eb" stroke-width="0.8" fill="none"/>
        </g>

        <g *ngIf="name === 'fromTableRange'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#2563eb" stroke-width="1.4"/>
          <rect x="3" y="4" width="18" height="4" fill="#2563eb"/>
          <rect x="7" y="10" width="10" height="7" stroke="#10b981" stroke-width="1.2" stroke-dasharray="2 1" fill="none"/>
        </g>

        <g *ngIf="name === 'fromPicture'">
          <rect x="4" y="4" width="16" height="16" rx="2" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="8" cy="8.5" r="2" fill="#eab308"/>
          <path d="M4 16l5-5 4 4 3-3 4 4" stroke="#2563eb" stroke-width="1.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'recentSources'">
          <ellipse cx="10" cy="8" rx="6" ry="2.5" fill="#3b82f6"/>
          <path d="M4 8v6c0 1.4 2.7 2.5 6 2.5s6-1.1 6-2.5V8" stroke="#1d4ed8" stroke-width="1" fill="none"/>
          <circle cx="16" cy="15" r="4" fill="#ffffff" stroke="#f59e0b" stroke-width="1.2"/>
          <path d="M16 13v2l1.5 1" stroke="#f59e0b" stroke-width="1" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'existingConnections'">
          <rect x="3" y="7" width="8" height="10" rx="1" fill="#3b82f6"/>
          <rect x="13" y="7" width="8" height="10" rx="1" fill="#10b981"/>
          <path d="M11 12h2" stroke="#475569" stroke-width="2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'refreshAll'">
          <path d="M19 8A8 8 0 0 0 5 8m0 0l3-3M5 8l3 3" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M5 16a8 8 0 0 0 14 0m0 0l-3-3m3 3l-3 3" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'queriesConnections'">
          <rect x="3" y="6" width="11" height="13" rx="1" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <rect x="9" y="3" width="11" height="13" rx="1" fill="#bfdbfe" stroke="#2563eb" stroke-width="1.2"/>
          <path d="M8 14h7l-2 2" stroke="#2563eb" stroke-width="1.2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'properties'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <circle cx="12" cy="12" r="3.5" stroke="#f59e0b" stroke-width="1.4" fill="none"/>
        </g>

        <g *ngIf="name === 'workbookLinks'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <path d="M9 14l6-4m-4-2l2-2a2 2 0 0 1 2.8 2.8l-1.4 1.4" stroke="#2563eb" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'sortAsc'">
          <text x="2" y="10" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">A</text>
          <text x="2" y="21" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">Z</text>
          <path d="M14 5v14m-3.5-3.5L14 19l3.5-3.5" stroke="#16a34a" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'sortDesc'">
          <text x="2" y="10" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">Z</text>
          <text x="2" y="21" font-family="Segoe UI, sans-serif" font-size="9" font-weight="bold" fill="currentColor">A</text>
          <path d="M14 5v14m-3.5-3.5L14 19l3.5-3.5" stroke="#16a34a" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'sortDialog'">
          <rect x="2.5" y="3" width="19" height="18" rx="2" fill="#ffffff" stroke="#2563eb" stroke-width="1.5"/>
          <text x="4.5" y="11" font-family="Segoe UI, sans-serif" font-size="8.5" font-weight="bold" fill="currentColor">A</text>
          <text x="4.5" y="19" font-family="Segoe UI, sans-serif" font-size="8.5" font-weight="bold" fill="currentColor">Z</text>
          <path d="M14 6v12m-3-3l3 3 3-3" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'filter'">
          <path d="M2.5 3h19l-7.5 9v8l-4 2v-10L2.5 3z" fill="#2563eb" stroke="#1d4ed8" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'clearFilter'">
          <path d="M3 4h14l-5 6v6l-3 1v-7L3 4z" fill="#93c5fd" stroke="#2563eb" stroke-width="1"/>
          <line x1="16" y1="12" x2="22" y2="18" stroke="#dc2626" stroke-width="2" stroke-linecap="round"/>
          <line x1="22" y1="12" x2="16" y2="18" stroke="#dc2626" stroke-width="2" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'reapplyFilter'">
          <path d="M3 4h14l-5 6v6l-3 1v-7L3 4z" fill="#93c5fd" stroke="#2563eb" stroke-width="1"/>
          <path d="M18 12a3 3 0 1 0 3 3" stroke="#16a34a" stroke-width="1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'advancedFilter'">
          <path d="M3 4h18l-7 8v7l-4 2v-9L3 4z" fill="#2563eb" stroke="#1d4ed8" stroke-width="1.2"/>
          <circle cx="18" cy="18" r="3" stroke="#f59e0b" stroke-width="1.4" fill="#ffffff"/>
        </g>

        <g *ngIf="name === 'textToColumns'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <rect x="4" y="6" width="6" height="12" fill="#bfdbfe"/>
          <rect x="14" y="6" width="6" height="12" fill="#bbf7d0"/>
          <path d="M11 12h2m-1-2l2 2-2 2" stroke="#2563eb" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'flashFill'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <path d="M13 5l-5 7h5l-2 7 7-9h-5l2-5z" fill="#eab308" stroke="#ca8a04" stroke-width="1"/>
        </g>

        <g *ngIf="name === 'removeDuplicates'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <rect x="3" y="4" width="18" height="4" fill="#3b82f6"/>
          <rect x="3" y="8" width="18" height="4" fill="#ffffff"/>
          <rect x="3" y="12" width="18" height="4" fill="#fee2e2"/>
          <line x1="5" y1="14" x2="19" y2="14" stroke="#dc2626" stroke-width="1.5"/>
          <circle cx="16" cy="15" r="4.5" fill="#dc2626"/>
          <line x1="14" y1="13" x2="18" y2="17" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
          <line x1="18" y1="13" x2="14" y2="17" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'dataValidation'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#475569" stroke-width="1.2"/>
          <circle cx="9" cy="12" r="4.5" fill="#16a34a"/>
          <path d="M7 12l1.5 1.5 3-3" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="16" cy="12" r="3.5" fill="#dc2626"/>
          <line x1="14" y1="12" x2="18" y2="12" stroke="#ffffff" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'consolidate'">
          <rect x="3" y="3" width="10" height="9" rx="1" fill="#bfdbfe" stroke="#2563eb" stroke-width="1"/>
          <rect x="11" y="3" width="10" height="9" rx="1" fill="#fed7aa" stroke="#ea580c" stroke-width="1"/>
          <rect x="7" y="12" width="10" height="9" rx="1" fill="#bbf7d0" stroke="#16a34a" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'manageDataModel'">
          <ellipse cx="8" cy="7" rx="5" ry="2" fill="#3b82f6"/>
          <path d="M3 7v6c0 1.1 2.2 2 5 2s5-.9 5-2V7" stroke="#1d4ed8" stroke-width="1" fill="none"/>
          <ellipse cx="16" cy="12" rx="5" ry="2" fill="#10b981"/>
          <path d="M11 12v6c0 1.1 2.2 2 5 2s5-.9 5-2v-6" stroke="#047857" stroke-width="1" fill="none"/>
          <line x1="9" y1="11" x2="14" y2="14" stroke="#475569" stroke-width="1.2"/>
        </g>

        <g *ngIf="name === 'analyzeData'">
          <rect x="3" y="4" width="18" height="16" rx="2" fill="#ffffff" stroke="#2563eb" stroke-width="1.2"/>
          <rect x="5.5" y="11" width="3" height="7" fill="#3b82f6"/>
          <rect x="10.5" y="7" width="3" height="11" fill="#3b82f6"/>
          <path d="M18 5l1 2 2 1-2 1-1 2-1-2-2-1 2-1 1-2z" fill="#eab308"/>
        </g>

        <g *ngIf="name === 'whatIfAnalysis'">
          <rect x="4" y="3" width="16" height="18" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <text x="9" y="16" font-family="Segoe UI, sans-serif" font-size="14" font-weight="bold" fill="#2563eb">?</text>
        </g>

        <g *ngIf="name === 'forecastSheet'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <path d="M5 14l4-5 4 2" stroke="#2563eb" stroke-width="1.5" stroke-linecap="round"/>
          <path d="M13 11l4-4 3 2" stroke="#ea580c" stroke-width="1.5" stroke-dasharray="2 1.5" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'groupRows'">
          <rect x="4" y="4" width="16" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <path d="M2 7h2v10H2" stroke="#2563eb" stroke-width="1.5" fill="none"/>
        </g>

        <g *ngIf="name === 'ungroupRows'">
          <rect x="4" y="4" width="16" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <path d="M4 12H1m2-2L1 12l2 2" stroke="#2563eb" stroke-width="1.4" stroke-linecap="round"/>
        </g>

        <g *ngIf="name === 'subtotal'">
          <rect x="3" y="4" width="18" height="16" rx="1.5" fill="#ffffff" stroke="#64748b" stroke-width="1.2"/>
          <text x="6" y="15" font-family="Segoe UI, sans-serif" font-size="11" font-weight="bold" fill="#2563eb">Σ</text>
        </g>

        <!-- DIALOG LAUNCHER & DROPDOWN -->
        <g *ngIf="name === 'dialogLauncher'">
          <rect x="1" y="9" width="8" height="8" rx="0.5" fill="none" stroke="currentColor" stroke-width="0.9"/>
          <path d="M5 13l3-3m0 0H5m3 0v3" stroke="currentColor" stroke-width="0.9" stroke-linecap="round" stroke-linejoin="round"/>
        </g>

        <g *ngIf="name === 'dropdownArrow'">
          <path d="M7 10l5 5 5-5z" fill="currentColor"/>
        </g>
      </svg>
    </span>
  `,
  styles: [`
    :host {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      line-height: 1;
    }
    .excel-icon-wrapper {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }
    .excel-svg {
      display: block;
      width: 100%;
      height: 100%;
    }

    /* Dark Theme Overrides */
    :host-context([data-theme="dark"]) svg [stroke="#334155"],
    :host-context([data-theme="dark"]) svg [stroke="#475569"],
    :host-context([data-theme="dark"]) svg [stroke="#1e293b"],
    :host-context([data-theme="dark"]) svg [stroke="#000000"],
    :host-context([data-theme="dark"]) svg [stroke="#323130"],
    :host-context([data-theme="dark"]) svg [stroke="#242424"],
    :host-context(.dark-theme) svg [stroke="#334155"],
    :host-context(.dark-theme) svg [stroke="#475569"],
    :host-context(.dark-theme) svg [stroke="#1e293b"],
    :host-context(.dark-theme) svg [stroke="#000000"],
    :host-context(.dark-theme) svg [stroke="#323130"],
    :host-context(.dark-theme) svg [stroke="#242424"] {
      stroke: #f3f2f1;
    }

    :host-context([data-theme="dark"]) svg [fill="#334155"],
    :host-context([data-theme="dark"]) svg [fill="#475569"],
    :host-context([data-theme="dark"]) svg [fill="#1e293b"],
    :host-context([data-theme="dark"]) svg [fill="#000000"],
    :host-context([data-theme="dark"]) svg [fill="#323130"],
    :host-context([data-theme="dark"]) svg [fill="#242424"],
    :host-context(.dark-theme) svg [fill="#334155"],
    :host-context(.dark-theme) svg [fill="#475569"],
    :host-context(.dark-theme) svg [fill="#1e293b"],
    :host-context(.dark-theme) svg [fill="#000000"],
    :host-context(.dark-theme) svg [fill="#323130"],
    :host-context(.dark-theme) svg [fill="#242424"] {
      fill: #f3f2f1;
    }
  `],
})
export class ExcelIconComponent {
  @Input() name: string = '';
  @Input() size: number = 18;
  @Input() color: string = 'currentColor';
}
