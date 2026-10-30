import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  Output,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PptIconComponent } from '../../ppt-icon/ppt-icon.component';
import { PptRibbonGroupComponent } from '../../ribbon-group/ppt-ribbon-group.component';
import { SlideTransition, TransitionDirection, TransitionId } from '../../../../models/ppt.types';
import { PptRibbonFoldService } from '../../../../services/ppt-ribbon-fold.service';
import { computePopupPosition, PopupPosition } from '../../../../utils/popup-position';
import { HostListener } from '@angular/core';

export interface TransitionTileDef {
  id: TransitionId;
  label: string;
  enabled: boolean;
}

@Component({
  selector: 'app-ppt-transitions-tab',
  standalone: true,
  imports: [CommonModule, PptIconComponent, PptRibbonGroupComponent],
  templateUrl: './ppt-transitions-tab.component.html',
  styleUrls: ['./ppt-transitions-tab.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class PptTransitionsTabComponent {
  @Input() currentTransition: SlideTransition = { type: 'none' };

  @Output() readonly selectTransition = new EventEmitter<SlideTransition>();
  @Output() readonly previewTransition = new EventEmitter<void>();

  get canPreview(): boolean {
    return !!this.currentTransition && this.currentTransition.type !== 'none';
  }

  // Primary visible tiles in gallery
  readonly primaryTiles: TransitionTileDef[] = [
    { id: 'none', label: 'None', enabled: true },
    { id: 'morph', label: 'Morph', enabled: false },
    { id: 'fade', label: 'Fade', enabled: true },
    { id: 'push', label: 'Push', enabled: true },
    { id: 'flip', label: 'Flip', enabled: true },
    { id: 'cube', label: 'Cube', enabled: true },
    { id: 'reveal', label: 'Reveal', enabled: false },
    { id: 'cut', label: 'Cut', enabled: false },
    { id: 'randomBars', label: 'Random Bars', enabled: false },
    { id: 'shape', label: 'Shape', enabled: false },
    { id: 'uncover', label: 'Uncover', enabled: false },
    { id: 'cover', label: 'Cover', enabled: false },
    { id: 'flash', label: 'Flash', enabled: false },
    { id: 'fallOver', label: 'Fall Over', enabled: false },
  ];

  isMorePanelOpen = false;
  isEffectOptionsOpen = false;
  morePos: PopupPosition = { left: 0, top: 0 };
  effectPos: PopupPosition = { left: 0, top: 0 };
  private lastMoreAnchor: HTMLElement | null = null;
  private lastEffectAnchor: HTMLElement | null = null;

  constructor(public foldService: PptRibbonFoldService) {}

  
  get isCompact(): boolean {
    return this.foldService.isTabCompact('transitions');
  }

  isOverflow(groupId: string): boolean {
    return this.foldService.isGroupInOverflow('transitions', groupId);
  }

  isFolded(groupId: string): boolean {
    return this.foldService.isGroupFolded('transitions', groupId);
  }

  isFoldOpen(groupId: string): boolean {
    return this.foldService.openFoldGroupId() === groupId;
  }

  onToggleFold(groupId: string): void {
    this.foldService.toggleFold(groupId);
  }

  onCloseFold(): void {
    this.foldService.closeFold();
  }

  onPickTransition(id: TransitionId, enabled: boolean): void {
    if (!enabled) return;
    this.isMorePanelOpen = false;
    const direction = id === 'push' ? (this.currentTransition.direction || 'fromRight') : undefined;
    this.selectTransition.emit({ type: id, direction });
  }

  onToggleMore(event: MouseEvent): void {
    this.lastMoreAnchor = event.currentTarget as HTMLElement;
    this.recalculatePositions();
    event.stopPropagation();
    this.isMorePanelOpen = !this.isMorePanelOpen;
    this.isEffectOptionsOpen = false;
  }

  onToggleEffectOptions(event: MouseEvent): void {
    this.lastEffectAnchor = event.currentTarget as HTMLElement;
    this.recalculatePositions();
    event.stopPropagation();
    if (this.currentTransition.type !== 'push') return;
    this.isEffectOptionsOpen = !this.isEffectOptionsOpen;
    this.isMorePanelOpen = false;
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    this.recalculatePositions();
  }

  @HostListener('document:click')
  onDocClick(): void {
    if (this.isMorePanelOpen || this.isEffectOptionsOpen) {
      this.isMorePanelOpen = false;
      this.isEffectOptionsOpen = false;
    }
  }

  @HostListener('document:keydown.escape')
  onEscape(): void {
    this.isMorePanelOpen = false;
    this.isEffectOptionsOpen = false;
  }

  private recalculatePositions(): void {
    if (this.lastMoreAnchor) {
      this.morePos = computePopupPosition(this.lastMoreAnchor, { width: 180, height: 160 });
    }
    if (this.lastEffectAnchor) {
      this.effectPos = computePopupPosition(this.lastEffectAnchor, { width: 140, height: 90 });
    }
  }

  onSelectDirection(dir: TransitionDirection): void {
    this.isEffectOptionsOpen = false;
    this.selectTransition.emit({
      type: 'push',
      direction: dir,
    });
  }
}

