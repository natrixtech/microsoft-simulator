import { RibbonTabId } from '../models/ppt.types';

export interface RibbonTabConfig {
  id: RibbonTabId;
  label: string;
  enabled: boolean;
}

export const RIBBON_TABS: RibbonTabConfig[] = [
  { id: 'file', label: 'File', enabled: false },
  { id: 'home', label: 'Home', enabled: true },
  { id: 'insert', label: 'Insert', enabled: true },
  { id: 'draw', label: 'Draw', enabled: false },
  { id: 'design', label: 'Design', enabled: true },
  { id: 'transitions', label: 'Transitions', enabled: true },
  { id: 'animations', label: 'Animations', enabled: false },
  { id: 'slideshow', label: 'Slide Show', enabled: false },
  { id: 'record', label: 'Record', enabled: false },
  { id: 'review', label: 'Review', enabled: false },
  { id: 'view', label: 'View', enabled: false },
  { id: 'help', label: 'Help', enabled: false },
];

export const TAB_FOLD_ORDER: Record<string, string[]> = {
  home: ['slides', 'paragraph', 'drawing', 'font', 'clipboard'],
  insert: ['illustrations', 'text', 'images', 'tables', 'slides', 'camera', 'zoom', 'links', 'comments', 'symbols', 'media'],
  design: ['edit', 'slides', 'customize'],
  transitions: ['timing', 'gallery'],
};

export const TAB_OVERFLOW_ORDER: Record<string, string[]> = {
  home: ['designer', 'addins', 'voice', 'editing'],
  insert: ['media', 'symbols', 'comments', 'links', 'zoom', 'camera'],
  design: ['edit', 'slides'],
  transitions: ['timing'],
};

export interface SearchCommand {
  id: string;
  label: string;
  group: string;
  tab: RibbonTabId;
  icon: string;
  demoTarget: string;
}

export const SEARCH_COMMANDS: SearchCommand[] = [
  { id: 'undo', label: 'Undo', group: 'Undo', tab: 'home', icon: 'undo', demoTarget: 'home-undo' },
  { id: 'redo', label: 'Redo', group: 'Undo', tab: 'home', icon: 'redo', demoTarget: 'home-redo' },
  { id: 'new-slide', label: 'New Slide', group: 'Slides', tab: 'home', icon: 'new-slide', demoTarget: 'home-new-slide' },
  { id: 'layout', label: 'Slide Layout', group: 'Slides', tab: 'home', icon: 'layout', demoTarget: 'home-layout' },
  { id: 'bold', label: 'Bold', group: 'Font', tab: 'home', icon: 'bold', demoTarget: 'home-bold' },
  { id: 'italic', label: 'Italic', group: 'Font', tab: 'home', icon: 'italic', demoTarget: 'home-italic' },
  { id: 'underline', label: 'Underline', group: 'Font', tab: 'home', icon: 'underline', demoTarget: 'home-underline' },
  { id: 'font-color', label: 'Font Color', group: 'Font', tab: 'home', icon: 'font-color', demoTarget: 'home-font-color' },
  { id: 'bullets', label: 'Bullets', group: 'Paragraph', tab: 'home', icon: 'bullets', demoTarget: 'home-bullets' },
  { id: 'numbering', label: 'Numbering', group: 'Paragraph', tab: 'home', icon: 'numbering', demoTarget: 'home-numbering' },
  { id: 'align-left', label: 'Align Left', group: 'Paragraph', tab: 'home', icon: 'align-left', demoTarget: 'home-align-left' },
  { id: 'align-center', label: 'Align Center', group: 'Paragraph', tab: 'home', icon: 'align-center', demoTarget: 'home-align-center' },
  { id: 'align-right', label: 'Align Right', group: 'Paragraph', tab: 'home', icon: 'align-right', demoTarget: 'home-align-right' },
  { id: 'insert-table', label: 'Table', group: 'Tables', tab: 'insert', icon: 'table', demoTarget: 'insert-table' },
  { id: 'insert-picture', label: 'Pictures', group: 'Images', tab: 'insert', icon: 'picture', demoTarget: 'insert-pictures' },
  { id: 'insert-shape', label: 'Shapes', group: 'Illustrations', tab: 'insert', icon: 'shapes', demoTarget: 'insert-shapes' },
  { id: 'insert-textbox', label: 'Text Box', group: 'Text', tab: 'insert', icon: 'text-box', demoTarget: 'insert-text-box' },
  { id: 'theme-gallery', label: 'Themes', group: 'Themes', tab: 'design', icon: 'themes', demoTarget: 'design-theme-focus' },
  { id: 'background-styles', label: 'Background Styles', group: 'Customize', tab: 'design', icon: 'background-styles', demoTarget: 'design-background-styles' },
  { id: 'transition-fade', label: 'Transition: Fade', group: 'Transitions', tab: 'transitions', icon: 'transition', demoTarget: 'transitions-fade' },
  { id: 'transition-push', label: 'Transition: Push', group: 'Transitions', tab: 'transitions', icon: 'transition', demoTarget: 'transitions-push' },
];

