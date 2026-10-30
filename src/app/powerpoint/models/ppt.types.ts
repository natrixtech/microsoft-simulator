export type ThemeId =
  | 'simple-light'
  | 'simple-dark'
  | 'streamline'
  | 'focus'
  | 'shift'
  | 'momentum'
  | 'paradigm'
  | 'material';

export interface ThemeDef {
  id: ThemeId;
  name: string;
  background: string;
  titleFont: string;
  titleColor: string;
  bodyFont: string;
  bodyColor: string;
  accentColors: string[]; // for theme gallery tiles preview
}

export type SlideLayout =
  | 'title'
  | 'titleAndBody'
  | 'titleOnly'
  | 'sectionHeader'
  | 'twoColumns'
  | 'blank';

export type PlaceholderType = 'title' | 'subtitle' | 'body' | 'body2';

export interface PlaceholderDef {
  type: PlaceholderType;
  x: number;
  y: number;
  w: number;
  h: number;
  fontSize: number;
  promptText: string;
  align: 'left' | 'center' | 'right';
  isBullet?: boolean;
}

export interface LayoutDef {
  id: SlideLayout;
  name: string;
  placeholders: PlaceholderDef[];
}

export type ElementKind = 'image' | 'textBox' | 'shape' | 'table';

export interface BaseSlideElement {
  id: string;
  kind: ElementKind;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ImageSlideElement extends BaseSlideElement {
  kind: 'image';
  file: string;
  alt?: string;
}

export interface TextBoxSlideElement extends BaseSlideElement {
  kind: 'textBox';
  text: string;
  html?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  align?: 'left' | 'center' | 'right';
}

export type ShapeKind =
  | 'rectangle'
  | 'roundedRectangle'
  | 'ellipse'
  | 'triangle'
  | 'arrow'
  | 'star';

export interface ShapeSlideElement extends BaseSlideElement {
  kind: 'shape';
  shape: ShapeKind;
  fill?: string;
  stroke?: string;
}

export interface TableSlideElement extends BaseSlideElement {
  kind: 'table';
  rows: number;
  columns: number;
  cells: string[][];
}

export type SlideElement =
  | ImageSlideElement
  | TextBoxSlideElement
  | ShapeSlideElement
  | TableSlideElement;

export type TransitionId =
  | 'none'
  | 'fade'
  | 'push'
  | 'flip'
  | 'cube'
  | 'morph'
  | 'wipe'
  | 'split'
  | 'reveal'
  | 'cut'
  | 'randomBars'
  | 'shape'
  | 'uncover'
  | 'cover'
  | 'flash'
  | 'fallOver';

export type TransitionDirection = 'fromRight' | 'fromLeft';

export interface SlideTransition {
  type: TransitionId;
  direction?: TransitionDirection;
}

export interface Slide {
  id: string;
  layout: SlideLayout;
  title?: string;
  titleHtml?: string;
  subtitle?: string;
  subtitleHtml?: string;
  body?: string;
  bodyHtml?: string;
  body2?: string;
  body2Html?: string;
  background?: string; // color name, hex, or image filename
  elements: SlideElement[];
  transition?: SlideTransition;
}

export interface Deck {
  theme: ThemeId;
  slides: Slide[];
}

export interface PptSnapshot {
  deck: Deck;
  currentSlideIndex: number;
  selectedElementId: string | null;
  editingPlaceholder: {
    slideId: string;
    field: PlaceholderType;
  } | null;
}

export type RibbonTabId =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'design'
  | 'transitions'
  | 'animations'
  | 'slideshow'
  | 'record'
  | 'review'
  | 'view'
  | 'help';

export interface RibbonButtonDef {
  id: string;
  label: string;
  icon: string;
  enabled: boolean;
  demoTarget: string;
  tooltip?: string;
}

export interface RibbonGroupDef {
  id: string;
  label: string;
  icon: string;
  buttons: RibbonButtonDef[];
}


// --- Goal Types ---

export interface PptSlideTextGoal {
  type: 'slideText';
  slide?: number;
  placeholder?: 'title' | 'subtitle' | 'body' | 'body2' | 'any';
  contains?: string | string[];
  equals?: string;
  notContains?: string[];
  minLines?: number;
  exact?: boolean;
}

export interface PptTextStyleGoal {
  type: 'textStyle';
  slide?: number;
  placeholder?: 'title' | 'subtitle' | 'body' | 'body2' | 'any';
  text?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font?: string;
  size?: number;
  color?: string;
  align?: 'left' | 'center' | 'right';
}

export interface PptListGoal {
  type: 'list';
  slide?: number;
  placeholder?: 'title' | 'subtitle' | 'body' | 'body2' | 'any';
  listType?: 'bullet' | 'numbered';
  minItems?: number;
}

export interface PptSlideCountGoal {
  type: 'slideCount';
  count?: number;
  min?: number;
}

export interface PptImageGoal {
  type: 'image';
  slide?: number;
  file?: string | null;
  position?: 'left' | 'right' | 'top' | 'bottom' | 'center';
  minWidth?: number;
  maxWidth?: number;
}

export interface PptTextBoxGoal {
  type: 'textBox';
  slide?: number;
  contains?: string;
}

export interface PptShapeGoal {
  type: 'shape';
  slide?: number;
  shape?: ShapeKind;
  count?: number;
}

export interface PptTableGoal {
  type: 'table';
  slide?: number;
  rows?: number;
  columns?: number;
  cells?: string[][];
}

export interface PptBackgroundGoal {
  type: 'background';
  slide?: number;
  color?: string;
  file?: string;
}

export interface PptLayoutGoal {
  type: 'layout';
  slide?: number;
  layout: SlideLayout;
}

export interface PptThemeGoal {
  type: 'theme';
  theme: string;
}

export interface PptTransitionGoal {
  type: 'transition';
  slide?: number;
  transition: string;
}

export interface PptSlideOrderGoal {
  type: 'slideOrder';
  titles: string[];
}

export interface PptActionGoal {
  type: 'action';
  action: string;
  slide?: number;
  menu?: string;
}

export type PptGoal =
  | PptSlideTextGoal
  | PptTextStyleGoal
  | PptListGoal
  | PptSlideCountGoal
  | PptImageGoal
  | PptTextBoxGoal
  | PptShapeGoal
  | PptTableGoal
  | PptBackgroundGoal
  | PptLayoutGoal
  | PptThemeGoal
  | PptTransitionGoal
  | PptSlideOrderGoal
  | PptActionGoal;

export interface PptGoalContext {
  currentSlideIndex: number;
  startingTheme: string;
  actionLog: any[];
}
