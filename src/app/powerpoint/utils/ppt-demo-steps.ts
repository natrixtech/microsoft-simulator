import { LessonTask } from '../../shared/models/activity.types';
import { Deck, PptGoal } from '../models/ppt.types';

export type DemoActionKind =
  | 'click'
  | 'rightClick'
  | 'dblClick'
  | 'type'
  | 'key'
  | 'select'
  | 'selectText'
  | 'drag'
  | 'ensureTab'
  | 'wait';

export interface DemoAction {
  kind: DemoActionKind;
  target?: string;
  tab?: 'home' | 'insert' | 'design' | 'transitions';
  text?: string;
  key?: string;
  value?: string | number;
  durationMs?: number;
  dragToTarget?: string;
  fromCoords?: { x: number; y: number };
  toCoords?: { x: number; y: number };
}

export interface DemoStep {
  label: string;
  goalIndex: number;
  actions: DemoAction[];
}

export function buildDemoSteps(
  goals: PptGoal[],
  startingDeck: Deck,
  task: LessonTask,
  startingTheme: string = 'simple-light',
  images: Array<{ file: string; alt?: string }> = []
): DemoStep[] {
  const steps: DemoStep[] = [];

  for (let i = 0; i < (goals || []).length; i++) {
    const goal = goals[i];
    const step = buildStepForGoal(goal, i, startingDeck, task, startingTheme, images);
    steps.push(step);
  }

  return steps;
}

function buildStepForGoal(
  goal: PptGoal,
  goalIndex: number,
  startingDeck: Deck,
  task: LessonTask,
  startingTheme: string,
  images: Array<{ file: string; alt?: string }>
): DemoStep {
  switch (goal.type) {
    case 'slideText':
      return buildSlideTextStep(goal, goalIndex, startingDeck);
    case 'textStyle':
      return buildTextStyleStep(goal, goalIndex, startingDeck);
    case 'list':
      return buildListStep(goal, goalIndex, startingDeck);
    case 'slideCount':
      return buildSlideCountStep(goal, goalIndex, startingDeck);
    case 'image':
      return buildImageStep(goal, goalIndex, images);
    case 'textBox':
      return buildTextBoxStep(goal, goalIndex);
    case 'shape':
      return buildShapeStep(goal, goalIndex);
    case 'table':
      return buildTableStep(goal, goalIndex);
    case 'background':
      return buildBackgroundStep(goal, goalIndex, images);
    case 'layout':
      return buildLayoutStep(goal, goalIndex);
    case 'theme':
      return buildThemeStep(goal, goalIndex, startingTheme);
    case 'transition':
      return buildTransitionStep(goal, goalIndex);
    case 'slideOrder':
      return buildSlideOrderStep(goal, goalIndex, startingDeck);
    case 'action':
      return buildActionStep(goal, goalIndex);
    default:
      return {
        label: 'Do the action',
        goalIndex,
        actions: [{ kind: 'click', target: 'slide-canvas' }],
      };
  }
}

function buildSlideTextStep(goal: any, goalIndex: number, deck: Deck): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });

  const f = goal.placeholder || 'any';
  const targetField = f === 'any' ? 'title' : f;
  const target = `placeholder-${targetField}`;

  let textToType = 'Text';
  let isMultilineList = false;
  let multilineArray: string[] = [];

  const equalsVal = goal.equals !== undefined && goal.equals !== null ? goal.equals : goal.equal;
  if (equalsVal !== undefined && equalsVal !== null) {
    textToType = String(equalsVal);
  } else if (goal.minLines && goal.minLines > 1) {
    isMultilineList = true;
    if (goal.contains) {
      if (Array.isArray(goal.contains)) {
        multilineArray = goal.contains.map(String);
      } else {
        multilineArray = [String(goal.contains)];
      }
    } else {
      multilineArray = [];
    }
    while (multilineArray.length < goal.minLines) {
      multilineArray.push(`Line ${multilineArray.length + 1}`);
    }
  } else if (goal.contains) {
    if (Array.isArray(goal.contains)) {
      textToType = goal.contains.join(' ');
    } else {
      textToType = String(goal.contains);
    }
  } else if (goal.text) {
    textToType = String(goal.text);
  }

  actions.push({ kind: 'click', target });
  actions.push({ kind: 'key', key: 'Ctrl+A' });

  if (isMultilineList) {
    for (let idx = 0; idx < multilineArray.length; idx++) {
      actions.push({ kind: 'type', text: multilineArray[idx] });
      if (idx < multilineArray.length - 1) {
        actions.push({ kind: 'key', key: 'Enter' });
      }
    }
  } else {
    actions.push({ kind: 'type', text: textToType });
  }

  actions.push({ kind: 'click', target: 'slide-canvas' });

  return {
    label: 'Type the text',
    goalIndex,
    actions,
  };
}

function buildTextStyleStep(goal: any, goalIndex: number, deck: Deck): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });

  const targetText = goal.text || '';
  let target = 'placeholder-title';
  const slideObj = deck.slides[slideNum - 1];
  if (slideObj) {
    if (slideObj.title && slideObj.title.includes(targetText)) {
      target = 'placeholder-title';
    } else if (slideObj.body && slideObj.body.includes(targetText)) {
      target = 'placeholder-body';
    } else if (slideObj.subtitle && slideObj.subtitle.includes(targetText)) {
      target = 'placeholder-subtitle';
    } else if (slideObj.elements) {
      const tbIdx = slideObj.elements.findIndex(
        (el) => el.kind === 'textBox' && (el as any).text?.includes(targetText)
      );
      if (tbIdx >= 0) {
        target = `slide-element-${tbIdx + 1}`;
      }
    }
  }

  actions.push({ kind: 'dblClick', target });
  if (targetText) {
    actions.push({ kind: 'selectText', text: targetText, target });
  }

  if (goal.bold) {
    actions.push({ kind: 'click', target: 'home-bold' });
  }
  if (goal.italic) {
    actions.push({ kind: 'click', target: 'home-italic' });
  }
  if (goal.underline) {
    actions.push({ kind: 'click', target: 'home-underline' });
  }
  if (goal.font) {
    actions.push({ kind: 'select', target: 'home-font-family', value: goal.font });
  }
  if (goal.size) {
    actions.push({ kind: 'select', target: 'home-font-size', value: goal.size });
  }
  if (goal.color) {
    actions.push({ kind: 'click', target: 'home-font-color-arrow' });
    actions.push({ kind: 'click', target: `color-swatch-${goal.color}` });
  }
  if (goal.align) {
    actions.push({ kind: 'click', target: `home-align-${goal.align}` });
  }

  actions.push({ kind: 'click', target: 'slide-canvas' });

  return {
    label: 'Format the text',
    goalIndex,
    actions,
  };
}

function buildListStep(goal: any, goalIndex: number, deck: Deck): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });

  const target = goal.placeholder ? `placeholder-${goal.placeholder}` : 'placeholder-body';
  actions.push({ kind: 'click', target });

  const slideObj = deck.slides[slideNum - 1];
  const bodyText = slideObj?.body || '';
  const lines = bodyText.split('\n').filter((l) => l.trim().length > 0);
  const needed = goal.minItems || 3;

  if (lines.length < needed) {
    actions.push({ kind: 'key', key: 'Ctrl+A' });
    for (let i = 1; i <= needed; i++) {
      actions.push({ kind: 'type', text: `Item ${i}` });
      if (i < needed) {
        actions.push({ kind: 'key', key: 'Enter' });
      }
    }
  }

  actions.push({ kind: 'key', key: 'Ctrl+A' });

  if (goal.listType === 'numbered') {
    actions.push({ kind: 'click', target: 'home-numbering' });
  } else {
    actions.push({ kind: 'click', target: 'home-bullets' });
  }

  actions.push({ kind: 'click', target: 'slide-canvas' });

  return {
    label: 'Make the list',
    goalIndex,
    actions,
  };
}

function buildSlideCountStep(goal: any, goalIndex: number, deck: Deck): DemoStep {
  const actions: DemoAction[] = [];
  const curCount = deck.slides.length;
  const targetCount = goal.count || goal.min || 2;

  if (curCount < targetCount) {
    for (let i = curCount; i < targetCount; i++) {
      actions.push({ kind: 'click', target: 'home-new-slide' });
    }
  } else if (curCount > targetCount) {
    for (let i = curCount; i > targetCount; i--) {
      actions.push({ kind: 'rightClick', target: `slide-thumb-${i}` });
      actions.push({ kind: 'click', target: 'context-delete-slide' });
    }
  } else {
    actions.push({ kind: 'click', target: 'home-new-slide' });
  }

  return {
    label: targetCount < curCount ? 'Delete a slide' : 'Add a slide',
    goalIndex,
    actions,
  };
}

function buildImageStep(
  goal: any,
  goalIndex: number,
  images: Array<{ file: string; alt?: string }>
): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'insert' });
  actions.push({ kind: 'click', target: 'insert-pictures' });

  let picIdx = 1;
  if (goal.file && images && images.length > 0) {
    const found = images.findIndex((img) => img.file === goal.file);
    if (found >= 0) picIdx = found + 1;
  }

  actions.push({ kind: 'click', target: `picture-item-${picIdx}` });
  actions.push({ kind: 'click', target: 'picture-dialog-insert' });

  let curX = 280;
  let curY = 145;
  let curW = 400;
  let curH = 250;

  if (goal.position) {
    let toX = 480;
    let toY = 270;
    if (goal.position === 'left') {
      toX = 240;
      toY = 270;
      curX = 40;
      curY = 145;
    } else if (goal.position === 'right') {
      toX = 720;
      toY = 270;
      curX = 520;
      curY = 145;
    } else if (goal.position === 'top') {
      toX = 480;
      toY = 135;
      curX = 280;
      curY = 10;
    } else if (goal.position === 'bottom') {
      toX = 480;
      toY = 405;
      curX = 280;
      curY = 280;
    } else if (goal.position === 'center') {
      toX = 480;
      toY = 270;
      curX = 280;
      curY = 145;
    }
    actions.push({
      kind: 'drag',
      target: 'slide-element-last',
      fromCoords: { x: 480, y: 270 },
      toCoords: { x: toX, y: toY },
    });
  }

  if (goal.minWidth !== undefined && curW < goal.minWidth) {
    const deltaW = goal.minWidth - curW + 40;
    const deltaH = Math.round(deltaW / 1.6);
    actions.push({
      kind: 'drag',
      target: 'handle-se',
      fromCoords: { x: curX + curW, y: curY + curH },
      toCoords: { x: curX + curW + deltaW, y: curY + curH + deltaH },
    });
  } else if (goal.maxWidth !== undefined && curW > goal.maxWidth) {
    const deltaW = goal.maxWidth - curW - 40;
    const deltaH = Math.round(deltaW / 1.6);
    actions.push({
      kind: 'drag',
      target: 'handle-se',
      fromCoords: { x: curX + curW, y: curY + curH },
      toCoords: { x: curX + curW + deltaW, y: curY + curH + deltaH },
    });
  }

  return {
    label: 'Insert a picture',
    goalIndex,
    actions,
  };
}

function buildTextBoxStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'insert' });
  actions.push({ kind: 'click', target: 'insert-text-box' });
  actions.push({
    kind: 'drag',
    target: 'slide-canvas',
    fromCoords: { x: 160, y: 380 },
    toCoords: { x: 560, y: 440 },
  });

  const text = goal.contains || goal.text || 'Text';
  actions.push({ kind: 'wait', durationMs: 400 });
  actions.push({ kind: 'type', text });
  actions.push({ kind: 'click', target: 'slide-canvas' });

  return {
    label: 'Add a text box',
    goalIndex,
    actions,
  };
}

function buildShapeStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'insert' });
  actions.push({ kind: 'click', target: 'insert-shapes' });

  let shapeTarget = 'rectangle';
  if (goal.shape === 'ellipse' || goal.shape === 'oval') {
    shapeTarget = 'oval';
  } else if (goal.shape === 'roundedRectangle' || goal.shape === 'rounded-rectangle') {
    shapeTarget = 'rounded-rectangle';
  } else if (goal.shape) {
    shapeTarget = goal.shape;
  }

  actions.push({ kind: 'click', target: `insert-shape-${shapeTarget}` });
  actions.push({
    kind: 'drag',
    target: 'slide-canvas',
    fromCoords: { x: 260, y: 150 },
    toCoords: { x: 560, y: 350 },
  });

  const count = goal.count || 1;
  for (let c = 2; c <= count; c++) {
    actions.push({ kind: 'click', target: 'insert-shapes' });
    actions.push({ kind: 'click', target: `insert-shape-${shapeTarget}` });
    actions.push({
      kind: 'drag',
      target: 'slide-canvas',
      fromCoords: { x: 260 + (c - 1) * 60, y: 150 },
      toCoords: { x: 560 + (c - 1) * 60, y: 350 },
    });
  }

  return {
    label: 'Add a shape',
    goalIndex,
    actions,
  };
}

function buildTableStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'insert' });
  actions.push({ kind: 'click', target: 'insert-table' });

  const rows = goal.rows || 3;
  const cols = goal.columns || 2;
  actions.push({ kind: 'click', target: `table-cell-${rows}x${cols}` });

  if (goal.cells && Array.isArray(goal.cells)) {
    for (let r = 0; r < goal.cells.length; r++) {
      const rowItem = goal.cells[r];
      if (Array.isArray(rowItem)) {
        for (let c = 0; c < rowItem.length; c++) {
          const val = rowItem[c];
          if (val !== undefined && val !== null && String(val).length > 0) {
            actions.push({ kind: 'click', target: `slide-table-cell-${r + 1}x${c + 1}` });
            actions.push({ kind: 'type', text: String(val) });
          }
        }
      } else if (rowItem && typeof rowItem === 'object') {
        const c = rowItem as any;
        if (c.text) {
          actions.push({ kind: 'click', target: `slide-table-cell-${c.row}x${c.col}` });
          actions.push({ kind: 'type', text: String(c.text) });
        }
      }
    }
  }

  return {
    label: 'Insert a table',
    goalIndex,
    actions,
  };
}

function buildBackgroundStep(
  goal: any,
  goalIndex: number,
  images: Array<{ file: string; alt?: string }>
): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'design' });
  actions.push({ kind: 'click', target: 'design-background-styles' });
  actions.push({ kind: 'click', target: 'design-format-background' });

  if (goal.file) {
    actions.push({ kind: 'click', target: 'format-bg-radio-picture' });
    actions.push({ kind: 'click', target: 'format-bg-insert-pic' });
    let picIdx = 1;
    if (images && images.length > 0) {
      const found = images.findIndex((img) => img.file === goal.file);
      if (found >= 0) picIdx = found + 1;
    }
    actions.push({ kind: 'click', target: `picture-item-${picIdx}` });
    actions.push({ kind: 'click', target: 'picture-dialog-insert' });
  } else {
    const color = goal.color === 'any' || !goal.color ? 'lightBlue' : goal.color;
    actions.push({ kind: 'click', target: 'format-bg-radio-solid' });
    actions.push({ kind: 'click', target: 'format-bg-color-btn' });
    actions.push({ kind: 'click', target: `color-swatch-${color}` });
  }

  actions.push({ kind: 'click', target: 'format-bg-close' });

  return {
    label: 'Change the background',
    goalIndex,
    actions,
  };
}

function buildLayoutStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'home' });
  actions.push({ kind: 'click', target: 'home-layout' });
  actions.push({ kind: 'click', target: `layout-choice-${goal.layout}` });

  return {
    label: 'Change the layout',
    goalIndex,
    actions,
  };
}

function buildThemeStep(goal: any, goalIndex: number, startingTheme: string): DemoStep {
  const actions: DemoAction[] = [];
  actions.push({ kind: 'ensureTab', tab: 'design' });

  let targetTheme = goal.theme;
  if (!targetTheme || targetTheme === 'any') {
    targetTheme = startingTheme === 'simple-light' ? 'simple-dark' : 'simple-light';
  }

  actions.push({ kind: 'click', target: `design-theme-${targetTheme}` });

  return {
    label: 'Change the theme',
    goalIndex,
    actions,
  };
}

function buildTransitionStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const slideNum = goal.slide || 1;
  actions.push({ kind: 'click', target: `slide-thumb-${slideNum}` });
  actions.push({ kind: 'ensureTab', tab: 'transitions' });

  let trans = (goal.transition || 'fade').toLowerCase();
  if (trans === 'any') trans = 'fade';

  if (trans.startsWith('slide') || trans === 'push') {
    actions.push({ kind: 'click', target: 'transitions-push' });
    const dirGoal = (goal.direction || '').toLowerCase();
    if (dirGoal || trans.includes('left') || trans.includes('right')) {
      const dir =
        dirGoal === 'fromleft' || trans.includes('left')
          ? 'transitions-effect-from-left'
          : 'transitions-effect-from-right';
      actions.push({ kind: 'click', target: 'transitions-effect-options' });
      actions.push({ kind: 'click', target: dir });
    }
  } else {
    actions.push({ kind: 'click', target: `transitions-${trans}` });
  }

  return {
    label: 'Add a transition',
    goalIndex,
    actions,
  };
}

function buildSlideOrderStep(goal: any, goalIndex: number, deck: Deck): DemoStep {
  const actions: DemoAction[] = [];
  const targetTitles: string[] = goal.titles || [];

  const currentOrder = deck.slides.map((s) => s.title || '');

  for (let toIdx = 0; toIdx < targetTitles.length; toIdx++) {
    const title = targetTitles[toIdx];
    const fromIdx = currentOrder.findIndex((t) => t.toLowerCase().includes(title.toLowerCase()));
    if (fromIdx !== -1 && fromIdx !== toIdx) {
      actions.push({
        kind: 'drag',
        target: `slide-thumb-${fromIdx + 1}`,
        dragToTarget: `slide-thumb-${toIdx + 1}`,
      });
      const [moved] = currentOrder.splice(fromIdx, 1);
      currentOrder.splice(toIdx, 0, moved);
    }
  }

  return {
    label: 'Reorder the slides',
    goalIndex,
    actions,
  };
}

function buildActionStep(goal: any, goalIndex: number): DemoStep {
  const actions: DemoAction[] = [];
  const act = goal.action;

  if (act === 'undo') {
    actions.push({ kind: 'click', target: 'home-new-slide' });
    actions.push({ kind: 'click', target: 'home-undo' });
  } else if (act === 'redo') {
    actions.push({ kind: 'click', target: 'home-new-slide' });
    actions.push({ kind: 'click', target: 'home-undo' });
    actions.push({ kind: 'click', target: 'home-redo' });
  } else if (act === 'selectSlide') {
    const s = goal.slide || 2;
    actions.push({ kind: 'click', target: `slide-thumb-${s}` });
  } else if (act === 'selectElement') {
    const s = goal.slide || 1;
    actions.push({ kind: 'click', target: `slide-thumb-${s}` });
    actions.push({ kind: 'click', target: 'slide-element-1' });
  } else if (act === 'openMenu') {
    if (goal.menu === 'Insert') {
      actions.push({ kind: 'click', target: 'insert-tab' });
    } else if (goal.menu === 'Slide') {
      const s = goal.slide || 1;
      actions.push({ kind: 'rightClick', target: `slide-thumb-${s}` });
    } else {
      actions.push({ kind: 'click', target: `${(goal.menu || 'home').toLowerCase()}-tab` });
    }
  } else {
    actions.push({ kind: 'click', target: 'slide-thumb-1' });
  }

  return {
    label: 'Do the action',
    goalIndex,
    actions,
  };
}
