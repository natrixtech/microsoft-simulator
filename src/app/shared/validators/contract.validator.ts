import { ActivityPayload, AppKind } from '../models/activity.types';
import { COLORS } from '../utils/editor.utils';

export const GOAL_TYPES: Record<AppKind, string[]> = {
  word: [
    'textStyle',
    'paragraph',
    'list',
    'textContains',
    'textCount',
    'textOrder',
    'image',
    'pageSetup',
    'table',
    'element',
    'action',
    'style',
    'chart',
  ],
  excel: [
    'cellValue',
    'cellFormula',
    'cellFormat',
    'sorted',
    'filter',
    'chart',
    'pivotTable',
    'comment',
    'conditionalFormat',
    'dataValidation',
    'noDuplicates',
    'action',
  ],
  powerpoint: [
    'slideText',
    'textStyle',
    'list',
    'slideCount',
    'image',
    'textBox',
    'shape',
    'table',
    'background',
    'layout',
    'theme',
    'transition',
    'slideOrder',
    'action',
  ],
};

export function parseCellAddress(addr: string): [number, number] {
  const m = String(addr).trim().toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!m) throw new Error(`Invalid cell address: ${addr}`);
  let col = 0;
  for (let i = 0; i < m[1].length; i++) {
    col = col * 26 + (m[1].charCodeAt(i) - 64);
  }
  return [parseInt(m[2], 10) - 1, col - 1];
}

export function validateCellRange(rng: string): void {
  const parts = String(rng).split(':');
  if (parts.length > 2) throw new Error(`Invalid range: ${rng}`);
  parts.forEach(parseCellAddress);
}

export function validateActivity(kind: AppKind, data: unknown): ActivityPayload {
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Activity must be an object.');
  }

  const act = data as Record<string, any>;

  if (typeof act['assetBaseUrl'] !== 'string') {
    throw new Error('assetBaseUrl is required.');
  }

  if (!['LIGHT', 'DARK'].includes(act['theme'])) {
    throw new Error('theme must be LIGHT or DARK.');
  }

  const validateShowMeHow = (val: unknown) => {
    if (val === undefined) return;
    if (kind === 'excel') {
      if (typeof val === 'boolean') return;
      if (typeof val === 'object' && val !== null && !Array.isArray(val)) {
        const smh = val as Record<string, any>;
        if (smh['description'] !== undefined && typeof smh['description'] !== 'string') {
          throw new Error('description must be a string.');
        }
        if (smh['sampleSheet'] !== undefined) {
          const ss = smh['sampleSheet'];
          if (!ss || typeof ss !== 'object' || Array.isArray(ss)) {
            throw new Error('sampleSheet.cells must be an object.');
          }
          if (
            ss['cells'] !== undefined &&
            (!ss['cells'] || typeof ss['cells'] !== 'object' || Array.isArray(ss['cells']))
          ) {
            throw new Error('sampleSheet.cells must be an object.');
          }
        }
        return;
      }
      throw new Error('showMeHow must be a boolean or an object.');
    } else {
      if (typeof val !== 'boolean') {
        throw new Error('showMeHow must be a boolean.');
      }
    }
  };

  if (act['showMeHow'] !== undefined) {
    validateShowMeHow(act['showMeHow']);
  }

  const validateGoals = (goals: unknown[]) => {
    if (!Array.isArray(goals)) throw new Error('goals must be an array.');

    for (const rawGoal of goals) {
      if (!rawGoal || typeof rawGoal !== 'object') {
        throw new Error('Goal must be an object.');
      }
      const g = rawGoal as Record<string, any>;
      if (!GOAL_TYPES[kind].includes(g['type'])) {
        throw new Error(`Unknown goal type: ${g['type']}`);
      }

      const requiredKeys: Record<string, string[]> =
        kind === 'word'
          ? {
              textStyle: ['text'],
              paragraph: ['text'],
              list: ['items', 'listType'],
              textCount: ['text', 'count'],
              textOrder: ['sequence'],
              element: ['element'],
              action: ['action'],
              style: ['name'],
            }
          : kind === 'excel'
          ? {
              cellValue: ['cells'],
              cellFormula: ['cell'],
              cellFormat: ['range'],
              sorted: ['range', 'byColumn', 'order'],
              filter: ['range'],
              comment: ['cell'],
              conditionalFormat: ['range'],
              dataValidation: ['range'],
              noDuplicates: ['range'],
              action: ['action'],
            }
          : {
              textStyle: ['text'],
              list: ['placeholder', 'listType'],
              layout: ['layout'],
              theme: ['theme'],
              slideOrder: ['titles'],
              action: ['action'],
            };

      const goalType = String(g['type']);
      const required = requiredKeys[goalType] || [];
      for (const key of required) {
        if (g[key] === undefined || g[key] === null) {
          throw new Error(`${goalType}.${key} is required.`);
        }
      }

      for (const key of [
        'text',
        'action',
        'element',
        'placeholder',
        'layout',
        'theme',
        'cell',
        'range',
        'name',
        'appliedTo',
        'title',
      ]) {
        if (g[key] !== undefined && typeof g[key] !== 'string') {
          throw new Error(`${goalType}.${key} must be text.`);
        }
      }

      for (const key of ['superscript', 'subscript']) {
        if (g[key] !== undefined && typeof g[key] !== 'boolean') {
          throw new Error(`${goalType}.${key} must be a boolean.`);
        }
      }

      for (const key of [
        'items',
        'sequence',
        'titles',
        'notContains',
        'allowed',
      ]) {
        if (
          g[key] !== undefined &&
          (!Array.isArray(g[key]) || g[key].some((v: unknown) => typeof v !== 'string'))
        ) {
          throw new Error(`${goalType}.${key} must be a list of strings.`);
        }
      }

      for (const key of [
        'count',
        'min',
        'minItems',
        'minLines',
        'rows',
        'columns',
        'slide',
      ]) {
        if (g[key] !== undefined && (!Number.isInteger(g[key]) || g[key] < 0)) {
          throw new Error(`${goalType}.${key} must be a non-negative integer.`);
        }
      }

      if (g['slide'] !== undefined && g['slide'] < 1) {
        throw new Error('Slide numbers start at 1.');
      }

      if (kind === 'word' && goalType === 'chart') {
        if (g['chartType'] !== undefined) {
          const allowedChartTypes = ['column', 'bar', 'line', 'pie'];
          if (!allowedChartTypes.includes(g['chartType'])) {
            throw new Error(`chart.chartType must be one of: ${allowedChartTypes.join(', ')}.`);
          }
        }
      }

      if (goalType === 'table') {
        if (g['style'] !== undefined) {
          const allowedStyles = ['Plain', 'Banded Rows', 'Grid', 'Grid Blue', 'Grid Green', 'List Blue'];
          if (!allowedStyles.includes(g['style'])) {
            throw new Error(`table.style must be one of: ${allowedStyles.join(', ')}.`);
          }
        }
        if (g['cellColors'] !== undefined) {
          if (!Array.isArray(g['cellColors'])) {
            throw new Error('table.cellColors must be an array.');
          }
          for (const cc of g['cellColors']) {
            if (
              !cc ||
              typeof cc !== 'object' ||
              !Number.isInteger(cc.row) ||
              cc.row < 0 ||
              !Number.isInteger(cc.col) ||
              cc.col < 0 ||
              typeof cc.color !== 'string'
            ) {
              throw new Error('table.cellColors must contain objects with non-negative integer row, col, and string color.');
            }
          }
        }
      }

      if (kind === 'excel') {
        for (const key of ['cell', 'range']) {
          if (g[key]) validateCellRange(g[key]);
        }
        if (goalType === 'cellValue') {
          const cells = g['cells'];
          if (typeof cells !== 'object' || Array.isArray(cells) || !cells) {
            throw new Error('cellValue.cells must be an object.');
          }
          Object.keys(cells).forEach(parseCellAddress);
        }

        if (g['numberFormat'] !== undefined) {
          const allowedNumberFormats = ['automatic', 'percent', 'currency', 'date'];
          if (!allowedNumberFormats.includes(g['numberFormat'])) {
            throw new Error(`numberFormat must be one of: ${allowedNumberFormats.join(', ')}.`);
          }
        }

        const allowedColors = Object.keys(COLORS);
        for (const colorKey of ['textColor', 'fillColor', 'color']) {
          if (g[colorKey] !== undefined && !allowedColors.includes(g[colorKey])) {
            throw new Error(`${colorKey} must be one of the 20 named colors.`);
          }
        }

        if (goalType === 'chart' && g['chartType'] !== undefined) {
          const allowedChartTypes = ['column', 'bar', 'line', 'pie'];
          if (!allowedChartTypes.includes(g['chartType'])) {
            throw new Error(`chart.chartType must be one of: ${allowedChartTypes.join(', ')}.`);
          }
        }

        if (goalType === 'pivotTable' && g['summarize'] !== undefined) {
          const allowedSummarize = ['SUM', 'COUNT', 'AVERAGE'];
          if (!allowedSummarize.includes(g['summarize'])) {
            throw new Error(`pivotTable.summarize must be one of: ${allowedSummarize.join(', ')}.`);
          }
        }

        if ((goalType === 'filter' || goalType === 'conditionalFormat') && g['condition'] !== undefined) {
          const allowedConditions = ['equals', 'greaterThan', 'lessThan', 'between', 'contains', 'isEmpty'];
          if (!allowedConditions.includes(g['condition'])) {
            throw new Error(`${goalType}.condition must be one of: ${allowedConditions.join(', ')}.`);
          }
        }
      }
    }
  };

  if (kind === 'word' && !act['tasks']) {
    if (
      typeof act['instruction'] !== 'string' ||
      typeof act['startingDocument'] !== 'string'
    ) {
      throw new Error('instruction and startingDocument are required.');
    }
    validateGoals(act['goals'] || []);
  } else {
    if (
      !['lesson', 'assessment'].includes(act['mode']) ||
      typeof act['title'] !== 'string' ||
      !Array.isArray(act['tasks']) ||
      !act['tasks'].length
    ) {
      throw new Error('mode, title and tasks are required.');
    }

    if (
      (act['mode'] === 'lesson' && act['tasks'].length !== 1) ||
      (act['mode'] === 'assessment' && act['tasks'].length < 2)
    ) {
      throw new Error(
        'A lesson needs one task; an assessment needs at least two.'
      );
    }

    act['tasks'].forEach((t: Record<string, any>, i: number) => {
      if (
        typeof t['title'] !== 'string' ||
        typeof t['explanation'] !== 'string' ||
        !Array.isArray(t['instructions']) ||
        t['instructions'].length < 1 ||
        t['instructions'].length > 3 ||
        t['instructions'].some((x: unknown) => typeof x !== 'string')
      ) {
        throw new Error(
          'Each task needs a title, explanation and 1–3 instructions.'
        );
      }

      if (t['showMeHow'] !== undefined) {
        validateShowMeHow(t['showMeHow']);
      }
      validateGoals(t['goals'] || []);

      if (kind === 'word') {
        const s = t['startingDocument'];
        if (s === null && i > 0) return;
        if (typeof s !== 'string') {
          throw new Error('startingDocument is required for the first task.');
        }
      } else {
        const key = kind === 'excel' ? 'startingSheet' : 'startingDeck';
        const s = t[key];
        if (s === null && i > 0) return;
        if (!s || typeof s !== 'object') {
          throw new Error(`${key} is required for the first task.`);
        }

        if (
          kind === 'excel' &&
          (!s.cells || typeof s.cells !== 'object' || Array.isArray(s.cells))
        ) {
          throw new Error('startingSheet.cells must be an object.');
        }

        if (kind === 'excel' && s.formats && typeof s.formats === 'object') {
          const allowedNumberFormats = ['automatic', 'percent', 'currency', 'date'];
          const allowedColors = Object.keys(COLORS);
          for (const fmt of Object.values(s.formats)) {
            if (fmt && typeof fmt === 'object') {
              const f = fmt as Record<string, any>;
              if (f['numberFormat'] !== undefined && !allowedNumberFormats.includes(f['numberFormat'])) {
                throw new Error(`numberFormat must be one of: ${allowedNumberFormats.join(', ')}.`);
              }
              for (const colorKey of ['textColor', 'fillColor', 'color']) {
                if (f[colorKey] !== undefined && !allowedColors.includes(f[colorKey])) {
                  throw new Error(`${colorKey} must be one of the 20 named colors.`);
                }
              }
            }
          }
        }

        if (
          kind === 'powerpoint' &&
          (!Array.isArray(s.slides) || !s.slides.length)
        ) {
          throw new Error('startingDeck needs at least one slide.');
        }
      }
    });
  }

  return act as ActivityPayload;
}
