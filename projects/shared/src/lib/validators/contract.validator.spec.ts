import { validateActivity, parseCellAddress, validateCellRange } from './contract.validator';

describe('ContractValidator', () => {
  it('should parse cell coordinates and ranges properly', () => {
    expect(parseCellAddress('A1')).toEqual([0, 0]);
    expect(parseCellAddress('B5')).toEqual([4, 1]);
    expect(parseCellAddress('AA10')).toEqual([9, 26]);

    expect(() => validateCellRange('A1:B10')).not.toThrow();
    expect(() => validateCellRange('INVALID')).toThrow();
  });

  it('should reject invalid input and unknown goals', () => {
    expect(() => validateActivity('word', null)).toThrowError(/Activity must be an object/);
    expect(() => validateActivity('word', { assetBaseUrl: '' })).toThrowError(/theme must be LIGHT or DARK/);
    expect(() => validateActivity('word', { assetBaseUrl: '', theme: 'LIGHT' })).toThrowError(/instruction and startingDocument are required/);

    expect(() =>
      validateActivity('word', {
        assetBaseUrl: '',
        theme: 'LIGHT',
        instruction: 'Test',
        startingDocument: '<p>Hi</p>',
        goals: [{ type: 'unknownGoal' }],
      })
    ).toThrowError(/Unknown goal type: unknownGoal/);
  });

  it('should allow valid open-ended activities with extra top-level fields', () => {
    const act = {
      assetBaseUrl: 'https://cdn.example.com/',
      theme: 'DARK',
      instruction: 'Format document',
      startingDocument: '<p>Document content</p>',
      goals: [],
      customEditorFlag: true,
    };
    expect(() => validateActivity('word', act)).not.toThrow();
  });

  it('should validate assessment tasks requiring initial state but allowing null subsequent state', () => {
    const invalidAssessment = {
      assetBaseUrl: '',
      theme: 'LIGHT',
      mode: 'assessment',
      title: 'Excel Assessment',
      tasks: [
        {
          title: 'Task 1',
          explanation: 'Exp',
          instructions: ['Do step 1'],
          goals: [],
          startingSheet: null, // First task cannot have null startingSheet
        },
        {
          title: 'Task 2',
          explanation: 'Exp',
          instructions: ['Do step 2'],
          goals: [],
          startingSheet: null,
        },
      ],
    };
    expect(() => validateActivity('excel', invalidAssessment)).toThrowError(/startingSheet is required for the first task/);

    const validAssessment = {
      assetBaseUrl: '',
      theme: 'LIGHT',
      mode: 'assessment',
      title: 'Excel Assessment',
      tasks: [
        {
          title: 'Task 1',
          explanation: 'Exp',
          instructions: ['Do step 1'],
          goals: [],
          startingSheet: { cells: { A1: '10' } },
        },
        {
          title: 'Task 2',
          explanation: 'Exp',
          instructions: ['Do step 2'],
          goals: [],
          startingSheet: null, // Task 2 can continue previous task state
        },
      ],
    };
    expect(() => validateActivity('excel', validAssessment)).not.toThrow();
  });
});
