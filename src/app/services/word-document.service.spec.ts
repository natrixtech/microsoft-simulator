import { WordDocumentService } from './word-document.service';

describe('WordDocumentService Goal Checks', () => {
  let service: WordDocumentService;
  let container: HTMLElement;

  beforeEach(() => {
    service = new WordDocumentService();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    container.remove();
  });

  it('should evaluate textStyle goals (bold, italic, underline)', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Make bold',
      startingDocument: '<p><b>Important</b> text</p>',
      goals: [],
    };
    service.init(act, act, container);

    expect(service.check({ type: 'textStyle', text: 'Important', bold: true })).toBeTrue();
    expect(service.check({ type: 'textStyle', text: 'Important', italic: true })).toBeFalse();
  });

  it('should evaluate textContains, textCount, and textOrder', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Check words',
      startingDocument: '<p>Apples and Oranges and Apples</p>',
      goals: [],
    };
    service.init(act, act, container);

    expect(
      service.check({
        type: 'textContains',
        contains: ['Apples', 'Oranges'],
        notContains: ['Bananas'],
      })
    ).toBeTrue();

    expect(service.check({ type: 'textCount', text: 'Apples', count: 2 })).toBeTrue();
    expect(service.check({ type: 'textCount', text: 'Oranges', count: 1 })).toBeTrue();

    expect(service.check({ type: 'textOrder', sequence: ['Apples', 'Oranges', 'Apples'] })).toBeTrue();
    expect(service.check({ type: 'textOrder', sequence: ['Oranges', 'Bananas'] })).toBeFalse();
  });

  it('should evaluate paragraph styles, align and indent', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Format paragraph',
      startingDocument: '<h1 data-paragraph-style="h1" style="text-align: center; margin-left: 40px;">Title</h1>',
      goals: [],
    };
    service.init(act, act, container);

    expect(service.check({ type: 'paragraph', text: 'Title', style: 'h1' })).toBeTrue();
    expect(service.check({ type: 'paragraph', text: 'Title', align: 'center' })).toBeTrue();
    expect(service.check({ type: 'paragraph', text: 'Title', indent: 1 })).toBeTrue();
  });

  it('should evaluate list goals', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Create list',
      startingDocument: '<ul><li>First</li><li>Second</li></ul>',
      goals: [],
    };
    service.init(act, act, container);

    expect(
      service.check({
        type: 'list',
        listType: 'bullet',
        items: ['First', 'Second'],
      })
    ).toBeTrue();
  });

  it('should evaluate table goals', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Create table',
      startingDocument: '<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>',
      goals: [],
    };
    service.init(act, act, container);

    expect(service.check({ type: 'table', rows: 2, columns: 2 })).toBeTrue();
    expect(
      service.check({
        type: 'table',
        cells: [
          ['A', 'B'],
          ['C', 'D'],
        ],
      })
    ).toBeTrue();
  });

  it('should evaluate image goals', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Insert image',
      startingDocument: '<p><img src="chart.png" data-file="chart.png" data-align="center" data-wrap="wrap" width="250" /></p>',
      goals: [],
    };
    service.init(act, act, container);

    expect(
      service.check({
        type: 'image',
        file: 'chart.png',
        align: 'center',
        wrap: 'wrap',
        minWidth: 200,
        maxWidth: 300,
      })
    ).toBeTrue();
  });

  it('should evaluate pageSetup and element goals', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Setup page',
      startingDocument: '<p>Document</p>',
      pageSetup: {
        orientation: 'landscape' as const,
        paperSize: 'A4' as const,
        margins: { top: 1.5, bottom: 1.5, left: 1, right: 1 },
      },
      goals: [],
    };
    service.init(act, act, container);
    service.header = 'My Document Header';
    service.numbering = true;

    expect(
      service.check({
        type: 'pageSetup',
        orientation: 'landscape',
        paperSize: 'A4',
        margins: { top: 1.5, bottom: 1.5 },
      })
    ).toBeTrue();

    expect(service.check({ type: 'element', element: 'header', text: 'My Document' })).toBeTrue();
    expect(service.check({ type: 'element', element: 'pageNumbers' })).toBeTrue();
  });

  it('should inherit pageSetup from activity payload when task has no pageSetup', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Setup page at activity level',
      startingDocument: '<p>Document</p>',
      pageSetup: {
        orientation: 'landscape' as const,
        paperSize: 'Legal' as const,
        margins: { top: 2, bottom: 1.5, left: 1.25, right: 1.25 },
      },
      goals: [],
    };
    const taskWithoutPageSetup = {
      startingDocument: '<p>Document</p>',
    };
    service.init(taskWithoutPageSetup, act, container);

    expect(service.setup.paperSize).toBe('Legal');
    expect(service.setup.orientation).toBe('landscape');
    expect(service.setup.margins).toEqual({ top: 2, bottom: 1.5, left: 1.25, right: 1.25 });
  });

  it('should evaluate action goals', () => {
    const act = {
      assetBaseUrl: '',
      theme: 'LIGHT' as const,
      instruction: 'Action test',
      startingDocument: '<p>Action document</p>',
      goals: [],
    };
    service.init(act, act, container);
    service.logAction('select', 'Action');
    service.logAction('copy', 'Action');
    service.logAction('undo');
    service.logAction('selectImage', 'photo.png');

    expect(service.check({ type: 'action', action: 'select', text: 'Action' })).toBeTrue();
    expect(service.check({ type: 'action', action: 'copy', text: 'Action' })).toBeTrue();
    expect(service.check({ type: 'action', action: 'undo' })).toBeTrue();
    expect(service.check({ type: 'action', action: 'selectImage' })).toBeTrue();
    expect(service.check({ type: 'action', action: 'selectImage', text: 'photo.png' })).toBeTrue();
  });
});
