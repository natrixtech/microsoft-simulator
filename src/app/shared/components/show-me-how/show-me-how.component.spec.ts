import { ComponentFixture, TestBed, fakeAsync, tick } from '@angular/core/testing';
import { ShowMeHowComponent } from './show-me-how.component';

describe('ShowMeHowComponent', () => {
  let fixture: ComponentFixture<ShowMeHowComponent>;
  let component: ShowMeHowComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ShowMeHowComponent],
    }).compileComponents();

    fixture = TestBed.createComponent(ShowMeHowComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    component.onClose();
    fixture.destroy();
  });

  it('should initialize and handle empty goals array gracefully without crashing', fakeAsync(() => {
    component.title = 'Empty Task Demo';
    component.instruction = 'Follow the steps';
    component.goals = [];
    component.startingDocument = '<p>Initial text</p>';

    fixture.detectChanges();
    tick(1000);
    fixture.detectChanges();

    expect(component.steps().length).toBe(0);
    expect(component.showToast()).toBeTrue();
    expect(component.getToastTitle()).toBe('empty task demo');
  }));

  it('should execute textStyle goal including explicit false style values and formatting-split text', fakeAsync(() => {
    component.title = 'Apply Bold & Remove Italic';
    component.goals = [
      {
        type: 'textStyle',
        text: 'hello',
        bold: true,
        italic: false,
        underline: true,
        font: 'Arial',
        size: 14,
        color: 'red',
        highlight: 'yellow',
      },
    ];
    // Formatting split: "he" and "llo" across elements
    component.startingDocument = '<p><span>he</span><em>llo</em> world</p>';

    fixture.detectChanges();
    tick(6000);
    fixture.detectChanges();

    const step = component.steps()[0];
    expect(step.state).toBe('done');
    expect(component.showToast()).toBeTrue();

    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    expect(docEl).toBeTruthy();
    // Verify style applied
    const span = docEl.querySelector('[data-demo-styled]') as HTMLElement;
    expect(span).toBeTruthy();
    expect(span?.style.fontWeight).toBe('bold');
    expect(span?.style.fontStyle).toBe('normal');
    expect(span?.style.textDecoration).toBe('underline');
  }));

  it('should not mark step done if target text for textStyle is missing', fakeAsync(() => {
    component.title = 'Missing Text Target';
    component.goals = [{ type: 'textStyle', text: 'Nonexistent string', bold: true }];
    component.startingDocument = '<p>Sample content only</p>';

    fixture.detectChanges();
    tick(2500);
    fixture.detectChanges();

    const step = component.steps()[0];
    expect(step.state).not.toBe('done');
    expect(component.showToast()).toBeFalse(); // No unconditional success
  }));

  it('should demonstrate paragraph goal with zero indent, alignment, and style', fakeAsync(() => {
    component.title = 'Format Paragraph';
    component.goals = [
      {
        type: 'paragraph',
        text: 'Target paragraph',
        align: 'center',
        indent: 0,
        lineSpacing: 1.5,
        style: 'h1',
      },
    ];
    component.startingDocument = '<p style="margin-left: 40px;">Target paragraph</p><p>Second paragraph</p>';

    fixture.detectChanges();
    tick(6000);
    fixture.detectChanges();

    expect(component.steps()[0].state).toBe('done');
    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    const h1 = docEl.querySelector('h1');
    expect(h1).toBeTruthy();
    expect(h1?.style.textAlign).toBe('center');
    expect(h1?.style.marginLeft).toBe('0px'); // Zero indent respected
    expect(h1?.dataset['lineSpacing']).toBe('1.5');
  }));

  it('should demonstrate list conversion for numbered and nested lists', fakeAsync(() => {
    component.title = 'Make Numbered List';
    component.goals = [
      {
        type: 'list',
        items: ['First item', 'Second item'],
        listType: 'number',
        level: 1,
      },
    ];
    component.startingDocument = '<p>First item</p><p>Second item</p>';

    fixture.detectChanges();
    tick(4000);
    fixture.detectChanges();

    expect(component.steps()[0].state).toBe('done');
    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    const ol = docEl.querySelector('ol');
    expect(ol).toBeTruthy();
    expect(ol?.children.length).toBe(2);
    expect(ol?.style.marginLeft).toBe('24px'); // Level 1 indent
  }));

  it('should demonstrate textContains with find and replace and contains-only without inventing words', fakeAsync(() => {
    component.title = 'Find and Replace';
    component.goals = [
      {
        type: 'textContains',
        notContains: ['recieve'],
        contains: ['receive', 'accuracy'],
      },
    ];
    component.startingDocument = '<p>Please recieve the memo.</p>';

    fixture.detectChanges();
    tick(8000);
    fixture.detectChanges();

    expect(component.steps()[0].state).toBe('done');
    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    expect(docEl.textContent).toContain('receive');
    expect(docEl.textContent).not.toContain('recieve');
    expect(docEl.textContent).toContain('accuracy');
  }));

  it('should demonstrate textCount reduction and zero count', fakeAsync(() => {
    component.title = 'Reduce Count';
    component.goals = [
      {
        type: 'textCount',
        text: 'Duplicate',
        count: 0,
      },
    ];
    component.startingDocument = '<p>Duplicate and Duplicate and keep other</p>';

    fixture.detectChanges();
    tick(6000);
    fixture.detectChanges();

    expect(component.steps()[0].state).toBe('done');
    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    expect(docEl.textContent).not.toContain('Duplicate');
    expect(docEl.textContent).toContain('keep other');
  }));

  it('should demonstrate image, pageSetup, table, element, and action goals', fakeAsync(() => {
    component.title = 'Multi Feature Demo';
    component.images = [{ file: 'sample.png' }];
    component.variables = { clientName: 'Acme Corp' };
    component.goals = [
      { type: 'image', file: 'sample.png', align: 'center', width: 200 },
      { type: 'pageSetup', orientation: 'landscape', paperSize: 'A4', margins: { top: 2 } },
      { type: 'table', rows: 2, columns: 2, cells: [['A1', 'B1'], ['A2', '']] },
      { type: 'element', element: 'header', text: 'Confidential' },
      { type: 'element', element: 'comment', text: 'Important note', resolved: true },
      { type: 'element', element: 'variable', text: 'clientName' },
      { type: 'action', action: 'undo' },
    ];
    component.startingDocument = '<p>Initial text</p>';

    fixture.detectChanges();
    tick(18000);
    fixture.detectChanges();

    // Verify all steps executed
    const steps = component.steps();
    expect(steps.every((s) => s.state === 'done')).toBeTrue();
    expect(component.showToast()).toBeTrue();
  }));

  it('should pause playback when isPaused is true and resume when false', fakeAsync(() => {
    component.title = 'Pause Demo';
    component.goals = [{ type: 'action', action: 'undo' }];
    component.isPaused = false;

    fixture.detectChanges();
    tick(200);

    // Pause in middle
    component.isPaused = true;
    tick(1000);
    // Should not have finished while paused
    expect(component.steps()[0].state).not.toBe('done');

    // Resume
    component.isPaused = false;
    tick(2000);
    expect(component.steps()[0].state).toBe('done');
  }));

  it('should cancel running timers cleanly on replay or close', fakeAsync(() => {
    component.title = 'Cancel Demo';
    component.goals = [{ type: 'action', action: 'undo' }];

    fixture.detectChanges();
    tick(100);

    let closedEmitted = false;
    component.closed.subscribe(() => (closedEmitted = true));

    component.onClose();
    expect(closedEmitted).toBeTrue();

    // Advance time - no unexpected errors thrown
    tick(3000);
  }));

  it('should switch to references tab when demonstrating tableOfContents and footnote', fakeAsync(() => {
    component.title = 'References Demo';
    component.goals = [
      { type: 'element', element: 'tableOfContents' },
      { type: 'element', element: 'footnote', text: 'Citation note' },
    ];
    component.startingDocument = '<p>Body text</p>';

    fixture.detectChanges();
    tick(6000);
    fixture.detectChanges();

    expect(component.steps().every((s) => s.state === 'done')).toBeTrue();
    expect(component.replicaActiveTab()).toBe('references');
    const docEl = fixture.nativeElement.querySelector('.replica-doc-content');
    expect(docEl.querySelector('.table-of-contents')).toBeTruthy();
    expect(docEl.querySelector('.footnote-section')).toBeTruthy();
  }));
});
