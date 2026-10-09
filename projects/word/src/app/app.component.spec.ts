import { TestBed } from '@angular/core/testing';
import { AppComponent } from './app.component';
import { LessonTask, TaskGoal } from 'shared';

function task(goals: TaskGoal[], hint?: string): LessonTask {
  return {
    title: 'Practice task',
    explanation: 'Do the thing.',
    instructions: ['Do the thing.'],
    hint,
    goals,
  };
}

describe('AppComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AppComponent],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(AppComponent);
    expect(fixture.componentInstance).toBeTruthy();
  });

  it('passes straight away when the task has no goals (Done button path)', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;

    app.currentActivity.set({
      mode: 'lesson',
      title: 'Practice',
      assetBaseUrl: './assets/',
      theme: 'LIGHT',
      tasks: [task([])],
    });
    app.taskIndex.set(0);
    app.onCheckWork();

    expect(app.isPassed()).toBeTrue();
    expect(app.feedbackMsg()).toBe('');
  });

  it('shows the hint and frowns when goals are not met', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    app.wordDoc.editorElement = document.createElement('div');
    app.wordDoc.editorElement.textContent = 'Hello world';

    app.currentActivity.set({
      mode: 'lesson',
      title: 'Practice',
      assetBaseUrl: './assets/',
      theme: 'LIGHT',
      tasks: [task([{ type: 'textContains', contains: 'missing-word' }], 'Try bolding the title.')],
    });
    app.taskIndex.set(0);
    app.onCheckWork();

    expect(app.isPassed()).toBeFalse();
    expect(app.feedbackMsg()).toBe('Try bolding the title.');
  });

  it('falls back to the default message when the task has no hint', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    app.wordDoc.editorElement = document.createElement('div');
    app.wordDoc.editorElement.textContent = 'Hello world';

    app.currentActivity.set({
      mode: 'lesson',
      title: 'Practice',
      assetBaseUrl: './assets/',
      theme: 'LIGHT',
      tasks: [task([{ type: 'textContains', contains: 'missing-word' }])],
    });
    app.taskIndex.set(0);
    app.onCheckWork();

    expect(app.feedbackMsg()).toBe('Not quite right. Try again.');
  });

  it('passes when every goal in the task is satisfied', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const editor = document.createElement('div');
    editor.innerHTML = '<p><b>My School</b> Apples Pears</p>';
    app.wordDoc.editorElement = editor;

    app.currentActivity.set({
      mode: 'lesson',
      title: 'Practice',
      assetBaseUrl: './assets/',
      theme: 'LIGHT',
      tasks: [
        task([
          { type: 'textContains', contains: ['My School', 'Apples'] },
          { type: 'textCount', text: 'Apples', count: 1 },
        ]),
      ],
    });
    app.taskIndex.set(0);
    app.onCheckWork();

    expect(app.isPassed()).toBeTrue();
  });

  it('marks the activity complete and sends one completed event on Continue', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const sent: boolean[] = [];
    spyOn(app.bridge, 'sendComplete').and.callFake((completed: boolean) => sent.push(completed));

    app.isPassed.set(true);
    app.onContinueNext();

    expect(app.isCompleted()).toBeTrue();
    expect(sent).toEqual([true]);
  });

  it('accepts a new PLUGIN_DATA_IN after finishing and fully resets for it', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const flush = async () => {
      fixture.detectChanges();
      await fixture.whenStable();
      fixture.detectChanges();
    };
    await flush(); // ngOnInit (bridge.init + effect) and ngAfterViewInit (fallback)

    const activity = (instruction: string, doc: string) => ({
      assetBaseUrl: './assets/',
      theme: 'LIGHT' as const,
      instruction,
      startingDocument: doc,
      goals: [{ type: 'textContains' as const, contains: 'Practice' }],
    });

    // Host sends activity 1 as PLUGIN_DATA_IN.
    app.bridge.loadActivity(activity('Do the thing', '<p>First document</p>'));
    await flush();
    expect(app.currentActivity()?.instruction).toBe('Do the thing');
    expect(app.isPassed()).toBeFalse();

    // The learner finishes it: card state + "already sent" flags are set.
    app.isPassed.set(true);
    app.isCompleted.set(true);
    app.feedbackMsg.set('Well done!');
    app.isDemoOpen.set(true); // e.g. a modal left open when the host moves on
    app.bridge.isCompleted.set(true);

    // Host answers with the NEXT activity as a separate PLUGIN_DATA_IN.
    app.bridge.loadActivity(activity('Next thing', '<p>Second document</p>'));
    await flush();

    expect(app.currentActivity()?.instruction).toBe('Next thing');
    expect(app.currentTask()?.instructions).toEqual(['Next thing']);
    expect(app.isPassed()).toBeFalse();
    expect(app.isCompleted()).toBeFalse();
    expect(app.feedbackMsg()).toBe('');
    expect(app.isDemoOpen()).toBeFalse();
    expect(app.showPageSetup()).toBeFalse();
    expect(app.showFindReplace()).toBeFalse();
    expect(app.bridge.isCompleted()).toBeFalse(); // can send activityCompleted again
    expect(app.wordDoc.editorElement?.innerHTML).toContain('Second document'); // document reloaded
  });
});
