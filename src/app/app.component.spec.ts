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

  describe('Show Me How link and snapshot contract', () => {
    it('shows the Show me how link when showMeHow === true, but not when missing or false', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      // Case 1: showMeHow: true -> should have Show me how
      app.currentActivity.set({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Try formatting',
        startingDocument: '<p>Document</p>',
        showMeHow: true,
        tasks: [
          {
            title: 'Task With Demo',
            explanation: 'Exp',
            instructions: ['Do step'],
            goals: [],
            showMeHow: true,
          },
        ],
      });
      fixture.detectChanges();
      expect(app.currentTask()?.showMeHow).toBeTrue();
      const cardEl = fixture.nativeElement.querySelector('lib-task-card');
      expect(cardEl).toBeTruthy();
      const helpBtn = fixture.nativeElement.querySelector('.help-button');
      expect(helpBtn).toBeTruthy();

      // Case 2: showMeHow: false -> no link
      app.currentActivity.set({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Try formatting',
        startingDocument: '<p>Document</p>',
        tasks: [
          {
            title: 'Task Without Demo',
            explanation: 'Exp',
            instructions: ['Do step'],
            goals: [],
            showMeHow: false,
          },
        ],
      });
      fixture.detectChanges();
      expect(app.currentTask()?.showMeHow).toBeFalse();
      expect(fixture.nativeElement.querySelector('.help-button')).toBeNull();

      // Case 3: missing showMeHow -> no link
      app.currentActivity.set({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Try formatting',
        startingDocument: '<p>Document</p>',
        tasks: [
          {
            title: 'Task Omitted Demo',
            explanation: 'Exp',
            instructions: ['Do step'],
            goals: [],
          },
        ],
      });
      fixture.detectChanges();
      expect(app.currentTask()?.showMeHow).toBeFalsy();
      expect(fixture.nativeElement.querySelector('.help-button')).toBeNull();
    });

    it('ignores legacy hintGif and never enables the demo through hintGif alone', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      // Activity with hintGif and no showMeHow
      app.currentActivity.set({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Try formatting',
        startingDocument: '<p>Document</p>',
        hintGif: 'demo.gif',
        tasks: [
          {
            title: 'Task With HintGif',
            explanation: 'Exp',
            instructions: ['Do step'],
            goals: [],
          },
        ] as any,
      });
      fixture.detectChanges();

      expect(app.currentTask()?.showMeHow).toBeFalsy();
      expect(fixture.nativeElement.querySelector('.help-button')).toBeNull();
    });

    it('captures task-start snapshot for inherited startingDocument: null and preserves snapshot despite student edits', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        mode: 'assessment',
        title: 'Two Task Assessment',
        tasks: [
          {
            title: 'Task 1',
            explanation: 'Exp',
            instructions: ['Do task 1'],
            goals: [],
            startingDocument: '<p>Initial state task 1</p>',
            showMeHow: true,
          },
          {
            title: 'Task 2',
            explanation: 'Exp',
            instructions: ['Do task 2'],
            goals: [],
            startingDocument: null, // Inherits document from task 1
            showMeHow: true,
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      // Student edits task 1
      if (app.editorRef) {
        app.editorRef.nativeElement.innerHTML = '<p>Finished task 1 edits</p>';
        app.onEditorInput();
      }

      // Transition to task 2
      app.loadTaskAtIndex(1);
      fixture.detectChanges();
      await fixture.whenStable();

      // Task 2 snapshot should match document at the beginning of task 2
      expect(app.currentTaskSnapshot()?.documentHtml).toBe('<p>Finished task 1 edits</p>');

      // Student now makes edits in task 2
      if (app.editorRef) {
        app.editorRef.nativeElement.innerHTML = '<p>Student broke everything in task 2</p>';
        app.onEditorInput();
      }

      // Open demo modal
      app.openDemoModal();
      fixture.detectChanges();

      // Demo snapshot must still have the pristine task-2 starting state, NOT the student's breaking edits!
      expect(app.currentTaskSnapshot()?.documentHtml).toBe('<p>Finished task 1 edits</p>');
    });

    it('preserves an intentionally empty starting document without substituting pizza fallback', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Empty start',
        startingDocument: '',
        showMeHow: true,
        tasks: [
          {
            title: 'Empty document task',
            explanation: 'Exp',
            instructions: ['Type something'],
            goals: [],
            startingDocument: '',
            showMeHow: true,
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      expect(app.currentTaskSnapshot()?.documentHtml).toBe('');
      app.openDemoModal();
      expect(app.currentTaskSnapshot()?.documentHtml).toBe('');
    });

    it('preserves student state, undo history, action logs, and completion after opening and closing demo', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Format text',
        startingDocument: '<p>Student text</p>',
        showMeHow: true,
        tasks: [
          {
            title: 'Demo task',
            explanation: 'Exp',
            instructions: ['Format text'],
            goals: [],
            startingDocument: '<p>Student text</p>',
            showMeHow: true,
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      // Student performs edits and logs actions
      app.wordDoc.logAction('bold', 'Student text');
      const actionCountBefore = app.wordDoc.actions.length;
      const htmlBefore = app.editorRef?.nativeElement.innerHTML;

      // Open and close demo
      app.openDemoModal();
      expect(app.isDemoOpen()).toBeTrue();
      app.isDemoOpen.set(false);

      // Student state is completely intact
      expect(app.editorRef?.nativeElement.innerHTML).toBe(htmlBefore);
      expect(app.wordDoc.actions.length).toBe(actionCountBefore);
      expect(app.isCompleted()).toBeFalse();
      expect(app.isPassed()).toBeFalse();
    });
  });
});

