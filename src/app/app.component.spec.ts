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

    it('logs undo action when formatCmd("undo") is executed', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Undo test',
        startingDocument: '<p>Initial text</p>',
        tasks: [
          {
            title: 'Undo task',
            explanation: 'Undo test',
            instructions: ['Click undo'],
            goals: [{ type: 'action', action: 'undo' }],
            startingDocument: '<p>Initial text</p>',
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      expect(app.wordDoc.check({ type: 'action', action: 'undo' })).toBeFalse();
      app.formatCmd('undo');
      expect(app.wordDoc.check({ type: 'action', action: 'undo' })).toBeTrue();
    });

    it('logs selectImage action when an image in the editor is clicked', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Select image test',
        startingDocument: '<p><img src="./assets/photo.png" data-file="photo.png" /></p>',
        tasks: [
          {
            title: 'Select image task',
            explanation: 'Select image test',
            instructions: ['Click image'],
            goals: [{ type: 'action', action: 'selectImage' }],
            startingDocument: '<p><img src="./assets/photo.png" data-file="photo.png" /></p>',
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      expect(app.wordDoc.check({ type: 'action', action: 'selectImage' })).toBeFalse();
      const img = app.editorRef?.nativeElement.querySelector('img');
      expect(img).toBeTruthy();
      app.onEditorClick({ target: img } as unknown as MouseEvent);
      expect(app.wordDoc.check({ type: 'action', action: 'selectImage' })).toBeTrue();
    });

    it('opens in-app comment dialog and allows confirming comment without native prompt', async () => {
      const fixture = TestBed.createComponent(AppComponent);
      const app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();

      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Comment test',
        startingDocument: '<p>Document section for review</p>',
        tasks: [
          {
            title: 'Comment task',
            explanation: 'Add comment',
            instructions: ['Insert comment'],
            goals: [
              { type: 'element', element: 'comment', text: 'Important note' },
              { type: 'action', action: 'comment' },
            ],
            startingDocument: '<p>Document section for review</p>',
          },
        ],
      });
      fixture.detectChanges();
      await fixture.whenStable();

      expect(app.showCommentDialog()).toBeFalse();

      // Trigger Insert Comment
      app.insertComment();
      fixture.detectChanges();

      expect(app.showCommentDialog()).toBeTrue();
      expect(app.wordDoc.comments.length).toBe(0);

      // Verify cancel does not add comment
      app.cancelInsertComment();
      fixture.detectChanges();
      expect(app.showCommentDialog()).toBeFalse();
      expect(app.wordDoc.comments.length).toBe(0);

      // Re-open and confirm comment
      app.insertComment();
      app.newCommentText.set('Important note');
      app.confirmInsertComment();
      fixture.detectChanges();

      expect(app.showCommentDialog()).toBeFalse();
      expect(app.wordDoc.comments.length).toBe(1);
      expect(app.wordDoc.comments[0].comment).toBe('Important note');
      expect(app.wordDoc.check({ type: 'element', element: 'comment', text: 'Important note' })).toBeTrue();
      expect(app.wordDoc.check({ type: 'action', action: 'comment' })).toBeTrue();
    });
  });

  describe('Deletion of all inserted document content types (Bugfix verification)', () => {
    let fixture: any;
    let app: AppComponent;

    beforeEach(async () => {
      fixture = TestBed.createComponent(AppComponent);
      app = fixture.componentInstance;
      fixture.detectChanges();
      await fixture.whenStable();
    });

    it('1. Image: can be selected and deleted via Backspace/Delete or delete button', () => {
      app.loadWordActivity({
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        instruction: 'Image test',
        startingDocument: '<p>Initial</p>',
        tasks: [{ title: 'T', explanation: 'E', instructions: ['I'], goals: [] }],
      });
      fixture.detectChanges();

      const img = document.createElement('img');
      img.src = './assets/books.svg';
      app.editorRef?.nativeElement.appendChild(img);
      expect(app.editorRef?.nativeElement.querySelector('img')).toBeTruthy();

      // Select image
      app.onEditorClick({ target: img } as unknown as MouseEvent);
      expect(app.selectedImageEl()).toBe(img);

      // Delete with Backspace/Delete key
      app.onDocumentKeydown(new KeyboardEvent('keydown', { key: 'Backspace' }));
      expect(app.editorRef?.nativeElement.querySelector('img')).toBeFalsy();
      expect(app.selectedImageEl()).toBeNull();
    });

    it('2. Table: can be selected and deleted', () => {
      app.insertGridTable(2, 2);
      const table = app.editorRef?.nativeElement.querySelector('table');
      expect(table).toBeTruthy();

      // Select table
      app.selectedTableEl.set(table || null);
      app.deleteSelectedTable();
      expect(app.editorRef?.nativeElement.querySelector('table')).toBeFalsy();
      expect(app.selectedTableEl()).toBeNull();
    });

    it('2b. Table typing safeguard: backspacing mid-text in a non-empty cell does not delete the table', () => {
      app.insertGridTable(2, 2);
      const table = app.editorRef?.nativeElement.querySelector('table') as HTMLTableElement;
      expect(table).toBeTruthy();
      const cell = table.rows[0].cells[0];
      cell.textContent = 'Testing';

      // Position cursor mid-text (offset 4, after 'Test', before 'ing')
      const textNode = cell.firstChild!;
      const sel = window.getSelection()!;
      const range = document.createRange();
      range.setStart(textNode, 4);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);

      // Even if table is marked as selectedTableEl from prior click
      app.selectedTableEl.set(table);

      // Backspace event dispatched
      const event = new KeyboardEvent('keydown', { key: 'Backspace', cancelable: true });
      app.onDocumentKeydown(event);
      app.onEditorKeydown(event);

      // Must NOT prevent default (allowing native deletion of just 1 character)
      expect(event.defaultPrevented).toBeFalse();

      // Table must still exist in DOM
      expect(app.editorRef?.nativeElement.querySelector('table')).toBeTruthy();
    });

    it('3. Page Break: can be deleted, merging pages back together', () => {
      app.editorRef!.nativeElement.innerHTML = '<p>Page 1 Text</p>';
      app.insertPageBreak();
      expect(app.hasPageBreak()).toBeTrue();

      // Backspace at beginning of page 2 merges back
      app.mergePage2IntoPage1();
      expect(app.hasPageBreak()).toBeFalse();
      expect(app.editorRef?.nativeElement.querySelector('[data-element="pageBreak"]')).toBeFalsy();
    });

    it('4. Header: can be cleared or removed', () => {
      app.toggleHeader();
      expect(app.wordDoc.header).toBe('Document Header');

      app.removeHeader();
      expect(app.wordDoc.header).toBe('');
      expect(app.activeRegion()).toBe('body');
    });

    it('5. Footer: can be cleared or removed', () => {
      app.toggleFooter();
      expect(app.wordDoc.footer).toBe('Document Footer');

      app.removeFooter();
      expect(app.wordDoc.footer).toBe('');
      expect(app.activeRegion()).toBe('body');
    });

    it('6. Page Numbers: can be selected and deleted with Backspace/Delete', () => {
      app.togglePageNumbers();
      expect(app.wordDoc.numbering).toBeTrue();

      app.selectedPageNum.set(true);
      app.onDocumentKeydown(new KeyboardEvent('keydown', { key: 'Delete' }));
      expect(app.wordDoc.numbering).toBeFalse();
      expect(app.selectedPageNum()).toBeFalse();
    });

    it('7. Table of Contents: is editable and can be deleted via Backspace/Delete or remove button', () => {
      app.editorRef!.nativeElement.innerHTML = '<h2>Title 1</h2><p>Content</p>';
      app.insertTableOfContents();
      const toc = app.editorRef?.nativeElement.querySelector('.table-of-contents') as HTMLElement;
      expect(toc).toBeTruthy();
      expect(toc.getAttribute('contenteditable')).not.toBe('false');

      // Click remove button or press Backspace when selected
      app.selectedTocEl.set(toc);
      app.onDocumentKeydown(new KeyboardEvent('keydown', { key: 'Backspace' }));
      expect(app.editorRef?.nativeElement.querySelector('.table-of-contents')).toBeFalsy();
    });

    it('8. Footnotes: footnote section is cleaned up when footnote reference is deleted', () => {
      app.editorRef!.nativeElement.innerHTML = '<p>Some text</p>';
      app.insertFootnote();
      const fnRef = app.editorRef?.nativeElement.querySelector('.footnote-ref');
      const fnSec = app.editorRef?.nativeElement.querySelector('.footnote-section');
      expect(fnRef).toBeTruthy();
      expect(fnSec).toBeTruthy();

      // Delete footnote ref with Backspace
      fnRef?.remove();
      app.onEditorInput();
      expect(app.editorRef?.nativeElement.querySelector('.footnote-section')).toBeFalsy();
    });

    it('9. Comments: can be deleted from the document', () => {
      app.wordDoc.comments = [];
      app.insertComment('A test comment', 'Target text');
      expect(app.wordDoc.comments.length).toBe(1);

      app.deleteComment(0);
      expect(app.wordDoc.comments.length).toBe(0);
    });

    it('10. Variables: empty variable tokens are cleaned up', () => {
      app.insertVariable('company', 'Acme');
      const v = app.editorRef?.nativeElement.querySelector('[data-element="variable"]');
      expect(v).toBeTruthy();

      // Clear text
      v!.textContent = '';
      app.onEditorInput();
      expect(app.editorRef?.nativeElement.querySelector('[data-element="variable"]')).toBeFalsy();
    });

    it('11. Lists: empty lists are cleaned up on editor input', () => {
      app.editorRef!.nativeElement.innerHTML = '<ul><li></li></ul>';
      app.onEditorInput();
      expect(app.editorRef?.nativeElement.querySelector('ul')).toBeFalsy();
    });

    it('12. Force Finish immediately sends activityCompleted with current progress', () => {
      spyOn(app.bridge, 'sendComplete');
      app.loadWordActivity({
        mode: 'assessment',
        title: 'Multi-Task Assessment',
        assetBaseUrl: './assets/',
        theme: 'LIGHT',
        tasks: [
          { title: 'Task 1', explanation: 'E1', instructions: ['I1'], goals: [] },
          { title: 'Task 2', explanation: 'E2', instructions: ['I2'], goals: [] },
        ],
      });
      fixture.detectChanges();

      // Simulate student having passed 1 out of 2 tasks
      app.passedTasks.set(new Set<number>([0]));

      // Trigger FORCE_FINISH
      app.bridge.forceFinishTrigger.update(v => v + 1);
      fixture.detectChanges();

      // Must call sendComplete with completed=false, tasksPassed=1, tasksTotal=2
      expect(app.bridge.sendComplete).toHaveBeenCalledWith(false, 1, 2);
      expect(app.isCompleted()).toBeTrue();
    });
  });
});

