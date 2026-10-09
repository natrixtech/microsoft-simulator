# Word Processing Simulator (Angular 19 Plugin)

Independent web-based mock word processing simulator modelled on Google Docs / Microsoft Word standards, running as an iframe plugin.

## 1. Architecture & Design
- **Framework**: Angular 19 (Standalone Components, Signals, Reactive State)
- **Deployment**: Single Page Application designed to load inside a host `<iframe>`
- **Isolation**: Self-contained state, internal clipboard (no system clipboard dependencies), and in-iframe dialogs (no browser popups/alerts).
- **Theme**: Supports both LIGHT and DARK themes dynamically toggled via host postMessage.

## 2. Build Instructions
To build the Word Simulator independently:
```bash
cd simulators
# If using npm scripts:
npm run build:word

# Or via Angular CLI directly:
npx ng build word --base-href ./
```
The production bundle will be generated at `simulators/dist/word/browser/`.
- Entry point: `index.html`
- QA Test Harness: `test-harness.html`
- Base href: `./` (enables deployment under any arbitrary subpath or static CDN)

## 3. Local Serving & Testing
To preview locally:
```bash
# From workspace root:
npm run serve:angular
```
Open:
- Simulator: `http://127.0.0.1:4174/word/index.html`
- QA Test Harness: `http://127.0.0.1:4174/word/test-harness.html`

## 4. Iframe PostMessage Protocol

### Lifecycle
1. **App Ready**:
   When loaded and ready to receive activities:
   ```javascript
   window.parent.postMessage({ type: 'PLUGIN_EVENT', name: 'pluginReady' }, '*');
   ```

2. **Receive Activity**:
   Host sends activity payload:
   ```javascript
   window.postMessage({
     type: 'PLUGIN_DATA_IN',
     payload: {
       instruction: "Make the word 'really' bold.",
       startingDocument: "<p>I really like pizza.</p>",
       goals: [{ type: "textStyle", text: "really", bold: true }],
       assetBaseUrl: "https://example.com/assets/",
       theme: "LIGHT"
     }
   }, '*');
   ```

3. **Activity Completed**:
   When user meets all goals and clicks Continue (or Done when goals is empty):
   ```javascript
   window.parent.postMessage({
     type: 'PLUGIN_EVENT',
     name: 'activityCompleted',
     payload: { completed: true }
   }, '*');
   ```

4. **Host Controls**:
   - Pause: `{ type: 'PAUSE', payload: { pause: true } }`
   - Resume: `{ type: 'PAUSE', payload: { pause: false } }`
   - Force Finish: `{ type: 'FORCE_FINISH' }`
   - Theme Toggle: `{ type: 'PREFERENCE_CHANGED', payload: { type: 'theme', value: 'DARK' } }`

## 5. Supported Goal Types
- `textStyle`: bold, italic, underline, font, size (pt), color, highlight (checks first occurrence of text).
- `paragraph`: align, style (normal, title, h1-h3), indent level.
- `list`: items, listType (bullet, numbered), level.
- `textContains`: contains, notContains.
- `textCount`: text, count.
- `textOrder`: sequence of text strings.
- `image`: file, align, wrap, minWidth, maxWidth.
- `pageSetup`: orientation, paperSize, margins (top, bottom, left, right).
- `table`: rows, columns, cells.
- `element`: pageBreak, header, footer, pageNumbers, tableOfContents, footnote, comment, variable.
- `action`: undo, copy, cut, paste, select, find, openPageSetup, openFindReplace, tabNewRow.

## 6. Libraries and Licenses
- **@angular/core**, **@angular/common**, **@angular/forms**: MIT License
- **rxjs**: Apache 2.0 License
- **canvas-confetti**: ISC License
