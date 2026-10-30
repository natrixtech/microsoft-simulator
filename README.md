# Word, Excel and PowerPoint Simulators

Microsoft Word, Microsoft Excel and Microsoft PowerPoint simulators that run one activity at a time inside an iframe and communicate with the host through `window.postMessage`. Built with Angular 16.

## Requirements

- Node.js 18.10 or newer (tested with Node 22)
- npm 10 or newer

## Run

```bash
npm install
npm start
```

Open the test harnesses:

- Word: http://localhost:4200/word-test-harness.html
- Excel: http://localhost:4200/excel-test-harness.html
- PowerPoint: http://localhost:4200/powerpoint-test-harness.html
- All three (links page): http://localhost:4200/test-harness.html

Each harness loads the app in an iframe, sends sample Input objects and shows the `activityCompleted` result it receives. The Word harness has one lesson per goal type plus a 3-task assessment. The Excel harness has 32 lessons covering every goal type. The PowerPoint harness has 7 lessons with 2–3 steps each plus a 3-task assessment; together they cover all 14 goal types and all 5 action types. All harnesses have Light/Dark, Pause/Resume and Force Finish controls, and a box for sending custom Input JSON.

Opening http://localhost:4200/ directly only shows "Waiting for an activity", because the app waits for the host to send `PLUGIN_DATA_IN`. The editor is chosen by the top-level `"app"` field of the Input object: `"word"`, `"excel"` or `"powerpoint"`.

## Build

```bash
npm run build
```

The production build is written to `dist`: a folder with `index.html` and its files, loadable from any URL path. All code, fonts and images are bundled; nothing is loaded from external CDNs at runtime. `dist/3rdpartylicenses.txt` contains the license notices of all bundled libraries.

## Spreadsheet and formula engine

The Excel simulator does not use any third-party spreadsheet or formula library. The grid, the formula parser and all supported functions are written as part of this project, so no extra license applies.

## Slide editor

The PowerPoint simulator does not use any third-party presentation or slide library. The slide editor, themes, layouts, shapes, tables, transition preview and the "Show me how" demo are written as part of this project, so no extra license applies. Slide text uses standard system fonts (Arial, Georgia, Times New Roman, Verdana, Courier New, Comic Sans MS); no font files are bundled for them.

## Libraries and licenses

| Library | Version | License |
|---|---|---|
| Angular (`@angular/animations`, `common`, `compiler`, `core`, `forms`, `platform-browser`, `platform-browser-dynamic`, `router`) | 16.2 | MIT |
| `canvas-confetti` | 1.9 | MIT |
| `rxjs` | 7.8 | Apache-2.0 |
| `tslib` | 2.x | 0BSD |
| `zone.js` | 0.13 | MIT |
| Lexend font (bundled) | variable | SIL Open Font License 1.1 |

Development-only tools (not part of the build): Angular CLI and build tools (MIT), TypeScript (Apache-2.0).