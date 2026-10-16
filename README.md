# Word Processing Simulator

A Microsoft Word simulator that runs one activity at a time inside an iframe and communicates with the host through `window.postMessage`. Built with Angular 16.

## Requirements

- Node.js 18.10 or newer (tested with Node 22)
- npm 10 or newer

## Run

```bash
npm install
npm start
```

Open the test harness: http://localhost:4200/test-harness.html

The harness loads the app in an iframe, sends sample Input objects (one lesson per goal type, plus a 3-task assessment) and shows the `activityCompleted` result it receives. It also has Light/Dark, Pause/Resume and Force Finish controls, and a box for sending custom Input JSON.

Opening http://localhost:4200/ directly only shows "Waiting for an activity", because the app waits for the host to send `PLUGIN_DATA_IN`. The editor is chosen by the top-level `"app"` field of the Input object (currently `"word"`).

## Build

```bash
npm run build
```

The production build is written to `dist`: a folder with `index.html` and its files, loadable from any URL path. All code, fonts and images are bundled; nothing is loaded from external CDNs at runtime. `dist` contains the license notices of all bundled libraries.

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
