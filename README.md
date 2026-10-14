# Word Processing Simulator (Microsoft Word 365)

An interactive, frontend-only Microsoft Word 365 learning simulator built with **Angular 16** and standalone components. Designed to run as an embedded iframe plugin inside LMS platforms and web training applications.

---

## 📋 Overview

- **Learner App (The Simulator):** A high-fidelity Microsoft Word 365 interface featuring the Ribbon UI, document canvas, paragraph styles, editing dialogs, comments margin gutter, variables dialog, and a responsive floating task mascot card.
- **QA Test Harness:** A parent-window host environment (`test-harness.html`) simulating the LMS platform. It enables loading lessons, assessments, theme switching, and live `postMessage` contract monitoring.
- **Shared Application Architecture:** Normal application modules under `src/app/shared/` containing the iframe communication bridge (`PluginBridgeService`), task mascot card component, and the animated "Show Me How" interactive demo component with coherent scaling.

---

## 🛠️ System Requirements

- **Node.js:** v20.x or v22.x (LTS recommended)
- **npm:** v10.x or higher
- **Modern Web Browser:** Chrome, Edge, Firefox, or Safari

---

## 🚀 Getting Started

### 1. Installation
Clone the repository and install all dependencies:
```bash
npm install
```

### 2. Development Server
Start the Angular development server:
```bash
npm start
# or: npx ng serve
```

Once compiled, access the applications at:
- **Learner Application (Simulator):** [http://localhost:4200/](http://localhost:4200/)
- **QA Test Harness (Host & Curriculum Runner):** [http://localhost:4200/test-harness.html](http://localhost:4200/test-harness.html)

---

## 📦 Building for Production

Compile the production-ready distribution bundles:
```bash
npm run build
# or: npm run build:word
```
The build output will be placed in:
```text
dist/word/
├── index.html           # Main simulator plugin entry point
├── test-harness.html    # QA test harness for testing iframe integration
├── favicon.ico          # Application favicon
├── assets/              # Mascots, icons, and SVG resources
├── main.*.js            # Angular compiled application bundle
├── polyfills.*.js       # Browser polyfills
└── styles.*.css         # Word 365 styling bundle
```

---

## 🧪 Testing

Run all unit tests via Karma and Jasmine:
```bash
npm test
# or: npx ng test --watch=false --browsers=ChromeHeadless
```

---

## 📂 Project Structure

```text
word-processing-simulator/
├── angular.json                 # Angular application configuration
├── package.json                 # Angular 16 dependencies and scripts
├── package-lock.json            # Deterministic dependency tree
├── tsconfig.json                # TypeScript compiler options & shared path mappings
├── tsconfig.app.json            # Application build TypeScript configuration
├── tsconfig.spec.json           # Karma test runner TypeScript configuration
├── .editorconfig                # Coding style settings
├── .gitignore                   # Git exclusion rules
├── README.md                    # Project documentation
├── public/                      # Static assets copied directly to root dist
│   ├── favicon.ico
│   ├── test-harness.html        # QA Test Harness runner
│   └── assets/                  # Robot mascots (happy, neutral, sad) & app icons
└── src/                         # Application source code
    ├── index.html               # Simulator entry point HTML
    ├── main.ts                  # Angular bootstrap entry
    ├── styles.css               # Global application styles
    ├── styles/
    │   └── theme.css            # Word 365 design tokens & dialog themes
    └── app/
        ├── app.component.ts     # Word editor controller
        ├── app.component.html   # Word 365 UI template (Ribbon, Document, Comments)
        ├── app.component.css    # Word 365 component styles
        ├── app.component.spec.ts
        ├── app.config.ts        # Application configuration providers
        ├── word-lessons.data.ts # Lesson & assessment definitions
        ├── services/            # Word core services
        │   ├── rich-text.service.ts
        │   ├── word-document.service.ts
        │   └── word-document.service.spec.ts
        └── shared/              # Shared application components & utilities
            ├── index.ts         # Public exports barrel
            ├── models/
            │   └── activity.types.ts
            ├── services/
            │   ├── confetti.service.ts
            │   ├── plugin-bridge.service.ts
            │   └── theme.service.ts
            ├── utils/
            │   └── editor.utils.ts
            ├── validators/
            │   ├── contract.validator.ts
            │   └── contract.validator.spec.ts
            └── components/
                ├── mascot/
                │   ├── mascot.component.ts
                │   ├── mascot.component.html
                │   └── mascot.component.css
                ├── task-card/
                │   ├── task-card.component.ts
                │   ├── task-card.component.html
                │   └── task-card.component.css
                └── show-me-how/
                    ├── show-me-how.component.ts
                    ├── show-me-how.component.html
                    ├── show-me-how.component.css
                    └── show-me-how.component.spec.ts
```

---

## 📜 Third-Party Dependencies & Licenses

All third-party libraries and direct dependencies utilized in this project are open-source and permissible for commercial embedding within learning management systems:

| Package | Version | License | Purpose | Source / Repository |
|---|---|---|---|---|
| `canvas-confetti` | `^1.9.4` | MIT | Celebratory confetti particle animations on task completion | [GitHub](https://github.com/catdad/canvas-confetti) |
| `@angular/animations` | `^16.2.12` | MIT | Animation engine and transitions primitives | [GitHub](https://github.com/angular/angular) |
| `@angular/common` | `^16.2.12` | MIT | Common Angular directives, pipes, and utilities | [GitHub](https://github.com/angular/angular) |
| `@angular/compiler` | `^16.2.12` | MIT | Ahead-of-Time and JIT template compiler | [GitHub](https://github.com/angular/angular) |
| `@angular/core` | `^16.2.12` | MIT | Application framework & reactive signal primitives | [GitHub](https://github.com/angular/angular) |
| `@angular/forms` | `^16.2.12` | MIT | Two-way data binding and form control primitives | [GitHub](https://github.com/angular/angular) |
| `@angular/platform-browser` | `^16.2.12` | MIT | Browser rendering and DOM management | [GitHub](https://github.com/angular/angular) |
| `@angular/platform-browser-dynamic` | `^16.2.12` | MIT | Browser bootstrap provider for JIT compilation | [GitHub](https://github.com/angular/angular) |
| `@angular/router` | `^16.2.12` | MIT | Client-side routing and URL state management | [GitHub](https://github.com/angular/angular) |
| `rxjs` | `~7.8.0` | Apache-2.0 | Reactive asynchronous event handling and stream primitives | [GitHub](https://github.com/ReactiveX/rxjs) |
| `tslib` | `^2.3.0` | 0BSD | Runtime TypeScript helper library | [GitHub](https://github.com/microsoft/tslib) |
| `zone.js` | `~0.13.3` | MIT | Zone execution context for change detection tracking | [GitHub](https://github.com/angular/angular) |
| `Lexend` | Variable (100–900) | SIL OFL 1.1 | Modern typography for lesson task card UI | [Google Fonts / GitHub](https://github.com/googlefonts/lexend) |

---

## 🏛️ Architecture & Engineering Decisions

### Rich Text Editing Engine (`RichEditor` & `document.execCommand`)
The document editor in `src/app/services/rich-text.service.ts` is implemented using browser DOM APIs and `document.execCommand` paired with custom DOM/Range tree walkers (`findText`, `styleMatches`), rather than an external heavyweight editor library (such as Quill, ProseMirror, or TipTap).

This deliberate architectural decision provides key operational advantages:
1. **Zero Bundle Overhead:** Keeps the production bundle under ~100 KB gzipped, ensuring near-instantaneous load times inside sandboxed host iframes.
2. **Direct DOM & Contract Alignment:** Generates clean, native semantic HTML elements (`<p>`, `<table>`, `<b>`, `<i>`, `<span>`, CSS styles, and `dataset` attributes) that map 1:1 to the 11 Word activity validation goal evaluators (`wordDoc.check(g)`) without intermediary virtual document models or AST abstractions.
3. **Cross-Browser Stability:** Fully supported across modern Chromium (Chrome, Edge), Firefox, and WebKit (Safari) engines.
4. **Encapsulated Abstraction Layer:** The editor logic is strictly encapsulated behind the `RichEditor` class. If headless schema constraints (e.g., TipTap/ProseMirror) are desired in future milestones, the editor engine can be swapped behind this service interface without altering the surrounding Ribbon UI, task card, or plugin bridge contract.
