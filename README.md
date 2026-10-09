# Word Processing Simulator (Microsoft Word 365)

An interactive, frontend-only Microsoft Word 365 learning simulator built with **Angular 19** and standalone components. Designed to run as an embedded iframe plugin inside LMS platforms and web training applications.

---

## 📋 Overview

- **Learner App (The Simulator):** A high-fidelity Microsoft Word 365 interface featuring the Ribbon UI, document canvas, paragraph styles, editing dialogs, and a responsive floating task mascot card.
- **QA Test Harness:** A parent-window host environment (`test-harness.html`) simulating the LMS platform. It enables loading lessons, assessments, theme switching, and live `postMessage` contract monitoring.
- **Shared Architecture:** Depends on `projects/shared`, an internal Angular library housing the iframe communication bridge (`PluginBridgeService`), task mascot card component, and the animated "Show Me How" interactive demo component. *(Note: This shared library is architectural and will also be shared with Excel and PowerPoint simulators in future multi-app deployments).*

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
# or: npx ng serve word
```

Once compiled, access the applications at:
- **Learner Application (Simulator):** [http://localhost:4200/](http://localhost:4200/)
- **QA Test Harness (Host & Curriculum Runner):** [http://localhost:4200/test-harness.html](http://localhost:4200/test-harness.html)

---

## 📦 Building for Production

Compile the production-ready distribution bundles:
```bash
npm run build:word
```
The build output will be placed in:
```text
dist/word/browser/
├── index.html           # Main simulator plugin entry point
├── test-harness.html    # QA test harness for testing iframe integration
├── main-*.js            # Angular compiled application bundle
├── polyfills-*.js       # Browser polyfills
└── styles-*.css         # Word 365 styling bundle
```

To build both the shared library and the Word simulator:
```bash
npm run build:all
```

---

## 🧪 Testing

Run unit tests via Karma and Jasmine:
```bash
npm test
```

---

## 📂 Project Structure

```text
word-processing-simulator/
├── angular.json                 # Angular workspace configuration (shared & word)
├── package.json                 # Dependencies and build scripts
├── package-lock.json            # Deterministic dependency tree
├── tsconfig.json                # TypeScript path mappings & compiler options
├── .editorconfig                # Coding style settings
├── .gitignore                   # Git exclusion rules
├── README.md                    # Project documentation
├── docs/
│   ├── ARCHITECTURE.md          # System architecture & postMessage flow
│   ├── HOW_TO_ADD_A_LESSON.md   # Lesson schema & task goal types
│   ├── LICENSES.md              # Third-party dependency licenses
│   └── KNOWN_ISSUES.md          # Spec considerations and notes
└── projects/
    ├── shared/                  # Shared Angular Library
    │   ├── src/lib/
    │   │   ├── components/      # Task card, Show Me How demo modal
    │   │   ├── services/        # PluginBridgeService (iframe postMessage)
    │   │   ├── models/          # Lesson, activity, and goal type definitions
    │   │   └── validators/      # Activity JSON schema validators
    │   ├── ng-package.json
    │   └── tsconfig.lib.json
    └── word/                    # Word Processing Simulator
        ├── public/
        │   ├── assets/          # Application icons, mascots, and media
        │   └── test-harness.html# QA Test Harness page
        ├── src/
        │   ├── app/
        │   │   ├── services/    # RichTextService & WordDocumentService
        │   │   ├── app.component.ts
        │   │   ├── app.component.html
        │   │   └── app.component.css
        │   ├── index.html       # Simulator entry HTML
        │   ├── main.ts          # Angular bootstrap entry
        │   └── styles.css       # Global styles
        ├── tsconfig.app.json
        └── tsconfig.spec.json
```
