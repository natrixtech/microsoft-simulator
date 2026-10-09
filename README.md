# Word Processing Simulator (Microsoft Word 365)

An interactive, frontend-only Microsoft Word 365 learning simulator built with **Angular 16** and standalone components. Designed to run as an embedded iframe plugin inside LMS platforms and web training applications.

---

## 📋 Overview

- **Learner App (The Simulator):** A high-fidelity Microsoft Word 365 interface featuring the Ribbon UI, document canvas, paragraph styles, editing dialogs, comments margin gutter, variables side panel, and a responsive floating task mascot card.
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
