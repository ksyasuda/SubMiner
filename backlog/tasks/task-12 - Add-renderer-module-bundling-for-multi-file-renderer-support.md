---
id: TASK-12
title: Add renderer module bundling for multi-file renderer support
status: To Do
assignee: []
created_date: '2026-02-11 08:21'
labels:
  - infrastructure
  - renderer
  - build
milestone: Codebase Clarity & Composability
dependencies:
  - TASK-5
references:
  - src/renderer/renderer.ts
  - src/renderer/index.html
  - package.json
  - tsconfig.json
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Currently renderer.ts is a single file loaded directly by Electron's renderer process via a script tag in index.html. To split it into modules (TASK-6), we need a bundling step since Electron renderer's default context doesn't support bare ES module imports without additional configuration.

Options:
1. **esbuild** — fast, minimal config, already used in many Electron projects
2. **Electron's native ESM support** — requires `"type": "module"` and sandbox configuration
3. **TypeScript compiler output** — if targeting a single concatenated bundle

The build pipeline already compiles TypeScript and copies renderer assets. Adding a bundling step for the renderer would slot into the existing `npm run build` script.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Renderer code can be split across multiple .ts files with imports
- [ ] #2 Build pipeline bundles renderer modules into a single output for Electron
- [ ] #3 Existing `make build` still works end-to-end
- [ ] #4 No runtime errors in renderer process
<!-- AC:END -->
