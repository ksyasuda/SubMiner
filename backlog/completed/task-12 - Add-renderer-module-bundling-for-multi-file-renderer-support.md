---
id: TASK-12
title: Add renderer module bundling for multi-file renderer support
status: Done
assignee: []
created_date: '2026-02-11 08:21'
updated_date: '2026-02-16 02:14'
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
priority: high
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

- [x] #1 Renderer code can be split across multiple .ts files with imports
- [x] #2 Build pipeline bundles renderer modules into a single output for Electron
- [x] #3 Existing `make build` still works end-to-end
- [x] #4 No runtime errors in renderer process
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Updated root npm build pipeline to use an explicit renderer bundle step via esbuild. Added `build:renderer` script to emit a single `dist/renderer/renderer.js` from `src/renderer/renderer.ts`; `build` now runs `pnpm run build:renderer` and preserves existing index/style copy and macOS helper step. Added `esbuild` to devDependencies.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented renderer bundling step and wired `build` to use it. This adds `pnpm run build:renderer` which bundles `src/renderer/renderer.ts` into a single `dist/renderer/renderer.js` for Electron to load. Also added `esbuild` as a dev dependency and aligned `pnpm-lock.yaml` importer metadata for dependency consistency. Kept `index.html`/`style.css` copy path unchanged, so renderer asset layout remains stable.

Implemented additional test-layer type fix after build breakage by correcting `makeDepsFromMecabTokenizer` and related `tokenizeWithMecab` mocks to match expected `Token` vs `MergedToken` shapes, keeping runtime behavior unchanged while satisfying TS checks.

<!-- SECTION:FINAL_SUMMARY:END -->
