---
id: TASK-152
title: Fix early Electron userData path casing to stay under SubMiner config dir
status: Done
assignee:
  - codex
created_date: '2026-03-10 06:46'
updated_date: '2026-03-10 06:51'
labels:
  - bug
  - config
  - electron
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main-entry.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - /home/sudacode/projects/japanese/SubMiner/src/config/path-resolution.ts
documentation:
  - /home/sudacode/projects/japanese/subminer-docs/development.md
  - /home/sudacode/projects/japanese/subminer-docs/architecture.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Electron startup touches the app object before the main-process bootstrap overrides userData, which can create/write the default lowercase ~/.config/subminer directory on Linux/macOS. Ensure early startup pins the app identity and userData path to the canonical SubMiner config directory before any Electron APIs can materialize the default path, and keep regression coverage around the bootstrap path behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Electron startup uses the canonical SubMiner config directory as userData before other Electron app calls can create the default lowercase directory.
- [x] #2 Regression tests cover the early bootstrap path setup and fail if startup falls back to a lowercase subminer config path.
- [x] #3 Existing config path resolution behavior for SubMiner casing remains intact.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression coverage for early Electron bootstrap path setup, including a case that would otherwise fall back to lowercase subminer.
2. Extract a pure helper that computes and applies the canonical app name/userData path from config resolution.
3. Call the helper from main-entry before any Electron app interactions that could materialize the default userData directory.
4. Run focused tests for startup/config path behavior, then the relevant fast gate if green.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added early Electron bootstrap path setup in src/main-entry so app name and userData are pinned to the canonical SubMiner config dir before single-instance/whenReady handling.

Added a regression test in src/main-entry-runtime.test.ts covering the lowercase subminer fallback case.

Validation: bun test src/main-entry-runtime.test.ts src/config/path-resolution.test.ts; bun run typecheck. bun run test:fast still fails on an existing unrelated renderer JLPT CSS test in src/renderer/subtitle-render.test.ts.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Pinned Electron's app identity and userData path during entry bootstrap so startup uses the canonical SubMiner config directory before any other Electron app calls can materialize the default lowercase path. Added a regression test covering the lowercase subminer fallback case and kept existing config-path resolution coverage green.

Validation:
- bun test src/main-entry-runtime.test.ts src/config/path-resolution.test.ts
- bun run typecheck
- bun run test:fast (fails on existing unrelated JLPT CSS renderer test in src/renderer/subtitle-render.test.ts)
<!-- SECTION:FINAL_SUMMARY:END -->
