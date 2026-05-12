---
id: TASK-338
title: Fix known-word highlight on standalone subtitle particles
status: Done
assignee:
  - codex
created_date: '2026-05-04 05:52'
updated_date: '2026-05-04 05:57'
labels:
  - bug
  - subtitle
  - tokenizer
dependencies: []
references:
  - src/core/services/tokenizer/annotation-stage.ts
  - src/core/services/tokenizer/subtitle-annotation-filter.ts
  - src/renderer/subtitle-render.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Standalone grammar particles such as に should not render as known-word green when they appear in the known-word cache as readings for other words. Keep known-word coloring for lexical tokens, but prevent grammar-excluded subtitle tokens from getting known-green.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Standalone grammar particles like に do not retain isKnown after subtitle annotation filtering.
- [x] #2 Lexical known-word tokens still render as known when not grammar-excluded.
- [x] #3 Focused regression test covers the particle false-positive path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression in `src/core/services/tokenizer/annotation-stage.test.ts` showing standalone particle `に` is grammar-excluded and does not retain `isKnown` even when `isKnownWord('に')` is true.
2. Run the focused tokenizer annotation test and confirm the new test fails for the current behavior.
3. Patch `src/core/services/tokenizer/annotation-stage.ts` so grammar-excluded tokens clear known status while still stripping N+1/frequency/JLPT/name metadata.
4. Run the focused test file, then inspect diff and update task acceptance criteria.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented tokenizer annotation filtering so grammar-excluded subtitle tokens clear known-word status instead of retaining green known coloring. Added focused regression for known-word-cache particle false positive and updated existing expectations for unified annotation clearing. Verification: `bun test src/core/services/tokenizer/annotation-stage.test.ts --test-name-pattern "clears known status from standalone particles"` failed before the production patch; after patch, `bun test src/core/services/tokenizer/annotation-stage.test.ts`, `bun test src/core/services/tokenizer.test.ts`, combined tokenizer tests, `bun run typecheck`, `bun run changelog:lint`, and `bun run test:fast` passed.

Full handoff gate follow-up: `bun run test:env` and `bun run build` passed. `bun run test:smoke:dist` failed outside this tokenizer change in `dist/core/services/overlay-manager.test.js` because current dirty overlay-window code calls `window.getTitle()` on a test mock that does not provide it.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Cleared `isKnown` for grammar-excluded subtitle tokens in the tokenizer annotation stage, preventing standalone particles such as `に` from rendering as known just because a known-word deck contains a matching reading.
- Added a focused regression test for the known-word-cache false positive and updated tokenizer expectations so helper/grammar spans consistently clear all subtitle annotations.
- Added changelog fragment `changes/338-known-word-particle-highlights.md`.

Verification:
- `bun test src/core/services/tokenizer/annotation-stage.test.ts --test-name-pattern "clears known status from standalone particles"` failed before the production patch.
- `bun test src/core/services/tokenizer/annotation-stage.test.ts`
- `bun test src/core/services/tokenizer.test.ts`
- `bun test src/core/services/tokenizer/annotation-stage.test.ts src/core/services/tokenizer.test.ts`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`

Blocked/External:
- `bun run test:smoke:dist` currently fails outside this tokenizer change in `dist/core/services/overlay-manager.test.js`: dirty overlay-window code calls `window.getTitle()` on a test mock without that method.
<!-- SECTION:FINAL_SUMMARY:END -->
