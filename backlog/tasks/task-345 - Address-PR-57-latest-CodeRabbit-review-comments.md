---
id: TASK-345
title: Address PR 57 latest CodeRabbit review comments
status: Done
assignee:
  - codex
created_date: '2026-05-12 06:35'
updated_date: '2026-05-12 06:38'
labels:
  - pr-review
  - coderabbit
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess the 2026-05-11 CodeRabbit review on PR #57 and address still-valid actionable comments with minimal changes. Current comments cover mpv subtitle playback unpause behavior, parser-selection empty reading handling, and overlay focus selector accessibility.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Still-valid CodeRabbit comments from the latest PR #57 review are fixed or explicitly documented as skipped with rationale.
- [x] #2 Regression coverage is added for behavior-affecting fixes where practical before production code changes.
- [x] #3 Relevant targeted checks pass locally.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect current code and nearby tests for the three latest CodeRabbit comments on PR #57.
2. Add regression tests first for behavior-impacting findings: mpv playNextSubtitle unpauses when pause state is unknown, and parser selection preserves combined reading for empty segment readings.
3. Apply minimal production fixes plus the CSS :focus-visible selector change.
4. Run targeted test commands for touched areas and update task notes/final status.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added red tests for the two behavior comments. `bun test src/core/services/mpv.test.ts` fails because playNextSubtitle skips unpause when pause state is null. `bun test src/core/services/tokenizer/parser-selection-stage.test.ts` fails because empty grammar-ending reading clears preceding combined reading.

Implemented all three latest CodeRabbit findings. Added regressions for mpv unknown pause state and parser-selection empty reading handling; changed overlay focus selectors to :focus-visible. Also fixed existing Prettier failure in src/core/services/stats-window.ts. Verification passed: targeted tests, typecheck, test:fast, test:env, format:check:src, build, test:smoke:dist, changelog:lint, changelog:pr-check.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed the latest CodeRabbit comments on PR #57.

Changes:
- `playNextSubtitle` now sends `pause=false` unless mpv is explicitly known to be playing, covering startup/reconnect unknown pause state.
- Parser selection no longer slices combined readings for empty grammar-ending readings, preserving preceding token readings.
- Overlay focus suppression now targets `:focus-visible` selectors.
- Applied Prettier to `src/core/services/stats-window.ts` to clear the existing formatting gate failure.

Verification:
- `bun test src/core/services/mpv.test.ts`
- `bun test src/core/services/tokenizer/parser-selection-stage.test.ts`
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run format:check:src`
- `bun run build`
- `bun run test:smoke:dist`
- `bun run changelog:lint`
- `bun run changelog:pr-check`
<!-- SECTION:FINAL_SUMMARY:END -->
