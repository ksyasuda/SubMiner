---
id: TASK-274
title: Stabilize current failing test regressions
status: Done
assignee:
  - codex
created_date: '2026-04-04 04:40'
updated_date: '2026-04-04 05:01'
labels: []
dependencies: []
documentation:
  - docs/workflow/verification.md
  - docs/architecture/README.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix the current src/test failures across stats CLI lifetime rebuild handling, immersion tracker lifetime rebuild idempotency, renderer test environment cleanup helpers, subtitle sidebar auto-follow behavior, and log retention pruning so the maintained test lanes pass again.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Stats CLI lifetime rebuild behavior passes the current regression coverage.
- [x] #2 Immersion tracker lifetime rebuild backfill remains idempotent under the existing runtime test.
- [x] #3 Renderer modal test helpers restore injected globals exactly to prior state.
- [x] #4 Log pruning removes files older than the configured retention window deterministically.
- [x] #5 Relevant targeted test files pass after the fixes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the failing specs in isolation to separate deterministic regressions from suite-order pollution.
2. Fix source or test-helper logic for the three isolated failures: log retention cutoff, stats CLI lifetime rebuild timestamp handling, and subtitle sidebar initial jump behavior.
3. Harden renderer modal cleanup regressions so tests verify descriptor restoration without assuming global window/document start absent.
4. Re-run the targeted failing files, then the required verification gate for the touched areas and record results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Targeted regressions fixed in log pruning, stats CLI lifetime logging/tests, subtitle sidebar auto-follow timing, and renderer global cleanup test isolation.

Verification: `bun test src/main/runtime/stats-cli-command.test.ts src/shared/log-files.test.ts src/renderer/modals/playlist-browser.test.ts src/renderer/modals/youtube-track-picker.test.ts src/renderer/modals/subtitle-sidebar.test.ts` passed.

Verification: `bun run test:src` still exits non-zero because of unrelated existing errors in `src/core/services/anilist/anilist-token-store.test.ts` (`Bun.serve is not a function`) plus one remaining non-task failure elsewhere; the originally reported regressions are green in the maintained lane.

User reported `test:full` still failing after the first regression pass. Reopened to clear the remaining `test:src` fail plus the existing unhandled test errors before handoff.

Verified final gate with `bun run test:launcher:unit:src` and `bun run test:full`; both pass after fixing the launcher AniSkip fallback title regression and the earlier src-lane regressions.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stabilized the failing test regressions across source and launcher lanes. Fixed log pruning cutoff math under Bun BigInt mtimes, subtitle sidebar auto-follow timing, renderer global cleanup test isolation, stats CLI lifetime rebuild logging/tests, stats-server node:http fallback isolation, and launcher AniSkip fallback title resolution so basename titles beat generic parent directories while episode-only filenames still prefer the series directory. Verification passed with `bun test launcher/aniskip-metadata.test.ts`, `bun run test:launcher:unit:src`, and `bun run test:full`.
<!-- SECTION:FINAL_SUMMARY:END -->
