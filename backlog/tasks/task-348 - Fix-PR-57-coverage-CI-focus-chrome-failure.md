---
id: TASK-348
title: Fix PR 57 coverage CI focus chrome failure
status: Done
assignee:
  - '@codex'
created_date: '2026-05-12 07:02'
updated_date: '2026-05-12 07:11'
labels:
  - ci
  - bug
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
  - 'https://github.com/ksyasuda/SubMiner/actions/runs/25718536412'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix current GitHub Actions `build-test-audit` failure on PR #57 (`tokenizer-updates`). CI fails during `bun run test:coverage:src` in the maintained source lane: `renderer stylesheet hides focus chrome on top-level overlay focus targets`.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Root cause of the focus chrome coverage failure is identified from CI/local test output.
- [x] #2 A focused fix is applied without broad unrelated changes.
- [x] #3 Relevant local coverage/test command passes.
- [x] #4 Remote PR check status is rechecked or next CI action is documented.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the CI failure locally with `bun test src/renderer/overlay-legacy-cleanup.test.ts`.
2. Update the stale legacy cleanup assertion to expect top-level `:focus-visible` suppression and reject broad `:focus` suppression.
3. Run the targeted test and `bun run test:coverage:src` to match CI's failing lane.
4. Recheck PR checks or document that CI needs a push/rerun.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
CI/local root cause: `src/renderer/style.css` was intentionally changed to `html/body/#overlay:focus-visible`, but `src/renderer/overlay-legacy-cleanup.test.ts` still required broad `:focus` selectors. The stale assertion fails in `test:coverage:src`.

Additional coverage-lane failure after first fix: `src/main/runtime/linux-mpv-fullscreen-overlay-refresh.test.ts` imported `updateLinuxMpvFullscreenOverlayRefreshBurst`, but `src/main/runtime/linux-mpv-fullscreen-overlay-refresh.ts` did not export/implement it. Added the helper to cancel existing bursts and schedule only while fullscreen is true.

Verification passed: `bun test src/renderer/overlay-legacy-cleanup.test.ts`; `bun test src/main/runtime/linux-mpv-fullscreen-overlay-refresh.test.ts`; `bun run test:coverage:src`; `bun run format:check:src`. `gh pr checks 57` still reports the old failed `build-test-audit` run at run 25718536412; branch needs push/rerun for remote green.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed current PR #57 `build-test-audit` CI blockers. Updated the stale overlay legacy cleanup assertion to expect `:focus-visible` top-level focus suppression and guard against reintroducing broad `:focus` suppression. Added the missing `updateLinuxMpvFullscreenOverlayRefreshBurst` export used by the Linux fullscreen overlay refresh tests. Verification passed locally: focused overlay legacy cleanup test, focused Linux fullscreen refresh test, `bun run test:coverage:src`, and `bun run format:check:src`. Remote PR checks still show the old failed `build-test-audit` run until these local changes are pushed and CI reruns.
<!-- SECTION:FINAL_SUMMARY:END -->
