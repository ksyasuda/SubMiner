---
id: TASK-322
title: Fix failing CI checks on PR 57
status: Done
assignee:
  - codex
created_date: '2026-05-03 06:27'
updated_date: '2026-05-03 06:31'
labels:
  - ci
  - bug
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix failing GitHub Actions checks on PR #57 (`tokenizer-updates`). Scope: use CI logs to identify root cause, apply focused local fix, and verify with relevant local checks.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Failing GitHub Actions check root cause is identified from logs.
- [x] #2 A focused code/test/docs fix is applied locally.
- [x] #3 Relevant local verification passes or blocked reason is documented.
- [x] #4 PR checks are rechecked or next CI action is documented.
- [x] #5 Actionable CodeRabbit PR comments are inspected and addressed or documented as non-actionable.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Fix CI changelog lint by adding a valid `type` frontmatter value to `changes/319-interjection-annotation-filter.md`.
2. Address unresolved CodeRabbit threads:
   - `scripts/test-plugin-session-bindings.lua`: make `.tmp` creation portable across Unix/Windows shells.
   - `src/core/services/tokenizer.ts`: pass `TokenizerAnnotationOptions` through `stripSubtitleAnnotationMetadata` paths so `sourceText` is honored.
   - `src/main/runtime/mpv-main-event-main-deps.ts`: align overlay-runtime quit-on-disconnect predicate with `hasInitialPlaybackQuitOnDisconnectArg`.
   - `src/renderer/handlers/mouse.test.ts`: make `elementFromPoint` stubs coordinate-sensitive.
3. Run focused checks: `bun run changelog:lint`, relevant tokenizer/main/mouse tests, and plugin Lua test path if available.
4. Recheck PR checks/comments after local verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
CI root cause: GitHub Actions `build-test-audit` failed during `bun run changelog:lint`; `changes/319-interjection-annotation-filter.md` must declare `type` as one of `added`, `changed`, `fixed`, `docs`, `internal`. Scope expanded by user to also address CodeRabbit comments on PR #57.

Implemented CI changelog metadata fix and unresolved CodeRabbit feedback locally. Full verification run: `bun run changelog:lint`, focused tests, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`, `bun run format:check:src`. Rechecked PR checks: remote `build-test-audit` still shows the old failing run until this branch is pushed; CodeRabbit remains pending remotely until review reruns.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed PR #57 CI failure by converting `changes/319-interjection-annotation-filter.md` to valid changelog fragment metadata. Addressed unresolved CodeRabbit feedback by making plugin test `.tmp` creation portable, threading tokenizer annotation options through metadata stripping, aligning quit-on-disconnect predicates for Jellyfin playback, and strengthening mouse hit-test assertions. Also formatted two existing PR files required by the source format gate. Verification passed locally: changelog lint, focused tests, typecheck, test:fast, test:env, build, smoke dist, and format check. Remote PR checks still show the previous failed `build-test-audit` run until these local changes are pushed.
<!-- SECTION:FINAL_SUMMARY:END -->
