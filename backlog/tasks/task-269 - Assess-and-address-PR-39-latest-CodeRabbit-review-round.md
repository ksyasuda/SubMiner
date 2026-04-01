---
id: TASK-269
title: 'Assess and address PR #39 latest CodeRabbit review round'
status: Done
assignee:
  - codex
created_date: '2026-04-01 07:22'
updated_date: '2026-04-01 07:55'
labels: []
dependencies: []
references:
  - src/main
  - docs/architecture/README.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess unresolved CodeRabbit review threads on PR #39, fix valid findings in the current branch, and document any non-actionable findings so the review state is clear for follow-up.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All unresolved CodeRabbit review threads on PR #39 are assessed and classified as fix or non-actionable with rationale
- [x] #2 Valid findings are addressed in code without regressing current runtime behavior
- [x] #3 Regression tests or targeted coverage are added when the reviewed behavior can be exercised locally
- [x] #4 Relevant verification commands are run and results recorded in the task summary
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Confirm each unresolved CodeRabbit thread against current branch state and mark already-addressed items as non-actionable without code churn.
2. Fix validated runtime issues in focused patches: AniList setup window stale close handler, headless known-word refresh effective config usage, main startup websocket probe wiring, startup warmup delegation, mpv media-path duplicate side effects, overlay visibility readiness guard, Discord presence service cleanup, and stats daemon self-stop handling.
3. Tighten headless startup runtime typing so custom bootstrap deps require an explicit factory instead of unsafe casting.
4. Add or extend targeted tests for behaviors that can be exercised locally, especially overlay visibility, Discord lifecycle cleanup, stats self-stop, and headless startup typing/runtime coverage.
5. Run targeted tests first, then the relevant broader verification lane, and capture which CodeRabbit items were fixed versus assessed as already addressed.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reassessed current branch after reset concern: most runtime/test fixes were still present; only live worktree delta was IPC playlist-browser cleanup plus the untracked backlog task record.

Validated and fixed CodeRabbit findings for stale AniList setup window clearing, effective headless Anki config usage, CLI websocket probe wiring, duplicate mpv media-path side effects, overlay visibility readiness, Discord presence service cleanup, stats self-stop, unsafe headless startup custom deps typing, and playlist-browser IPC wiring.

Assessed main-startup-runtime-bootstrap warmup comment as non-actionable at this layer: the guarded startBackgroundWarmupsIfAllowed wrapper exists in main-startup-bootstrap, while this lower bootstrap still needs to expose the raw mpv warmup command.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the unresolved CodeRabbit round on PR #39 and applied the confirmed fixes across the main-process runtime split. Updated AniList setup window teardown to avoid clearing newer windows, made headless known-word refresh use one effective Anki config end-to-end, wired the real mpv websocket probe into CLI startup, removed duplicate mpv media-path side effects, guarded overlay visibility changes behind overlay readiness, stopped Discord presence services before disable/replace, and prevented stats daemon stop from SIGTERMing the current Electron process. Also tightened headless startup typing so custom bootstrap deps require an explicit factory, normalized stats coordinator note-id typing, and completed playlist-browser IPC wiring/types so typecheck stays green.

Added targeted regression coverage for overlay visibility initialization, Discord lifecycle cleanup, stats self-stop, and headless startup custom-deps typing expectations. Verification run from the current tree: `bun run typecheck`, targeted `bun test` for overlay/discord/stats/headless startup files, and full `bun run test:fast` all passed.

Assessment outcome for remaining review noise: the playlist-browser open action was already wired in IPC bootstrap, and the warmup-delegation comment on `main-startup-runtime-bootstrap.ts` was not applied because the guarded wrapper lives one layer higher in `main-startup-bootstrap.ts`; this lower bootstrap still needs to surface the raw mpv warmup command consumed by that wrapper.
<!-- SECTION:FINAL_SUMMARY:END -->
