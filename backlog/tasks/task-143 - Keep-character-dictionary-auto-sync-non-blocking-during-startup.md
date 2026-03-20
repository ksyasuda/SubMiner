---
id: TASK-143
title: Keep character dictionary auto-sync non-blocking during startup
status: In Progress
assignee:
  - codex
created_date: '2026-03-09 01:45'
updated_date: '2026-03-20 09:22'
labels:
  - dictionary
  - startup
  - performance
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/current-media-tokenization-gate.ts
priority: high
ordinal: 38500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Keep character dictionary auto-sync running in parallel during startup without delaying playback. Only tokenization readiness should gate playback; character dictionary import/settings updates should wait until tokenization is already ready and then refresh annotations afterward.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Character dictionary snapshot/build work can run immediately during startup.
- [x] #2 Yomitan dictionary mutation work waits until current-media tokenization is ready.
- [x] #3 Regression coverage verifies auto-sync builds before the gate and only mutates Yomitan after the gate resolves.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test for startup autoplay release surviving delayed mpv readiness or late subtitle refresh after dictionary sync.
2. Harden the autoplay-ready release path so paused startup keeps retrying until mpv is actually released or media changes, without resuming user-paused playback later.
3. Keep the existing character-dictionary revisit fixes and paused-startup OSD fixes aligned with the autoplay change, then run targeted runtime tests and typecheck.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a small current-media tokenization gate in main runtime. Media changes reset the gate, the first tokenization-ready event marks it ready, and auto-sync now waits on that gate only before Yomitan dictionary inspection/import/settings updates. Snapshot generation and merged ZIP build still run immediately in parallel.

2026-03-20: User reports startup remains paused after annotations/tokenization are visible and only resumes after character-dictionary generation/import finishes. Investigating autoplay-ready release regression vs dictionary sync completion refresh.

2026-03-20: Added startup autoplay retry-budget helper so paused startup retries cover the full plugin gate window instead of only ~2.8s. Verification: bun test src/main/runtime/startup-autoplay-release-policy.test.ts src/main/runtime/character-dictionary-auto-sync.test.ts src/main/runtime/startup-osd-sequencer.test.ts src/main/runtime/character-dictionary-auto-sync-completion.test.ts; bun run typecheck; bun run test:fast; bun run test:env; bun run build; bun run test:smoke:dist; runtime-compat verifier passed at .tmp/skill-verification/subminer-verify-20260320-022106-nM28Nk. Pending real installed-app/mpv validation.
<!-- SECTION:NOTES:END -->
