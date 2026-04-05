---
id: TASK-279
title: Fix Linux AppImage child-process libffmpeg resolution
status: In Progress
assignee:
  - '@codex'
created_date: '2026-04-05 17:17'
updated_date: '2026-04-05 17:55'
labels: []
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/issues/41'
documentation:
  - docs/workflow/verification.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the Linux AppImage packaging so Chromium child processes relaunched from the bundled binary can resolve the packaged libffmpeg shared library and SubMiner starts cleanly instead of crash-looping on network-service restarts.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Linux AppImage packaging ensures bundled Chromium child processes can resolve the packaged libffmpeg shared library during relaunch.
- [x] #2 Regression coverage exercises the Linux packaging/build configuration that provides the AppImage shared-library path.
- [x] #3 Release notes/changelog reflect the Linux AppImage startup fix.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression tests for Linux release packaging that assert the build config invokes an `afterPack` hook and that the hook stages bundled `libffmpeg.so` into `usr/lib` for AppImage runtime lookup.
2. Implement a small electron-builder `afterPack` hook that runs only for Linux, copies `libffmpeg.so` from the packaged app root into `usr/lib`, and no-ops when the source library is absent.
3. Wire the hook into `package.json` build config and add a changelog fragment for the Linux AppImage startup fix.
4. Run the focused test lane first, then the default handoff gate because the change touches release-sensitive packaging behavior.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Chose a repo-local electron-builder `afterPack` hook instead of patching/forking `electron-builder`. The hook copies bundled `libffmpeg.so` from the packaged Linux app root into `usr/lib`, matching the AppImage runtime's existing `LD_LIBRARY_PATH` search path.

Added regression coverage for both config wiring (`src/release-workflow.test.ts`) and the hook behavior (`scripts/electron-builder-after-pack.test.ts`), then wired the new script test into `test:fast` so the maintained lane keeps exercising the fix.

Verification passed: `bun test scripts/electron-builder-after-pack.test.ts src/release-workflow.test.ts`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`.

Addressed PR #45 CodeRabbit review thread: Linux `afterPack` staging now hard-fails when `libffmpeg.so` is missing instead of silently no-oping. Updated focused hook tests to assert the new failure contract and that `afterPack` propagates Linux staging errors.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a shared electron-builder `afterPack` hook at `scripts/electron-builder-after-pack.cjs` and wired it into `package.json` so Linux packaging stages the bundled `libffmpeg.so` into `usr/lib` inside the packaged app. This keeps Chromium child relaunches compatible with the AppImage runtime's existing `LD_LIBRARY_PATH` layout without forking or patching upstream `electron-builder`.

Regression coverage now checks both the packaging config and the hook behavior: `src/release-workflow.test.ts` asserts the hook stays wired into release config, and `scripts/electron-builder-after-pack.test.ts` verifies Linux copies `libffmpeg.so` into `usr/lib` while non-Linux and missing-library cases no-op safely. The new script test is included in `test:fast`, and a changelog fragment was added under `changes/fix-appimage-libffmpeg-path.md`.

Verification passed with `bun test scripts/electron-builder-after-pack.test.ts src/release-workflow.test.ts`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, and `bun run test:smoke:dist`.
<!-- SECTION:FINAL_SUMMARY:END -->
