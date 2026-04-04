---
id: TASK-276
title: Restore canonical SubMiner Linux app-id metadata
status: Done
assignee:
  - codex
created_date: '2026-04-04 06:31'
updated_date: '2026-04-04 06:34'
labels:
  - bug
  - linux
  - electron
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the Linux/Wayland packaged app metadata so the OS-facing desktop/app-id metadata stays canonical `SubMiner` instead of the lowercase npm package name `subminer`. Add regression coverage around packaged metadata so future Electron/runtime changes do not silently reintroduce the lowercase class/app-id behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Top-level package metadata provides the canonical capitalized app name used by Electron runtime bootstrap on Linux
- [x] #2 Packaged Linux app metadata resolves to `SubMiner`/`SubMiner.desktop` instead of lowercase `subminer`
- [x] #3 Regression coverage fails before the fix and passes after it
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused packaging/runtime regression test that reads the packaged app metadata source and asserts the Linux Electron bootstrap-visible fields resolve to canonical `SubMiner` / `SubMiner.desktop`.
2. Run the targeted test first to capture the failing pre-fix state.
3. Update top-level package metadata in `package.json` with the canonical Electron runtime-facing fields needed for Linux bootstrap.
4. Re-run the targeted test and a lightweight packaging validation to confirm the packaged metadata now stays canonical.
5. Record verification notes and complete the task if all acceptance criteria pass.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a regression in `src/release-workflow.test.ts` asserting top-level `productName` and `desktopName` stay canonical for Linux Electron runtime bootstrap. Verified the new test failed before the fix because both fields were missing from top-level package metadata.

Updated top-level `package.json` metadata with `productName: SubMiner` and `desktopName: SubMiner.desktop` so packaged `app.asar` exposes the canonical Linux startup identity Electron reads before app code runs.

Verification passed with `bun test src/release-workflow.test.ts`, `bun run build && ./node_modules/.bin/electron-builder --linux dir --publish never`, packaged `release/linux-unpacked/resources/app.asar` inspection showing `{ name: subminer, productName: SubMiner, desktopName: SubMiner.desktop }`, and `bun run changelog:lint`.

Ran the full default handoff gate after the targeted/package verification: `bun run typecheck`, `bun run test:fast`, `bun run test:env`, and `bun run test:smoke:dist` all passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored canonical Linux/Wayland app identity metadata by adding top-level Electron runtime fields to `package.json`: `productName: SubMiner` and `desktopName: SubMiner.desktop`. This fixes the packaged app metadata Electron reads before user code runs, so native Wayland compositors no longer need to derive the app-id/class from the lowercase npm package name alone.

Added a regression test in `src/release-workflow.test.ts` that asserts the runtime-visible top-level metadata stays canonical. The new test was run first and failed before the fix because `productName` was missing, then passed after the metadata update.

Verification: `bun test src/release-workflow.test.ts`; `bun run build && ./node_modules/.bin/electron-builder --linux dir --publish never`; inspected `release/linux-unpacked/resources/app.asar` and confirmed `productName: SubMiner` plus `desktopName: SubMiner.desktop`; `bun run changelog:lint`. Added changelog fragment `changes/2026-04-04-linux-app-id-metadata.md`.

Full default handoff gate also passed: `bun run typecheck`; `bun run test:fast`; `bun run test:env`; `bun run test:smoke:dist`.
<!-- SECTION:FINAL_SUMMARY:END -->
