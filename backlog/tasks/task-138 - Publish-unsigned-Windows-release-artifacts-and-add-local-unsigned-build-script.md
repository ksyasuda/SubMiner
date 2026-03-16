---
id: TASK-138
title: Publish unsigned Windows release artifacts and add local unsigned build script
status: Done
assignee:
  - codex
created_date: '2026-03-09 00:00'
updated_date: '2026-03-16 05:13'
labels:
  - release
  - windows
dependencies: []
references:
  - .github/workflows/release.yml
  - package.json
  - src/release-workflow.test.ts
priority: high
ordinal: 44500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Stop the tag-driven release workflow from depending on SignPath and publish unsigned Windows `.exe` and `.zip` artifacts directly. Add an explicit local `build:win:unsigned` script without changing the existing `build:win` command.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Windows release CI builds unsigned artifacts without requiring SignPath secrets.
- [x] #2 The Windows release job uploads `release/*.exe` and `release/*.zip` directly as the `windows` artifact.
- [x] #3 The repo exposes a local `build:win:unsigned` script for explicit unsigned Windows packaging.
- [x] #4 Regression coverage fails if the workflow reintroduces SignPath submission or drops the unsigned script.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update workflow regression tests to assert unsigned Windows release behavior and the new local script.
2. Patch `package.json` to add `build:win:unsigned`.
3. Patch `.github/workflows/release.yml` to build unsigned Windows artifacts and upload them directly.
4. Add the release changelog fragment and run focused verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed the Windows SignPath secret validation and submission steps from `.github/workflows/release.yml`. The Windows release job now runs `bun run build:win:unsigned` and uploads `release/*.exe` and `release/*.zip` directly as the `windows` artifact consumed by the release job.

Added `scripts/build-win-unsigned.mjs` plus the `build:win:unsigned` package script. The wrapper clears Windows code-signing environment variables and disables identity auto-discovery before invoking `electron-builder`, so release CI stays unsigned even if signing credentials are configured elsewhere.

Updated `src/release-workflow.test.ts` to assert the unsigned workflow contract and added the release changelog fragment in `changes/unsigned-windows-release-builds.md`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Windows release CI now publishes unsigned artifacts directly and no longer depends on SignPath. Local developers also have an explicit `bun run build:win:unsigned` path for unsigned packaging without changing the existing `build:win` command.

Verification:

- `bun test src/release-workflow.test.ts`
- `bun run typecheck`
- `node --check scripts/build-win-unsigned.mjs`
<!-- SECTION:FINAL_SUMMARY:END -->
