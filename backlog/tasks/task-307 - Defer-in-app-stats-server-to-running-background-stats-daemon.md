---
id: TASK-307
title: Defer in-app stats server to running background stats daemon
status: Done
assignee: []
created_date: '2026-04-27 01:57'
updated_date: '2026-04-27 02:02'
labels:
  - stats
  - runtime
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When normal SubMiner app startup has stats auto-start enabled, it should detect an already-running background stats daemon and avoid starting a second in-app stats server. Stats overlay/dashboard URL resolution should point at the background daemon.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 If a live background stats daemon state exists for another process, in-app stats auto-start does not start a local stats server.
- [x] #2 Stats URL resolution returns the background daemon URL when the background daemon is live.
- [x] #3 Stale or dead background daemon state is cleared and normal in-app stats startup still works.
- [x] #4 Regression tests cover the deferral behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add unit tests for stats server routing decisions around live/stale background daemon state.
2. Implement a small routing helper used by main stats startup.
3. Wire `ensureStatsServerStarted()` through the helper.
4. Run targeted tests/typecheck/changelog lint and finalize the task.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Extracted stats server URL routing into `src/main/runtime/stats-server-routing.ts` and wired `main.ts` through it. The helper returns the background daemon URL without calling local server startup when a live external daemon exists; dead/self-owned stale state is removed before falling back to local startup. Added the new test to `test:core:src`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added a pure stats server routing helper that chooses between a live background daemon and local in-app stats server startup.
- Updated main stats URL resolution to defer to another process's background daemon and only start the in-app server when no live daemon is available.
- Added regression tests for live daemon deferral, dead daemon cleanup, self-owned stale state cleanup, and local server reuse.
- Added the routing test to the core source test lane and added a changelog fragment.

Verification:
- `bun test src/main/runtime/stats-server-routing.test.ts src/main-entry-runtime.test.ts src/main/runtime/stats-cli-command.test.ts src/stats-daemon-control.test.ts`
- `bun run test:core:src`
- `bun run typecheck`
- `bun run changelog:lint`
<!-- SECTION:FINAL_SUMMARY:END -->
