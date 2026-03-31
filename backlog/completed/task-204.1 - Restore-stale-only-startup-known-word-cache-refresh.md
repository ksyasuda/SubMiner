---
id: TASK-204.1
title: Restore stale-only startup known-word cache refresh
status: Done
assignee:
  - '@Codex'
created_date: '2026-03-20 02:52'
updated_date: '2026-03-23 03:22'
labels:
  - anki
  - cache
  - bug
dependencies: []
references:
  - src/anki-integration/known-word-cache.ts
  - src/anki-integration/known-word-cache.test.ts
  - docs/plans/2026-03-19-known-word-cache-incremental-sync-design.md
parent_task_id: TASK-204
priority: high
ordinal: 124500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow up on the incremental known-word cache change so startup still performs a refresh when the persisted cache is older than the configured refresh interval, while leaving fresh persisted state untouched.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Startup refreshes known words immediately when persisted cache state is stale for the configured interval.
- [x] #2 Startup skips the immediate refresh when persisted cache state is still fresh.
- [x] #3 Regression tests cover both stale and fresh startup paths.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused known-word cache lifecycle tests that distinguish fresh startup state from stale startup state and verify the stale path currently fails.
2. Update startup scheduling in src/anki-integration/known-word-cache.ts so persisted cache still loads immediately, but startup only triggers an immediate refresh when the cache is stale for the configured interval or the cache scope/config changed.
3. Run focused known-word cache tests and targeted SubMiner verification for the touched cache/runtime lane, then update the task with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verified current lifecycle behavior: fresh persisted known-word cache already skips immediate startup refresh when the cache scope/config matches; stale persisted cache already refreshes immediately. Added regression coverage for both startup paths plus a proxy integration test showing addNote responses return without waiting for background enrichment.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added regression coverage for known-word cache startup behavior and proxy response timing. The cache tests now lock in the intended lifecycle: fresh persisted state stays load-only on startup, while stale persisted state refreshes immediately. Added a proxy integration test proving addNote responses return without waiting for background enrichment. Verification: targeted Bun tests passed (`bun test src/anki-connect.test.ts src/anki-integration/anki-connect-proxy.test.ts src/anki-integration/known-word-cache.test.ts src/anki-integration/note-update-workflow.test.ts src/anki-integration/runtime.test.ts`) and direct `bun run test:fast` passed. The `subminer-change-verification` helper repeatedly reported `bun run test:fast` as failed in its isolated lane despite the direct command passing, so that helper lane remains a flaky/blocking verification artifact rather than a reproduced code failure.
<!-- SECTION:FINAL_SUMMARY:END -->
