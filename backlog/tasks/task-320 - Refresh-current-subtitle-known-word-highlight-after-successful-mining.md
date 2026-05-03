---
id: TASK-320
title: Refresh current subtitle known-word highlight after successful mining
status: Done
assignee:
  - Codex
created_date: '2026-05-03 03:22'
updated_date: '2026-05-03 03:29'
labels:
  - bug
  - anki
  - subtitle-annotations
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After a sentence card is mined successfully, the mined word is added to the known-word cache and future subtitle appearances render as known. The currently displayed subtitle must also be refreshed immediately so the mined word turns known-color without waiting for a later cue.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Successful sentence-card mining refreshes the current displayed subtitle so newly mined known words render immediately.
- [x] #2 Unsuccessful/no-op mining does not refresh the current subtitle.
- [x] #3 Regression coverage verifies the successful and unsuccessful mining paths.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test around AnkiIntegration known-word cache appends: when mined note info changes known words, a callback fires.
2. Make KnownWordCacheManager.appendFromNoteInfo report whether it changed the immediate known-word cache.
3. Add an AnkiIntegration known-word-cache-updated callback and invoke it after successful immediate append.
4. Wire main process callback to subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText), forcing active-line retokenization after popup/proxy or local mining updates the known-word cache.
5. Add a changelog fragment and run targeted tests plus typecheck.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented generic known-word-cache update notification instead of shortcut-only refresh. KnownWordCacheManager.appendFromNoteInfo now returns whether in-memory known words changed; AnkiIntegration notifies a callback after successful append. Main process wires that callback to subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText), forcing retokenization without using stale prefetch/cache data. Added regression coverage in anki-integration.test.ts.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added a known-word-cache update callback on AnkiIntegration and wired it in the main process to refresh the current subtitle after mined note info changes known words.
- Made KnownWordCacheManager.appendFromNoteInfo report whether it changed the known-word cache, so refresh only happens after an actual immediate known-word append.
- Added regression coverage proving mined note info updates known words and emits the update notification.

Verification:
- bun test src/anki-integration.test.ts src/anki-integration/known-word-cache.test.ts src/main/runtime/anki-actions.test.ts src/main/runtime/anki-actions-main-deps.test.ts
- bun run typecheck
- bun run changelog:lint currently blocked by pre-existing invalid metadata in changes/319-interjection-annotation-filter.md.
<!-- SECTION:FINAL_SUMMARY:END -->
