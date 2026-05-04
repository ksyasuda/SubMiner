---
id: TASK-328
title: Keep subtitle prefetch running after immediate cached annotation render
status: Done
assignee:
  - codex
created_date: '2026-05-04 01:26'
updated_date: '2026-05-04 01:30'
labels: []
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/mpv-main-event-actions.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-processing-controller.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/backlog/completed/task-197 -
    Eliminate-per-line-plain-subtitle-flash-on-prefetch-cache-hit.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/backlog/completed/task-196 -
    Fix-subtitle-prefetch-cache-key-mismatch-and-active-cue-window.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Cached subtitle annotation hits should render annotated subtitles immediately without starving the subtitle prefetcher. Current evidence: the mpv subtitle-change path emits the cached payload before forwarding the subtitle change; in the runtime, the cached emit resumes prefetch, then the forwarded change pauses it, and no async controller emit follows on a cache hit to resume it again.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Cached subtitle annotation payloads still render immediately without a plain subtitle flash.
- [x] #2 A cache-hit subtitle-change event leaves subtitle prefetch eligible to continue after the immediate annotated emit.
- [x] #3 Cache-miss subtitle-change behavior still shows plain text immediately while async annotation processing runs.
- [x] #4 Regression coverage proves the cache-hit ordering that prevents prefetch from staying paused.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression test in `src/main/runtime/mpv-main-event-actions.test.ts` proving cache-hit subtitle changes pause live prefetch work before emitting the immediate annotated payload, so the emit resumes prefetch last.
2. Change `createHandleMpvSubtitleChangeHandler` ordering in `src/main/runtime/mpv-main-event-actions.ts`: set current text, consume cache, forward `onSubtitleChange(text)`, then emit cached payload or plain fallback, then refresh Discord presence.
3. Preserve existing behavior: cache hits emit annotated payload synchronously; cache misses emit `{ text, tokens: null }` synchronously.
4. Run focused tests for `mpv-main-event-actions`; run adjacent controller/prefetch tests if ordering touches cache assumptions.
5. Update TASK-328 acceptance criteria and add a changelog fragment if the repo requires one for this user-visible fix.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Red/green: added cache-hit ordering regression in `src/main/runtime/mpv-main-event-actions.test.ts`; first run failed with actual order `emit:annotated` before `process:line`. Fix narrows ordering change to cache hits only: cache hit calls `onSubtitleChange` before immediate annotated emit; cache miss keeps plain broadcast before processing.

Verification: `bun test src/main/runtime/mpv-main-event-actions.test.ts` passed; `bun test src/core/services/subtitle-processing-controller.test.ts` passed; `bun test src/core/services/subtitle-prefetch.test.ts` passed; combined targeted test command passed 35 tests; `bun run typecheck` passed. `bun run changelog:lint` blocked by unrelated pre-existing `changes/326-anilist-time-position-post-watch.md` missing a valid `type` metadata line.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the subtitle cache-hit ordering that could leave subtitle prefetch paused after an immediate annotated render. Cache hits now forward the subtitle change first, then emit the cached annotated payload, so the runtime pause happens before the emit path resumes prefetch. Cache misses keep the previous plain-subtitle-first path so fallback text still appears immediately while tokenization runs.

Added a regression test for the cache-hit ordering and a changelog fragment for the overlay fix. Verified with targeted subtitle runtime/controller/prefetch tests and `bun run typecheck`; changelog lint is blocked by an unrelated existing malformed fragment for TASK-326.
<!-- SECTION:FINAL_SUMMARY:END -->
