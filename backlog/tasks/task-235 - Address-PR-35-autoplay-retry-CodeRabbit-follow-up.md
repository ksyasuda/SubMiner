---
id: TASK-235
title: 'Address PR #35 autoplay retry CodeRabbit follow-up'
status: Done
assignee:
  - codex
created_date: '2026-03-26 04:30'
updated_date: '2026-03-31 19:37'
labels:
  - review-comments
  - coderabbit
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
ordinal: 176500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess and implement the latest CodeRabbit follow-up on PR #35 concerning stale autoplay-ready fallback retries interfering with a new app-owned YouTube playback flow in main.ts.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Starting a new app-owned YouTube playback flow invalidates any pending autoplay-ready fallback retries from older playback state before mpv prep begins.
- [x] #2 Relevant verification for the touched main.ts autoplay retry logic passes locally.
- [x] #3 Task notes/final summary capture the fix and verification.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a helper in src/main.ts that invalidates pending autoplay-ready fallback retry state by clearing the tracked media path and advancing the autoplay generation counter.
2. Invoke that helper at the start of runYoutubePlaybackFlowMain before app-owned YouTube playback takes over so stale retries cannot unpause reused playback.
3. Run relevant verification for the touched main.ts path and record results in the task notes/final summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added invalidatePendingAutoplayReadyFallbacks() in src/main.ts to clear the tracked autoplay-ready media path and advance the autoplay generation before a new app-owned YouTube flow claims playback. This invalidates stale fallback retry closures even when the reused playback path is the same.

Verification: bun test src/main/runtime/mpv-main-event-actions.test.ts; bun test src/main/runtime/startup-autoplay-release-policy.test.ts; bun run typecheck.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed the latest PR #35 CodeRabbit follow-up by invalidating pending autoplay-ready fallback retries before a new app-owned YouTube playback flow takes over in src/main.ts. The new helper clears the tracked autoplay media path and advances the autoplay generation counter, so retry closures from older playback state cannot later unpause the newly prepared flow when reusing the same media path.

Verification:
- bun test src/main/runtime/mpv-main-event-actions.test.ts
- bun test src/main/runtime/startup-autoplay-release-policy.test.ts
- bun run typecheck
<!-- SECTION:FINAL_SUMMARY:END -->
