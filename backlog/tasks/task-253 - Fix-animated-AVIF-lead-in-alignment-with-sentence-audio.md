---
id: TASK-253
title: Fix animated AVIF lead-in alignment with sentence audio
status: Done
assignee:
  - codex
created_date: '2026-03-30 01:59'
updated_date: '2026-03-30 02:03'
labels: []
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/anki-integration/animated-image-sync.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/anki-integration.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/stats-server.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/media-generator.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Animated AVIF cards currently freeze only for the existing word-audio duration. Because generated sentence audio starts with configured audio padding before the spoken subtitle begins, animation motion can begin early instead of lining up with the spoken sentence. Update the shared lead-in calculation so animated motion begins when sentence speech begins after the chosen word audio finishes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Animated AVIF lead-in calculation includes both the chosen word-audio duration and the generated sentence-audio start offset so motion begins with spoken sentence audio
- [x] #2 Shared animated-image sync behavior is applied consistently across the Anki note update, card creation, and stats server media-generation paths
- [x] #3 Regression tests cover the corrected lead-in timing calculation and fail before the fix
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Approved plan:
1. Add a failing unit test proving animated-image lead-in must include sentence-audio start offset in addition to chosen word-audio duration.
2. Update shared animated-image lead-in resolution to add the configured sentence-audio offset used by generated sentence audio.
3. Thread the shared calculation through note update, card creation, and stats-server generation paths without duplicating timing logic.
4. Run targeted tests first, then the relevant fast verification lane for touched files.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User approved implementation on 2026-03-29 local time. Root cause: lead-in omitted sentence-audio padding offset, so AVIF motion began before spoken sentence audio.

Implemented shared animated-image lead-in fix in src/anki-integration/animated-image-sync.ts by adding the same sentence-audio start offset used by generated audio (`audioPadding`) after summing the chosen word-audio durations.

Added regression coverage in src/anki-integration/animated-image-sync.test.ts for explicit `audioPadding` lead-in alignment and kept the zero-padding case covered.

Verification passed: `bun test src/anki-integration/animated-image-sync.test.ts src/anki-integration/note-update-workflow.test.ts src/media-generator.test.ts`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed animated AVIF lead-in alignment so motion starts when the spoken sentence starts, not at the padded beginning of the generated sentence-audio clip. The shared resolver in `src/anki-integration/animated-image-sync.ts` now adds the configured/default `audioPadding` offset after summing the selected word-audio durations, which keeps note update, card creation, and stats-server generation paths aligned through the same logic.

Added regression coverage in `src/anki-integration/animated-image-sync.test.ts` for both zero-padding and explicit padding cases to prove the lead-in math matches sentence-audio timing.

Verification:
- `bun test src/anki-integration/animated-image-sync.test.ts src/anki-integration/note-update-workflow.test.ts src/media-generator.test.ts`
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run test:smoke:dist`
<!-- SECTION:FINAL_SUMMARY:END -->
