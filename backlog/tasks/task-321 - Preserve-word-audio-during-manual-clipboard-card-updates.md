---
id: TASK-321
title: Preserve word audio during manual clipboard card updates
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-03 06:22'
updated_date: '2026-05-03 06:23'
labels:
  - anki
  - mining
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Manual Ctrl+Shift+C/Ctrl+V card updates on already-mined cards should refresh the sentence content and generated sentence media without removing or replacing the existing word/expression audio. The word is unchanged in this flow, so the configured word audio field must be left untouched while sentence audio remains forced-overwrite behavior from TASK-299.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Manual clipboard subtitle update replaces the resolved sentence audio field with newly generated sentence audio.
- [x] #2 Manual clipboard subtitle update does not include the configured word/expression audio field in Anki field updates.
- [x] #3 Animated image generation still uses the existing word audio duration for lead-in sync when configured.
- [x] #4 A regression test covers preserving word/expression audio during manual clipboard update.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update the focused manual clipboard card update regression so generated audio is written only to the resolved sentence audio field and the configured word/expression audio field is absent from updateNoteFields payloads.
2. Run the focused test and confirm it fails for the existing TASK-299 behavior.
3. Change CardCreationService.updateLastAddedFromClipboard to stop merging/updating expression audio while preserving forced overwrite for sentence audio.
4. Run the focused test; then run adjacent Anki card-creation tests if the focused gate passes.
5. Update task acceptance criteria/final notes with verification results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented narrow manual clipboard update change in CardCreationService.updateLastAddedFromClipboard: generated audio now force-overwrites only the resolved sentence audio field and no longer writes the configured word/expression audio field. Animated AVIF lead-in still runs from the original note info before image generation, preserving existing word-audio sync behavior.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Manual Ctrl+Shift+C/Ctrl+V card updates now leave the configured word/expression audio field untouched while force-replacing the resolved sentence audio field.
- Updated the regression test to assert the Anki update payload omits ExpressionAudio and only merges SentenceAudio with forced overwrite.
- Updated docs-site behavior notes and added a changelog fragment for the sentence-only manual audio replacement behavior.

Verification:
- bun test src/anki-integration/card-creation-manual-update.test.ts src/anki-integration/card-creation.test.ts src/anki-integration/animated-image-sync.test.ts
- bun run typecheck
- bun run docs:test
- bun run docs:build
- git diff --check -- src/anki-integration/card-creation.ts src/anki-integration/card-creation-manual-update.test.ts docs-site/mining-workflow.md docs-site/anki-integration.md docs-site/configuration.md changes/322-preserve-word-audio-manual-update.md

Blocked gate:
- bun run changelog:lint is blocked by pre-existing malformed changes/319-interjection-annotation-filter.md, which is outside this task's files.
<!-- SECTION:FINAL_SUMMARY:END -->
