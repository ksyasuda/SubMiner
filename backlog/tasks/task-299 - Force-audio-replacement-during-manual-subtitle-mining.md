---
id: TASK-299
title: Force audio replacement during manual subtitle mining
status: Done
assignee:
  - Codex
created_date: '2026-04-26 00:10'
updated_date: '2026-04-26 02:42'
labels:
  - anki
  - mining
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Manual subtitle mining via the Ctrl+C/Ctrl+V flow should replace expression and sentence audio fields even when the user has configured media overwrite fields to false. These fields can already contain proxy-inserted SubMiner audio on a new card, and manual update intent is to replace that generated content.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Manual subtitle mining replaces existing expression audio content regardless of configured audio overwrite settings.
- [x] #2 Manual subtitle mining replaces existing sentence audio content regardless of configured audio overwrite settings.
- [x] #3 Non-manual mining/update flows continue to respect configured audio overwrite settings.
- [x] #4 A regression test covers manual audio replacement when overwrite settings are disabled.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Locate the manual subtitle mining Ctrl+C/Ctrl+V flow and the Anki media field overwrite gate.
2. Add a failing regression test showing manual mining overwrites expression and sentence audio when configured audio overwrite is disabled.
3. Implement the smallest path-specific override so only manual subtitle mining forces audio replacement.
4. Run the focused mining test and update task acceptance criteria/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented focused manual clipboard update behavior in CardCreationService.updateLastAddedFromClipboard: generated manual audio is written to both resolved sentence audio and expression audio fields with forced overwrite. Other update flows still use existing overwrite config paths.

Verification: focused Anki tests passed; typecheck passed; changelog lint and diff check passed. Full bun run test:fast was attempted but is blocked by unrelated existing tokenizer annotation-stage failures tied to dirty task 298 worktree changes.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Manual clipboard subtitle updates now resolve both sentence audio and expression audio fields and replace both with the newly generated audio regardless of ankiConnect.behavior.overwriteAudio.
- Added a regression test for the Ctrl+C/Ctrl+V manual update path with existing proxy-inserted audio and overwriteAudio disabled.
- Registered the regression test in test:fast, documented the overwrite exception in user docs, and added a changelog fragment.

Verification:
- bun test src/anki-integration/card-creation-manual-update.test.ts src/anki-integration/card-creation.test.ts
- bun run tsc --noEmit
- bun test src/main-entry-runtime.test.ts src/anki-integration.test.ts src/anki-integration/card-creation-manual-update.test.ts src/anki-integration/anki-connect-proxy.test.ts src/anki-integration/field-grouping-workflow.test.ts src/anki-integration/field-grouping.test.ts src/anki-integration/field-grouping-merge.test.ts src/release-workflow.test.ts src/prerelease-workflow.test.ts src/ci-workflow.test.ts scripts/build-changelog.test.ts scripts/electron-builder-after-pack.test.ts scripts/mkv-to-readme-video.test.ts scripts/run-coverage-lane.test.ts scripts/update-aur-package.test.ts
- bun run changelog:lint
- bun run docs:test
- bun run docs:build
- git diff --check

Blocked gate:
- bun run test:fast currently fails in unrelated src/core/services/tokenizer/annotation-stage.test.ts tests for kana-only non-independent noun helper merges; those files have pre-existing dirty changes outside this task.
<!-- SECTION:FINAL_SUMMARY:END -->
