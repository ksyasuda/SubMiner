---
id: TASK-195
title: Keep final card-mine OSD result from being overwritten by progress spinner
status: Done
assignee:
  - Codex
created_date: '2026-03-18 19:40'
updated_date: '2026-03-18 19:49'
labels:
  - anki
  - ui
  - bug
milestone: m-1
dependencies: []
references:
  - src/anki-integration/ui-feedback.ts
  - src/anki-integration.ts
  - src/anki-integration/card-creation.ts
priority: medium
ordinal: 105610
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When a card mine finishes, the mpv OSD currently tries to show the final status text but the in-flight Anki progress spinner can immediately overwrite it on the next tick. Stop the spinner first, then show a single-line final result with a success/failure marker and the mined-word notification.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Successful mine/update OSD results render after the spinner is stopped and do not get overwritten by a later spinner tick.
- [x] #2 Failure results that replace the spinner show an `x` marker and stay visible on the same OSD line.
- [x] #3 Regression coverage locks the spinner teardown/result-notification ordering.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a focused failing regression test around the Anki UI-feedback spinner/result helper.
2. Add a helper that stops progress before emitting the final OSD result line with `✓`/`x`.
3. Route mine/update result notifications through that helper, then run targeted verification.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->

Added a dedicated Anki UI-feedback result helper that force-clears the in-flight spinner state before emitting the final OSD result line. Successful card-update notifications now render as `✓ Updated card: ...`, and sentence-card creation failures now render as `x Sentence card failed: ...` without a later spinner tick reclaiming the line.

Verification:

- `bun test src/anki-integration/ui-feedback.test.ts`
- `bun test src/anki-integration/ui-feedback.test.ts src/anki-integration/note-update-workflow.test.ts src/anki-integration.test.ts src/core/services/mining.test.ts src/main/runtime/mining-actions.test.ts`
- `bun x prettier --check src/anki-integration/ui-feedback.ts src/anki-integration/ui-feedback.test.ts src/anki-integration.ts src/anki-integration/card-creation.ts "backlog/tasks/task-195 - Keep-final-card-mine-OSD-result-from-being-overwritten-by-progress-spinner.md" changes/2026-03-18-mine-osd-spinner-result.md`
- `bun run changelog:lint`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/anki-integration/ui-feedback.ts src/anki-integration/ui-feedback.test.ts src/anki-integration.ts src/anki-integration/card-creation.ts changes/2026-03-18-mine-osd-spinner-result.md`
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260318-194614-uZMrAx/`

<!-- SECTION:OUTCOME:END -->
