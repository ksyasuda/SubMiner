---
id: TASK-309
title: Accept modified follow-up digits for multi-line sentence mining
status: Done
assignee:
  - '@codex'
created_date: '2026-04-27 20:06'
updated_date: '2026-04-27 20:15'
labels:
  - bug
  - linux
  - shortcuts
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

On Linux, `Ctrl+Shift+S` starts multi-line sentence-card mining, but the follow-up digit is not accepted and the prompt times out. Restore reliable digit capture for the multi-mine flow, including the common case where the original shortcut modifiers are still held briefly while pressing the digit.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 `Ctrl+Shift+S` followed by a number-row digit creates a counted `mineSentenceMultiple` request instead of timing out.
- [x] #2 Follow-up digit capture works when the user has not fully released `Ctrl`/`Shift` after the starter shortcut.
- [x] #3 Regression coverage includes renderer session bindings and mpv plugin numeric selection.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Backlog MCP unavailable in this session, so this task is tracked via repo-local backlog files.

Implemented renderer digit extraction from `KeyboardEvent.code` for pending numeric selection, so shifted number-row events such as `Ctrl+Shift+Digit3` still dispatch count `3`. Updated the mpv plugin session-binding numeric selector to register bare digits plus the starter shortcut modifier combinations, so plugin-owned `Ctrl+Shift+S` can accept a follow-up digit before the modifiers are fully released.

Verification:

- `bun test src/renderer/handlers/keyboard.test.ts src/core/services/overlay-shortcut-handler.test.ts src/core/services/overlay-window.test.ts`
- `bun run test:plugin:src`
- `bun run changelog:lint`
- `bun x prettier --check src/renderer/handlers/keyboard.ts src/renderer/handlers/keyboard.test.ts package.json 'changes/309-multi-mine-modified-digits.md' 'backlog/tasks/task-309 - Accept-modified-follow-up-digits-for-multi-line-sentence-mining.md'`

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Restored multi-line sentence-card digit capture for the case where `Ctrl`/`Shift` are still held after `Ctrl+Shift+S`. The renderer now accepts digits by physical `Digit1`-`Digit9`/`Numpad1`-`Numpad9` code during pending numeric selection, and the mpv plugin registers the matching modified digit bindings for session-binding numeric prompts.

<!-- SECTION:FINAL_SUMMARY:END -->
