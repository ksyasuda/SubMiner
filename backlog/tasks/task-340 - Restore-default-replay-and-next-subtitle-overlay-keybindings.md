---
id: TASK-340
title: Restore default replay and next subtitle overlay keybindings
status: Done
assignee:
  - Codex
created_date: '2026-05-04 06:25'
updated_date: '2026-05-04 06:49'
labels:
  - bug
  - keybindings
  - overlay
  - mpv
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Default overlay/mpv keybindings for replaying the current subtitle line and playing the next subtitle line are not firing. Shift+H and Shift+L subtitle jumps still work, but Ctrl+Shift+H should replay the current subtitle and pause at subtitle end, and Ctrl+Shift+L should play the next subtitle and pause at subtitle end. Keep the other built-in defaults working.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Default keybindings include working replay-current-subtitle and play-next-subtitle bindings on Ctrl+Shift+H and Ctrl+Shift+L.
- [x] #2 Replay-current-subtitle dispatch reaches the existing runtime path that pauses at the subtitle end.
- [x] #3 Play-next-subtitle dispatch reaches the existing runtime path that pauses at the subtitle end.
- [x] #4 Existing default keybindings continue to compile/register without regressions.
- [x] #5 Focused regression tests cover the broken default bindings.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression coverage that the resolved defaults compile on Linux without dropping Ctrl+Shift+H/L, and that those keys map to replayCurrentSubtitle/playNextSubtitle session actions.
2. Move the default session-help shortcut off Ctrl/Cmd+Shift+H to a non-conflicting shortcut, then update generated/default config docs so shipped defaults match documentation.
3. Add/adjust coverage for default replay/next bindings and run targeted Bun tests plus plugin session-binding smoke.

4. Follow-up after live test: fix the mpv plugin shifted-letter key-name conversion so `Ctrl+Shift+KeyL` registers using mpv's uppercase letter form and add Lua regression coverage for both `Ctrl+Shift+L` and `Shift+L`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: default `shortcuts.openSessionHelp = CommandOrControl+Shift+H` canonicalized to `ctrl+shift+KeyH` on Linux/Windows, conflicting with the built-in replay-current-subtitle keybinding. The session-binding compiler drops conflicted bindings, so replay did not register. Moved default session help to `CommandOrControl+Slash` and added regression coverage that defaults compile without a conflict and keep replay/next actions on `Ctrl+Shift+H/L`.

Follow-up from live test: `Ctrl+Shift+H` works after resolving the help shortcut conflict, but `Ctrl+Shift+L` still behaves like native/other `Ctrl+L`. Investigating mpv/plugin key-name generation for shifted letter chords.

Follow-up fix: mpv normalizes shifted letter chords to uppercase letter key names (for example `Ctrl+Shift+l` becomes `Ctrl+L`). The plugin previously emitted `Ctrl+Shift+l`, which let live `Ctrl+Shift+L` fall through as the `Ctrl+L` key path. `plugin/subminer/session_bindings.lua` now emits uppercase letters and omits the Shift modifier for shifted `Key[A-Z]` bindings. Lua regression coverage now checks `Ctrl+Shift+KeyL -> Ctrl+L`, `Shift+KeyL -> L`, and the play-next CLI dispatch.

Second live follow-up: `Ctrl+Shift+L` routed to play-next but still behaved like `Shift+L` when playback was already paused because `MpvIpcClient.playNextSubtitle()` explicitly cleared `pendingPauseAtSubEnd` and only sent `sub-seek 1` in paused state. Changed play-next to always arm pause-at-sub-end, clear stale pause target, seek to next subtitle, and unpause when currently paused. Existing sub-end/time-pos handling then pauses at the next subtitle end.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Changed the default session-help shortcut from `CommandOrControl+Shift+H` to `CommandOrControl+Slash` so `Ctrl+Shift+H` remains available for replay-current-subtitle and `Ctrl+Shift+L` remains available for play-next-subtitle. Updated config examples, docs-site shortcut/config/usage docs, and added changelog fragment `changes/340-default-subtitle-keybindings.md`.

Fixed both follow-up issues from live testing. First, the mpv plugin key-name converter now uses mpv's uppercase key form for shifted letter bindings (`Ctrl+Shift+KeyL` registers as `Ctrl+L`, `Shift+KeyL` as `L`). Second, `MpvIpcClient.playNextSubtitle()` now starts playback even when mpv is paused, keeps the pause-at-sub-end path armed, and lets existing subtitle-end timing pause again at the next subtitle end.

Regression coverage now includes compiled default bindings, Lua plugin shifted-letter registration/CLI dispatch, and paused-state play-next behavior.

Verification passed: targeted Bun session/mpv/protocol tests, `bun run test:plugin:src`, `bun run changelog:lint`, `bun run build`, and `bun run test:smoke:dist`. Earlier full gate also passed before the follow-ups: `bun run typecheck`, `bun run test:fast`, `bun run test:env`, docs/config checks, and dist smoke.
<!-- SECTION:FINAL_SUMMARY:END -->
