---
id: TASK-111
title: Fix subtitle-cycle OSD labels for J keybindings
status: Done
assignee:
  - Codex
created_date: '2026-03-07 23:45'
updated_date: '2026-03-16 05:13'
labels: []
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/ipc-command.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/mpv.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/ipc-command.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/mpv-control.test.ts
ordinal: 72500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When cycling subtitle tracks with the default J/Shift+J keybindings, the mpv OSD currently shows raw template text like `${sid}` instead of a resolved subtitle label. Update the keybinding OSD behavior so users see the active subtitle selection clearly when cycling tracks, and ensure placeholder-based OSD messages sent through the mpv client API render correctly.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Pressing the primary subtitle cycle keybinding shows a resolved subtitle label on the OSD instead of a raw `${sid}` placeholder.
- [x] #2 Pressing the secondary subtitle cycle keybinding shows a resolved subtitle label on the OSD instead of a raw `${secondary-sid}` placeholder.
- [x] #3 Proxy OSD messages that rely on mpv property expansion render resolved values when sent through the mpv client API.
- [x] #4 Regression tests cover the subtitle-cycle OSD behavior and the placeholder-expansion OSD path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused failing tests for subtitle-cycle OSD labels and mpv placeholder-expansion behavior.
2. Update the IPC mpv command handler to resolve primary and secondary subtitle track labels from mpv `track-list` data after cycling subtitle tracks.
3. Update the mpv OSD runtime path so placeholder-based `show-text` messages sent through the client API opt into property expansion.
4. Run focused tests, then the relevant core test lane, and record results in the task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Initial triage: `ipc-command.ts` emits raw `${sid}`/`${secondary-sid}` placeholder strings, and `showMpvOsdRuntime` sends `show-text` via mpv client API without enabling property expansion.

User approved implementation plan on 2026-03-07.

Implementation: proxy mpv command OSD now supports an async resolver so subtitle track cycling can show human-readable labels instead of raw `${sid}` placeholders.

Implementation: `showMpvOsdRuntime` now prefixes placeholder-based messages with mpv client-api `expand-properties`, which fixes raw `${...}` OSD output for subtitle delay/position messages.

Testing: `bun test src/core/services/ipc-command.test.ts src/core/services/mpv-control.test.ts src/main/runtime/mpv-proxy-osd.test.ts src/main/runtime/ipc-mpv-command-main-deps.test.ts src/main/runtime/ipc-bridge-actions.test.ts src/main/runtime/ipc-bridge-actions-main-deps.test.ts src/main/runtime/composers/ipc-runtime-composer.test.ts` passed.

Testing: `bun x tsc --noEmit` passed.

Testing: `bun run test:core:src` passed (423 pass, 6 skip, 0 fail).

Docs: no update required because no checked-in docs or help text describe the J/Shift+J OSD output behavior.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed subtitle-cycle OSD handling for the default J/Shift+J keybindings. The IPC mpv command path now supports resolving proxy OSD text asynchronously, and the main-runtime resolver reads mpv `track-list` state so primary and secondary subtitle cycling show human-readable track labels instead of raw `${sid}` / `${secondary-sid}` placeholders.

Also fixed the lower-level mpv OSD transport so placeholder-based `show-text` messages sent through the client API opt into `expand-properties`. That preserves existing template-based OSD messages like subtitle delay and subtitle position without leaking the raw `${...}` syntax.

Added regression coverage for the async proxy OSD path, the placeholder-expansion `showMpvOsdRuntime` path, and the runtime subtitle-track label resolver. Verification run: `bun x tsc --noEmit`; focused mpv/IPC tests; and the maintained `bun run test:core:src` lane (423 pass, 6 skip, 0 fail).
<!-- SECTION:FINAL_SUMMARY:END -->
