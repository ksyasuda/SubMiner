---
id: TASK-255
title: Add overlay playlist browser modal for sibling video files and mpv queue
status: In Progress
assignee:
  - codex
created_date: '2026-03-30 05:46'
updated_date: '2026-03-30 09:22'
labels:
  - feature
  - overlay
  - mpv
  - launcher
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add an in-session overlay modal that opens from a keybinding during active playback and lets the user browse video files from the current file's parent directory alongside the active mpv playlist. The modal should sort local files in best-effort episode order, highlight the current item, and allow keyboard/mouse interaction to add files into the mpv queue, remove queued items, and reorder queued items without leaving playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 An overlay modal can be opened during active playback from a dedicated keybinding and closed without disrupting existing modal behavior.
- [ ] #2 The modal shows video files from the current media file's parent directory in best-effort episode order and highlights the current file when present.
- [ ] #3 The modal shows the active mpv playlist/queue with enough metadata to identify the current item and queued order.
- [ ] #4 The user can add a directory file to the mpv playlist, remove playlist items, and reorder playlist items from the modal using both mouse and keyboard interactions.
- [ ] #5 Modal state stays in sync after playlist mutations so the rendered queue reflects mpv's current playlist order.
- [ ] #6 Feature coverage includes automated tests for ordering/playlist behavior and docs or shortcut/help updates for the new modal.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add playlist-browser domain types, IPC channels, overlay modal registration, special command, and default keybinding for Ctrl+Alt+P.
2. Write failing tests for best-effort episode sorting and main playlist-browser runtime snapshot/mutation behavior.
3. Implement playlist-browser main/runtime helpers for local sibling video discovery, mpv playlist normalization, and append/play/remove/move operations with refreshed snapshots.
4. Wire preload and main-process IPC handlers that expose snapshot and mutation methods to the renderer.
5. Write failing renderer and keyboard tests for modal open/close, split-pane interaction, keyboard controls, and degraded states.
6. Implement playlist-browser modal markup, DOM/state, renderer composition, keyboard routing, and session-help labeling.
7. Run targeted test lanes first, then the maintained verification gate relevant to the touched surfaces; update task notes/criteria as checks pass.

2026-03-30 CodeRabbit follow-up: 1) add failing runtime coverage for unreadable playlist-browser file stat failures, 2) add failing renderer coverage for stale snapshot UI reset on refresh failure/close, 3) add failing renderer coverage to block playlist-browser open when another modal already owns the overlay, 4) implement minimal fixes, 5) rerun targeted tests plus typecheck for touched surfaces.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented overlay playlist browser modal with split directory/playlist panes, Ctrl+Alt+P keybinding, main/preload IPC, mpv queue mutations, and best-effort sibling episode sorting.

Added tests for sort/runtime logic, IPC wiring, keyboard routing, and playlist-browser modal behavior.

Verification: `bun run typecheck` passed; targeted playlist-browser and IPC tests passed; `bun run build` passed; `bun run test:smoke:dist` passed.

Repo gate blockers outside this feature: `bun run test:fast` hits existing Bun `node:test` NotImplementedError cases plus unrelated immersion-tracker failures; `bun run test:env` fails in existing immersion-tracker sqlite tests.

2026-03-30: Fixed playlist-browser local playback regression where subtitle track IDs leaked across episode jumps. `playPlaylistBrowserIndexRuntime` now reapplies local subtitle auto-selection defaults (`sub-auto=fuzzy`, `sid=auto`, `secondary-sid=auto`) before `playlist-play-index` for local filesystem targets only; remote playlist entries remain untouched. Added runtime regression tests for both paths.

2026-03-30: Follow-up subtitle regression fix. Pre-jump `sid=auto` was ineffective because mpv resolved it against the current episode before `playlist-play-index`. Local playlist jumps now set `sub-auto=fuzzy`, switch episodes, then schedule a delayed rearm of `sid=auto` and `secondary-sid=auto` so selection happens against the new file's tracks. Added failing-first runtime coverage for delayed local rearm and remote no-op behavior.

2026-03-30: Cleaned up playlist-browser runtime local-play subtitle-rearm flow by extracting focused helpers without changing behavior. Added public docs/readme coverage for the default `Ctrl+Alt+P` playlist browser keybinding and modal, plus changelog fragment `changes/260-playlist-browser.md`. Verification: `bun test src/main/runtime/playlist-browser-runtime.test.ts`, `bun run typecheck`, `bun run docs:test`, `bun run docs:build`, `bun run changelog:lint`, `bun run build`.

2026-03-30: Pulled unresolved CodeRabbit review threads for PR #37. Actionable set is three items: unreadable-file stat error handling in playlist-browser runtime, stale playlist-browser DOM after failed refresh/close, and missing modal-ownership guard before opening the playlist-browser overlay. Proceeding test-first for each.

2026-03-30: Addressed current CodeRabbit follow-up findings for PR #37. Fixed playlist-browser unreadable-file stat handling, stale playlist-browser DOM reset on refresh failure/close, modal-ownership guard before opening the playlist-browser overlay, async rejection surfacing for PLAYLIST_BROWSER_OPEN IPC commands, overlay bootstrap before playlist-browser open dispatch, texthooker option normalization in the mpv plugin, and superseded local subtitle-rearm suppression. Added targeted regressions plus new playlist-browser-open helper coverage. Verification: `bun test src/main/runtime/playlist-browser-runtime.test.ts src/main/runtime/playlist-browser-open.test.ts src/core/services/ipc-command.test.ts src/renderer/modals/playlist-browser.test.ts`, `lua scripts/test-plugin-start-gate.lua`, `bun run typecheck`, `bun run build`.
<!-- SECTION:NOTES:END -->
