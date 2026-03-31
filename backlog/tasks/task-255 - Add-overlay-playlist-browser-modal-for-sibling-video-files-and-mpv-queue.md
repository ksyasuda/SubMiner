---
id: TASK-255
title: Add overlay playlist browser modal for sibling video files and mpv queue
status: Done
assignee:
  - '@codex'
created_date: '2026-03-30 05:46'
updated_date: '2026-03-31 19:37'
labels:
  - feature
  - overlay
  - mpv
  - launcher
dependencies: []
ordinal: 180500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add an in-session overlay modal that opens from a keybinding during active playback and lets the user browse video files from the current file's parent directory alongside the active mpv playlist. The modal should sort local files in best-effort episode order, highlight the current item, and allow keyboard/mouse interaction to add files into the mpv queue, remove queued items, and reorder queued items without leaving playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 An overlay modal can be opened during active playback from a dedicated keybinding and closed without disrupting existing modal behavior.
- [x] #2 The modal shows video files from the current media file's parent directory in best-effort episode order and highlights the current file when present.
- [x] #3 The modal shows the active mpv playlist/queue with enough metadata to identify the current item and queued order.
- [x] #4 The user can add a directory file to the mpv playlist, remove playlist items, and reorder playlist items from the modal using both mouse and keyboard interactions.
- [x] #5 Modal state stays in sync after playlist mutations so the rendered queue reflects mpv's current playlist order.
- [x] #6 Feature coverage includes automated tests for ordering/playlist behavior and docs or shortcut/help updates for the new modal.
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

2026-03-30 current CodeRabbit round: verify 4 unresolved threads, ignore already-fixed outdated dblclick thread if current code matches, add failing-first coverage for selection preservation / timestamp fixture consistency / string test-clock alignment, implement minimal fixes, rerun targeted tests plus typecheck.

2026-03-30 latest CodeRabbit round on PR #37: 1) add failing coverage for negative fractional numeric __subminerTestNowMs input so nowMs() matches the string-backed path, 2) add failing coverage that playlist-browser modal tests restore absent window/document globals without leaving undefined-valued properties behind, 3) refactor repeated playlist-browser modal test harness into a shared setup/teardown fixture while preserving assertions, 4) implement minimal fixes, 5) rerun touched tests plus typecheck.

2026-03-30 latest CodeRabbit follow-up after ff760ea: tighten the new cleanup regression so env.restore() always runs under assertion failure, and make the keydown test's append mock return a post-append mutated snapshot before exercising Ctrl+ArrowDown. Re-run targeted playlist-browser tests plus typecheck.
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

Addressed CodeRabbit follow-ups on the playlist browser PR: clamped stale playingIndex values, failed mutation paths when MPV rejects send(), added temp-dir cleanup in runtime tests, and blocked action-button dblclick bubbling in the renderer. Verification: `bun run typecheck`, `bun run build`, `bun test src/main/runtime/playlist-browser-runtime.test.ts src/renderer/modals/playlist-browser.test.ts`.

Additional follow-up: moved playlist-browser keydown handling ahead of keyboard-driven lookup controls so KeyH/ArrowLeft/ArrowRight and related chords are routed to the modal first. Verification refreshed with `bun test src/main/runtime/playlist-browser-runtime.test.ts src/renderer/modals/playlist-browser.test.ts src/renderer/handlers/keyboard.test.ts`, `bun run typecheck`, and `bun run build`.

Split playlist-browser UI row rendering into `src/renderer/modals/playlist-browser-renderer.ts` and left `src/renderer/modals/playlist-browser.ts` as the controller/wiring layer. Moved playlist-browser IPC/runtime wiring into `src/main/runtime/playlist-browser-ipc.ts` and collapsed the `src/main.ts` registration block to use that helper. Verification after refactor: `bun run typecheck`, `bun run build`, `bun test src/main/runtime/playlist-browser-runtime.test.ts src/renderer/modals/playlist-browser.test.ts src/renderer/handlers/keyboard.test.ts`.

2026-03-30 PR #37 unresolved CodeRabbit threads currently reduce to three likely-actionable items plus one outdated renderer dblclick thread to verify against HEAD before touching code.

2026-03-30 Addressed latest unresolved CodeRabbit items on PR #37: preserved playlist-browser selection across mutation snapshots, taught nowMs() to honor string-backed test clocks so it stays aligned with currentDbTimestamp(), and normalized maintenance test timestamp fixtures to toDbTimestamp(). The older playlist-browser dblclick thread remains unresolved in GitHub state but current HEAD already contains that fix in playlist-browser-renderer.ts.

2026-03-30 latest CodeRabbit remediation on PR #37: switched nowMs() numeric test-clock branch from Math.floor() to Math.trunc() so numeric and string-backed mock clocks agree for negative fractional values. Refactored playlist-browser modal tests onto a shared setup/teardown fixture that restores global window/document descriptors correctly, and added regression coverage that injected globals are deleted when originally absent. Verification: `bun test src/core/services/immersion-tracker/time.test.ts src/renderer/modals/playlist-browser.test.ts`, `bun run typecheck`.

2026-03-30 CodeRabbit follow-up: wrapped the injected-globals cleanup regression in try/finally so restore always runs, and changed the keydown test append mock to return createMutationSnapshot() before exercising Ctrl+ArrowDown. Verified with `bun test src/renderer/modals/playlist-browser.test.ts` and `bun run typecheck`.

2026-03-31 assessment: the playlist-browser feature is landed on `main` via `d51e7fe4 Add playlist browser overlay modal (#37)` with runtime, IPC, renderer, keybinding, and changelog/docs coverage present. Verified passes: `bun test src/main/runtime/playlist-browser-runtime.test.ts src/main/runtime/playlist-browser-open.test.ts src/main/runtime/playlist-browser-sort.test.ts src/renderer/handlers/keyboard.test.ts src/core/services/ipc.test.ts src/core/services/ipc-command.test.ts src/config/definitions/domain-registry.test.ts`.

Remaining action item before close: fix `src/renderer/modals/playlist-browser.test.ts` so the cleanup regression does not assume `globalThis.window` / `globalThis.document` start absent under Bun, rerun the playlist-browser modal lane (and then typecheck/build if you want the full closeout proof), then finalize the task.
<!-- SECTION:NOTES:END -->
