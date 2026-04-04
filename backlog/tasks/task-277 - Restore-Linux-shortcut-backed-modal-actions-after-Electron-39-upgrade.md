---
id: TASK-277
title: Restore Linux shortcut-backed modal actions after Electron 39 upgrade
status: Done
assignee:
  - codex
created_date: '2026-04-04 06:49'
updated_date: '2026-04-04 07:08'
labels:
  - bug
  - linux
  - electron
  - shortcuts
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Linux builds on Electron 39 no longer open the runtime options, Jimaku, or Subsync flows from their configured shortcuts (`Ctrl/Cmd+Shift+O`, `Ctrl+Shift+J`, `Ctrl+Alt+S`). Other renderer-driven modals like session help, subtitle sidebar, and playlist browser still work, which points to the Linux `globalShortcut` / overlay shortcut path rather than modal rendering in general. Investigate the Electron 39 regression and restore these Linux shortcut-triggered actions without regressing macOS or Windows behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 On Linux, the configured runtime options, Jimaku, and Subsync shortcuts trigger their existing actions again under Electron 39.
- [x] #2 The fix is covered by automated tests around the Linux shortcut/backend behavior that regressed.
- [x] #3 A changelog fragment is added for the Linux shortcut regression fix.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Extend the special-command / mpv IPC command path to cover Jimaku alongside the existing runtime-options and subsync commands.
2. Add tests for IPC command dispatch and any keybinding/session-help metadata affected by the new special command.
3. Route Linux modal-opening shortcuts through the working mpv/IPC command path instead of depending on Electron globalShortcut delivery for those actions.
4. Re-run targeted tests, then the handoff verification gate, and update the changelog/task summary to reflect the final root cause and fix.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User approved plan; proceeding with test-first implementation.

Root cause: Linux startup unconditionally enabled Electron's GlobalShortcutsPortal path, so Electron 39 X11 sessions were routed through the portal-backed globalShortcut path even though that path should only be used for Wayland-style launches. The affected runtime-options, Jimaku, and Subsync actions all depend on that shortcut path, while renderer-driven modals did not regress.

User retested after the portal gating fix; issue persists. Updating investigation scope from portal selection alone to Linux overlay shortcut delivery/registration under Electron 39.

Revised fix path: Linux renderer now matches configured overlay shortcuts for runtime options, subsync, and Jimaku locally, then dispatches the existing mpv/IPC special-command route instead of depending on Electron globalShortcut delivery.

Added Jimaku mpv special command plumbing through IPC/runtime deps and exposed an overlay-runtime helper for opening the Jimaku modal from that IPC path.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored Linux shortcut-backed modal actions by routing runtime options, Subsync, and Jimaku through the working mpv/IPC special-command path. Added renderer-side Linux shortcut matching for configured overlay shortcuts, threaded a new Jimaku special command through IPC/runtime deps, refreshed configured shortcuts on hot reload, and covered the regression with IPC/runtime keyboard tests. Verification: bun test src/core/services/ipc-command.test.ts src/renderer/handlers/keyboard.test.ts src/main/runtime/ipc-mpv-command-main-deps.test.ts; bun run typecheck; bun run test:fast; bun run test:env; bun run build; bun run test:smoke:dist.
<!-- SECTION:FINAL_SUMMARY:END -->
