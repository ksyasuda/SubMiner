---
id: TASK-19
title: Enable overlay keybinds whenever app runtime is active
status: To Do
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-12 09:40'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Restored task after accidental cleanup. Ensure keybindings are available whenever the Electron runtime is active, while respecting focused overlay/input contexts that require local key handling.

<!-- SECTION:DESCRIPTION:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Started implementation: tracing overlay shortcut registration lifecycle and runtime activation gating.

Root cause: overlay shortcut sync executes during overlay runtime initialization before overlayRuntimeInitialized is set true, so registration could be skipped until a later visibility toggle.

Implemented fix in initializeOverlayRuntime: after setting overlayRuntimeInitialized = true, immediately call syncOverlayShortcuts() to register overlay keybinds for active runtime state.

Follow-up from repro: startup path with --start did not initialize overlay runtime (commandNeedsOverlayRuntime excluded start), so overlay keybinds stayed unavailable until first overlay visibility command.

Updated CLI runtime gating so --start initializes overlay runtime, which activates overlay shortcut registration immediately.

User clarified MPV-only workflow requirement. Added MPV plugin keybindings and script messages for mining/runtime actions (copy/mine/multi/mode/field-grouping/subsync/audio-card/runtime-options) so these actions are available from mpv chord bindings without relying on overlay global shortcuts.

Per user direction, reverted all shortcut/runtime/plugin changes from this implementation cycle. Desired behavior is to keep keybindings working only when overlay is active.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Ensured overlay shortcuts are available as soon as overlay runtime becomes active by resyncing after activation flag is set. This prevents startup states where shortcuts remained inactive until a later overlay visibility change.

Follow-up fix: included --start in overlay-runtime-required commands so keybinds are active right after startup, even before toggling visible/invisible overlays.

<!-- SECTION:FINAL_SUMMARY:END -->
