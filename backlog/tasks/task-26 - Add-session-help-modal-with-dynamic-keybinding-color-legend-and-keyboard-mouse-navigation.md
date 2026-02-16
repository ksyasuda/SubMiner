---
id: TASK-26
title: >-
  Add session help modal with dynamic keybinding/color legend and keyboard/mouse
  navigation
status: Done
assignee: []
created_date: '2026-02-13 16:49'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Create a help modal that auto-generates its content from the project/app layout and current configuration, showing the active keybindings and color keys for the current session in a presentable way. The modal should be navigable with arrow keys and Vim-style hjkl-style movement behavior internally but without labeling hjkl in UI, and support mouse interaction. Escape closes/goes back. Open the modal with `y-h`, and if that binding is already taken, use `y-k` as fallback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Help modal content is generated automatically from current keybinding config and project/app layout rather than hardcoded static text.
- [x] #2 Modal displays current session keybindings and active color-key mappings in a clear, grouped layout with section separation for readability.
- [x] #3 Modal can be opened with `y-h` when available; if `y-h` is already bound, modal opens with `y-k` instead.
- [x] #4 Close behavior: `Escape` exits the modal and returns to previous focus/state.
- [x] #5 Modal supports mouse-based interaction for standard focus/selection actions.
- [x] #6 Navigation inside modal supports arrow keys for movement between focusable items.
- [x] #7 Modal supports internal navigation semantics equivalent to hjkl directional movement for users, while UI text/labels do not mention hjkl keys.
- [x] #8 No visible UI mention of Vim key names is shown in modal labels/help copy.
- [x] #9 Modal opens from current session without requiring a restart and reflects updated config changes without code changes.
- [x] #10 If the shortcut is unavailable due to conflicts, user-visible fallback behavior/error is deterministic and documented.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Auto-generated help modal displays up-to-date keybinding + color mapping data and supports both keyboard (arrow/fallback path) and mouse navigation with Escape-to-close.
<!-- DOD:END -->
