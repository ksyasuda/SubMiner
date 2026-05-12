---
id: TASK-343
title: Fix macOS character dictionary selector session shortcut
status: Done
assignee: []
created_date: '2026-05-11 08:05'
updated_date: '2026-05-11 08:06'
labels:
  - bug
  - macos
  - character-dictionary
  - plugin
dependencies: []
modified_files:
  - plugin/subminer/session_bindings.lua
  - scripts/test-plugin-session-bindings.lua
priority: medium
ordinal: 181500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Opening the character dictionary AniList selector from mpv/session shortcuts should work on macOS. Current generated session bindings include the openCharacterDictionary session action, but the Lua plugin CLI dispatch table does not map that action to the app flag, so the shortcut cannot reach the runtime selector.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The openCharacterDictionary session action invokes the app with --open-character-dictionary from the mpv plugin.
- [x] #2 Regression coverage proves the Lua session-binding CLI map forwards openCharacterDictionary correctly.
- [x] #3 Existing session-binding regression coverage still passes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
TDD red: `lua scripts/test-plugin-session-bindings.lua` failed because openCharacterDictionary did not emit --open-character-dictionary. Green after adding the missing Lua CLI mapping.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the mpv plugin session action mapping so the character dictionary selector shortcut dispatches `--open-character-dictionary` to the app. Added Lua regression coverage for the macOS-style Alt+Meta+A binding and verified adjacent TypeScript session binding tests.
<!-- SECTION:FINAL_SUMMARY:END -->
