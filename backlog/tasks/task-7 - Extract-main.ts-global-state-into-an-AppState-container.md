---
id: TASK-7
title: Extract main.ts global state into an AppState container
status: In Progress
assignee: []
created_date: '2026-02-11 08:20'
updated_date: '2026-02-14 08:45'
labels:
  - refactor
  - main
  - architecture
milestone: Codebase Clarity & Composability
dependencies: []
references:
  - src/main.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
main.ts has 30+ module-level `let` declarations holding all application state: mpvClient, yomitanExt, yomitanSettingsWindow, yomitanParserWindow, reconnectTimer, currentSubText, currentSubAssText, subtitlePosition, currentMediaPath, mecabTokenizer, keybindings, ankiIntegration, secondarySubMode, runtimeOptionsManager, overlayRuntimeInitialized, subsyncInProgress, shortcutsRegistered, etc.

These are mutated through closures from many different functions, making state changes hard to trace and impossible to test.

Consolidate into a typed AppState object (or small set of domain-specific state containers) that provides controlled access to state. This also enables passing state to services without building 50-property dependency objects.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 All mutable state consolidated into typed container(s)
- [ ] #2 No bare `let` declarations at module scope for application state
- [ ] #3 State access goes through the container rather than closures
- [ ] #4 Dependency objects for services shrink significantly (reference the container instead)
- [ ] #5 TypeScript compiles cleanly
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Started centralizing module-level application state in `src/main.ts` via `appState` container and routing most state reads/writes through it. Initial rewrite completed; behavior verification pending and dependency-surface shrink pass still needed.

Implemented Task-7 state migration to `appState` in main.ts and removed module-scope mutable state declarations; fixed a broken regression where several appState references were left as bare expressions in object literals (e.g., `appState.autoStartOverlay`), restoring valid typed dependency construction.
<!-- SECTION:NOTES:END -->
