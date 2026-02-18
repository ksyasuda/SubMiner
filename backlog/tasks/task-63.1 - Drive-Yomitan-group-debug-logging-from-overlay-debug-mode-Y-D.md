---
id: TASK-63.1
title: Drive Yomitan group debug logging from overlay debug mode (Y-D)
status: Done
assignee: []
created_date: '2026-02-16 23:48'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
parent_task_id: TASK-63
priority: low
ordinal: 21000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Remove dedicated runtime/config toggle for Yomitan group logging and instead enable logs only when overlay debug mode is active via the existing Y-D debug flow.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 No runtime option or config key is required for Yomitan group logging.
- [x] #2 Yomitan group logs are emitted only when overlay debug mode is enabled (Y-D/DevTools debug state).
- [x] #3 When overlay debug mode is disabled, Yomitan group logs are not emitted.
- [x] #4 Build/tests for touched modules pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Remove the `debugYomitanGroups` runtime/config option wiring from types/config registries so it no longer appears in runtime options.
2. Keep tokenizer-level debug logging gate but drive it from existing overlay debug state (`overlayDebugVisualizationEnabled`) which is toggled by Y-D/DevTools flow.
3. Rebuild and run tokenizer/config/runtime-option tests to confirm behavior and no registry regressions.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Removed `anki.debugYomitanGroups` from runtime option ID union and removed `ankiConnect.behavior.debugYomitanGroups` from config defaults/registry entries.

Updated `main.ts` tokenizer dependency wiring so `getYomitanGroupDebugEnabled` now directly reads `appState.overlayDebugVisualizationEnabled` (the existing debug visualization state toggled via Y-D/DevTools).

Validated with `pnpm run build && node dist/core/services/tokenizer-service.test.js && node --test dist/config/config.test.js dist/core/services/runtime-options-ipc-service.test.js`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Switched Yomitan group debug logging to follow the existing overlay debug mode (Y-D/DevTools state) and removed the dedicated runtime/config option surface. The tokenizer still logs `Selected Yomitan token groups` only when the debug gate is true, but the gate now comes from `appState.overlayDebugVisualizationEnabled` in main runtime wiring. Removed the temporary runtime-option/config definitions and updated related registry expectations. Build and relevant tests are passing.

<!-- SECTION:FINAL_SUMMARY:END -->
