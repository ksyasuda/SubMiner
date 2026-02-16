---
id: TASK-56
title: Extract remaining main.ts runtime functions to dedicated modules
status: To Do
assignee: []
created_date: '2026-02-16 04:47'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
main.ts is still 1481 lines after previous refactoring efforts. While significant progress has been made, there are still opportunities to extract runtime functions into dedicated modules to further reduce its size and improve maintainability.

Current opportunities:
1. **JLPT dictionary lookup functions** (lines 470-535) - initializeJlptDictionaryLookup, ensureJlptDictionaryLookup, getJlptDictionarySearchPaths
2. **Media path utilities** (lines 552-590) - updateCurrentMediaPath, updateCurrentMediaTitle, resolveMediaPathForJimaku
3. **Overlay visibility helpers** (lines 1273-1360) - updateVisibleOverlayVisibility, updateInvisibleOverlayVisibility, syncInvisibleOverlayMousePassthrough

These functions are largely self-contained and could be moved to:
- `src/main/jlpt-runtime.ts`
- `src/main/media-runtime.ts`  
- `src/main/overlay-visibility-runtime.ts`

Goal: Reduce main.ts to under 1000 lines (target: ~800-900 lines)

Benefits:
- Faster navigation and comprehension of main.ts
- Easier to test extracted modules independently
- Clearer separation of concerns
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Extract JLPT dictionary lookup functions to dedicated module
- [ ] #2 Extract media path utilities to dedicated module
- [ ] #3 Extract overlay visibility helpers to dedicated module
- [ ] #4 Update main.ts imports to use new modules
- [ ] #5 Ensure all functionality remains intact
- [ ] #6 Run full test suite
- [ ] #7 Verify main.ts line count is reduced to under 1000 lines
<!-- AC:END -->
