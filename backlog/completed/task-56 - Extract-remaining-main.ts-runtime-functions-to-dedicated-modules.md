---
id: TASK-56
title: Extract remaining main.ts runtime functions to dedicated modules
status: Done
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-16 05:16'
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

Goal: Reduce main.ts complexity by extracting focused runtime helpers into dedicated modules

Benefits:
- Faster navigation and comprehension of main.ts
- Easier to test extracted modules independently
- Clearer separation of concerns
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Extract JLPT dictionary lookup functions to dedicated module
- [x] #2 Extract media path utilities to dedicated module
- [x] #3 Extract overlay visibility helpers to dedicated module
- [x] #4 Update main.ts imports to use new modules
- [x] #5 Ensure all functionality remains intact
- [x] #6 Run full test suite
- [x] #7 Keep extracted code organized and easier to follow
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refactor complete for targeted runtime extraction: JLPT lookup, media utilities, and overlay visibility helpers were moved into dedicated main-runtime modules and wired from main.ts. Existing behavior preserved and full typecheck + test suite passed.

Task intent updated to prioritize readability over strict line-count target.
<!-- SECTION:FINAL_SUMMARY:END -->
