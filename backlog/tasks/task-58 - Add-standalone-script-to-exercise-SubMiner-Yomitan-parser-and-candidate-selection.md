---
id: TASK-58
title: >-
  Add standalone script to exercise SubMiner Yomitan parser and candidate
  selection
status: Done
assignee: []
created_date: '2026-02-16 22:04'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
ordinal: 27000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Create `scripts/test-yomitan-parser.ts` as a standalone CLI tool that reuses SubMiner's Yomitan parser logic to inspect parse output and candidate selection behavior for test inputs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A new script exists at `scripts/test-yomitan-parser.ts`.
- [x] #2 The script can be run standalone and accepts input text for parsing.
- [x] #3 The script uses existing SubMiner parser logic rather than duplicating parser behavior.
- [x] #4 The script prints parsed results and candidate selection details (including which candidate is chosen).
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a standalone parser test script at `scripts/test-yomitan-parser.ts` that reuses `tokenizeSubtitleService` from SubMiner, initializes optional Yomitan+Electron runtime, fetches raw parse candidates via Yomitan `parseText`, and reports which candidate(s) match the final selected token output. Added package scripts `test-yomitan-parser` and `test-yomitan-parser:electron` for direct and Electron-backed runs.
<!-- SECTION:FINAL_SUMMARY:END -->
