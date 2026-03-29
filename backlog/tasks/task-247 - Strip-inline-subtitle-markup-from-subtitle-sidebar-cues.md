---
id: TASK-247
title: Strip inline subtitle markup from subtitle sidebar cues
status: Done
assignee:
  - codex
created_date: '2026-03-29 10:01'
updated_date: '2026-03-29 10:10'
labels: []
dependencies: []
references:
  - src/core/services/subtitle-cue-parser.ts
  - src/renderer/modals/subtitle-sidebar.ts
  - src/core/services/subtitle-cue-parser.test.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle sidebar should display readable subtitle text when loaded subtitle files include inline markup such as HTML-like font tags. Parsed cue text currently preserves markup, causing raw tags to appear in the sidebar instead of clean subtitle content.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Subtitle sidebar cue text omits inline subtitle markup such as HTML-like font tags while preserving visible subtitle content.
- [x] #2 Parsed subtitle cues used by the sidebar keep timing order and expected line-break behavior after markup sanitization.
- [x] #3 Regression tests cover markup-bearing subtitle cue parsing so raw tags do not reappear in the sidebar.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression tests in src/core/services/subtitle-cue-parser.test.ts for subtitle cues containing HTML-like font tags, including multi-line content.
2. Verify the new parser test fails against current behavior to confirm the bug is covered.
3. Update src/core/services/subtitle-cue-parser.ts to sanitize inline subtitle markup while preserving visible text and expected newline handling.
4. Re-run focused parser tests, then run broader verification commands required for handoff as practical.
5. Update task notes/acceptance criteria based on verified results and finalize the task record.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User approved implementation on 2026-03-29.

Implemented parser-level subtitle cue sanitization for HTML-like tags so loaded sidebar cues render readable text while preserving cue line breaks.

Added regression coverage for SRT and ASS cue parsing with <font ...> markup.

Verification: bun test src/core/services/subtitle-cue-parser.test.ts; bun run typecheck; bun run test:fast; bun run test:env; bun run build; bun run test:smoke:dist.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Sanitized parsed subtitle cue text in src/core/services/subtitle-cue-parser.ts so HTML-like inline markup such as <font ...> is removed before cues reach the subtitle sidebar. The sanitizer is shared across SRT/VTT-style parsing and ASS parsing, while existing cue timing and line-break semantics remain intact.

Added regression tests in src/core/services/subtitle-cue-parser.test.ts covering markup-bearing SRT lines and ASS dialogue lines with \N breaks, and verified the original failure before implementing the fix.

Tests run: bun test src/core/services/subtitle-cue-parser.test.ts; bun run typecheck; bun run test:fast; bun run test:env; bun run build; bun run test:smoke:dist.
<!-- SECTION:FINAL_SUMMARY:END -->
