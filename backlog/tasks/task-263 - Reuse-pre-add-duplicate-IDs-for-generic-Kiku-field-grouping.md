---
id: TASK-263
title: Reuse pre-add duplicate IDs for generic Kiku field grouping
status: Done
assignee: []
created_date: '2026-03-31 20:44'
updated_date: '2026-03-31 20:48'
labels:
  - anki
  - kiku
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Avoid the extra post-add duplicate lookup on the generic sentence-card creation path by capturing duplicate note IDs before add and reusing that result for Kiku field grouping. Keep Yomitan semantics aligned where practical so duplicate selection is consistent across mining paths.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Generic sentence-card creation captures duplicate note IDs before add and reuses them for Kiku field grouping instead of running the existing post-add duplicate finder
- [x] #2 Duplicate selection remains deterministic when multiple matching notes exist
- [x] #3 Regression tests cover the generic path duplicate reuse behavior and preserve existing non-Kiku behavior
- [x] #4 Internal docs/config comments are updated if the behavior or operator-facing semantics changed
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
No docs update was required because this is internal duplicate-selection plumbing and does not change user-facing config surface.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Generic sentence-card creation now captures exact duplicate note IDs before add when Kiku field grouping is enabled and stores that context by created note ID. Manual field grouping reuses the tracked duplicate IDs first and deterministically picks the most recent matching note, falling back to the legacy duplicate finder only when no tracked context exists. Verified with bun test src/anki-integration/duplicate.test.ts src/anki-integration/card-creation.test.ts src/anki-integration/field-grouping.test.ts and bun run typecheck.
<!-- SECTION:FINAL_SUMMARY:END -->
