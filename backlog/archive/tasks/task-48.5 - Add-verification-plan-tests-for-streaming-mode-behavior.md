---
id: TASK-48.5
title: Add verification plan/tests for streaming mode behavior
status: To Do
assignee: []
created_date: '2026-02-14 06:03'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - qa
dependencies: []
parent_task_id: TASK-48
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Create a validation plan or tests for CLI flag behavior, stream resolution, and subtitle precedence/fallback rules so streaming mode changes are measurable and regressions are caught.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Document/manual checklist covers `-s` and `--stream` invocation and streaming-only behavior.
- [ ] #2 Include cases for (a) stream with JP subtitles, (b) no JP subtitles with Jimaku fallback, (c) primary-language not Japanese.
- [ ] #3 Run or provide reproducible checks to confirm non-streaming behavior unchanged.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Superseded by TASK-51. Verification plan moved to deferred reimplementation context.

<!-- SECTION:NOTES:END -->
