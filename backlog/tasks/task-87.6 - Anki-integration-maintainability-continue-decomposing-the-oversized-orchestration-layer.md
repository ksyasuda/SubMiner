---
id: TASK-87.6
title: >-
  Anki integration maintainability: continue decomposing the oversized
  orchestration layer
status: To Do
assignee: []
created_date: '2026-03-06 03:20'
updated_date: '2026-03-06 03:21'
labels:
  - tech-debt
  - anki
  - maintainability
milestone: m-0
dependencies:
  - TASK-87.1
references:
  - src/anki-integration.ts
  - src/anki-integration/field-grouping-workflow.ts
  - src/anki-integration/note-update-workflow.ts
  - src/anki-integration/card-creation.ts
  - src/anki-integration/anki-connect-proxy.ts
  - src/anki-integration.test.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
  - docs/anki-integration.md
parent_task_id: TASK-87
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

src/anki-integration.ts remains an oversized orchestration file even after earlier extractions. It still mixes config normalization, polling setup, media generation, duplicate resolution, field grouping workflows, and user feedback coordination in one class. This task should continue the decomposition so the remaining orchestration surface is smaller and easier to reason about, while preserving existing Anki, proxy, field grouping, and note update behavior.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 The responsibilities currently concentrated in src/anki-integration.ts are split into clearer modules or services with narrow ownership boundaries.
- [ ] #2 The resulting orchestration surface is materially smaller and easier to review, with at least one mixed-responsibility cluster extracted behind a well-named interface.
- [ ] #3 Existing Anki integration behavior remains covered by automated verification, including note update, field grouping, and proxy-related flows that the refactor touches.
- [ ] #4 Any developer-facing docs or notes needed to understand the new structure are updated in the same task.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Map the remaining responsibility clusters inside src/anki-integration.ts and choose one or more extraction seams that reduce mixed concerns without changing behavior.
2. Move logic behind narrow interfaces/modules rather than creating another giant helper; keep orchestration readable.
3. Preserve coverage for field grouping, note update, proxy, and card creation flows touched by the refactor.
4. Update docs or internal notes if the new structure changes where contributors should look for a given behavior.
<!-- SECTION:PLAN:END -->
