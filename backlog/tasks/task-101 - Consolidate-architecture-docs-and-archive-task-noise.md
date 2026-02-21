---
id: TASK-101
title: Consolidate architecture docs and archive task noise
status: To Do
assignee: []
created_date: '2026-02-21 07:15'
updated_date: '2026-02-21 07:15'
labels:
  - documentation
  - maintainability
dependencies:
  - TASK-96
  - TASK-97
  - TASK-98
  - TASK-99
  - TASK-100
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Architecture guidance is fragmented across long-lived docs and task notes. Consolidate canonical runtime architecture docs and move stale/noisy task-level detail to archive.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Audit architecture and development docs for duplicated runtime composition guidance.
2. Select one canonical architecture section for runtime composition and dependency boundaries.
3. Refactor duplicate sections to brief pointers to the canonical section.
4. Move stale task-evidence prose from persistent docs to backlog/archive where appropriate.
5. Add a compact architecture diagram and update links in contributor docs.
6. Validate docs build and link integrity.
7. Record what moved and why in task notes for traceability.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Runtime composition guidance exists in one canonical location.
- [ ] #2 Duplicated/stale architecture notes removed from long-lived docs.
- [ ] #3 Task evidence retained in backlog/archive, not lost.
- [ ] #4 Docs build and links pass after consolidation.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Change log of moved/removed doc sections included in task notes.
- [ ] #2 `bun run docs:build` passes.
- [ ] #3 Contributor-facing entry points link to canonical architecture section.
<!-- DOD:END -->

