---
id: TASK-113
title: Refresh architecture doc content to match current codebase
status: Done
assignee:
  - codex-architecture-doc-refresh
created_date: '2026-02-22 19:39'
updated_date: '2026-02-22 19:44'
labels:
  - docs
  - architecture
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User requested a thorough review of the current codebase and `docs/architecture.md`, then updating document content to reflect current implementation details. Mermaid/flow diagrams are explicitly out of scope for this pass.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Architecture content sections align with current source tree/module ownership.
- [x] #2 Service/composer/runtime descriptions are updated for current behavior and module names.
- [x] #3 Mermaid/flow chart sections remain unchanged in this pass.
<!-- AC:END -->

## Notes

- Reviewed current module ownership across `src/main.ts`, `src/main/runtime/*`, `src/core/services/*`, `src/config/*`, `src/renderer/*`, `launcher/*`, and `plugin/subminer.lua`.
- Updated `docs/architecture.md` content sections only; Mermaid blocks intentionally left untouched.
- Verification: `bun run docs:build` passed.
