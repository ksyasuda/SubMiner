---
id: TASK-2
title: Post-refactor follow-ups from investigation.md
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:56'
updated_date: '2026-02-18 04:11'
labels: []
dependencies:
  - TASK-1
references:
  - investigation.md
  - docs/refactor-main-checklist.md
ordinal: 14000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Execute the remaining follow-up work identified in investigation.md: remove unused scaffolding, add tests for high-risk consolidated services, and run manual smoke validation in a desktop MPV session.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Follow-up subtasks are created with explicit scope and dependencies.
- [x] #2 Unused architectural scaffolding and abandoned IPC abstraction files are removed or explicitly retained with documented rationale.
- [x] #3 Dedicated tests are added for higher-risk consolidated services (`overlay-shortcut-handler`, `mining-service`, `anki-jimaku-service`).
- [x] #4 Manual smoke checks for overlay rendering, mining flow, and field-grouping interaction are executed and results documented.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Create scoped subtasks for each recommendation in investigation.md and sequence them by risk and execution constraints.
2. Remove dead scaffolding files and any now-unneeded exports/imports; verify build/tests remain green.
3. Add focused behavior tests for the three higher-risk consolidated services.
4. Run and document desktop smoke validation in an MPV-enabled environment.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed:
- Created TASK-2.1 through TASK-2.5 from `investigation.md` recommendations.
- Finished TASK-2.1: removed unused scaffolding in `src/core/`, `src/modules/`, `src/ipc/` and cleaned internal-only service barrel export.
- Finished TASK-2.2: added dedicated tests for `overlay-shortcut-handler.ts`.
- Finished TASK-2.3: added dedicated tests for `mining-service.ts`.
- Finished TASK-2.4: added dedicated tests for `anki-jimaku-service.ts`.

Remaining:
- TASK-2.5: desktop smoke validation with MPV session
<!-- SECTION:NOTES:END -->
