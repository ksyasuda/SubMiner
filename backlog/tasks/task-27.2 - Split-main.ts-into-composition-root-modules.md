---
id: TASK-27.2
title: Split main.ts into composition-root modules
status: To Do
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-14 02:23'
labels:
  - 'owner:backend'
  - 'owner:architect'
dependencies:
  - TASK-27.1
  - TASK-7
references:
  - src/main.ts
documentation:
  - docs/architecture.md
parent_task_id: TASK-27
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce main.ts complexity by extracting bootstrap, lifecycle, overlay, IPC, and CLI wiring into explicit modules while keeping runtime behavior unchanged.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Create modules under src/main/ for bootstrap/lifecycle/ipc/overlay/cli concerns.
- [ ] #2 main.ts no longer owns session-specific business state; it only composes services and starts the app.
- [ ] #3 Public service behavior, startup order, and flags remain unchanged, validated by existing integration/manual smoke checks.
- [ ] #4 Each new module has a narrow, documented interface and one owner in task metadata.
- [ ] #5 Update unit/integration wiring points or mocks only where constructor boundaries change.
- [ ] #6 Add a migration note in docs/structure-roadmap.md.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Dependency Context

**TASK-7 is a hard prerequisite.** main.ts currently has 30+ module-level `let` declarations (mpvClient, yomitanExt, reconnectTimer, currentSubText, subtitlePosition, keybindings, ankiIntegration, secondarySubMode, etc.). Splitting main.ts into src/main/ submodules without first consolidating this state into a typed AppState container would scatter mutable state across files, making data flow even harder to trace.

**TASK-9 (remove trivial wrappers)** depends on TASK-7 and should ideally complete before this task starts, since it reduces the surface area of main.ts. However, it's not a hard blocker — wrapper removal can happen during or after the split.

**Sequencing note:** This task should run AFTER TASK-27.3 (anki-integration split) completes, because AnkiIntegration is instantiated and heavily wired in main.ts. Changing both the composition root and the Anki facade simultaneously creates integration risk. Let 27.3 stabilize the Anki module boundaries first, then split main.ts around the stable API.

## Folded-in work from TASK-9 and TASK-10
TASK-9 (remove trivial wrappers) and TASK-10 (naming conventions) have been deprioritized to low. Their scope is largely subsumed by this task:
- When main.ts is split into composition-root modules, trivial wrappers will naturally be eliminated or inlined at each module boundary.
- Naming conventions should be standardized per-module as they are extracted, not as a separate global pass.
Refer to TASK-9 and TASK-10 acceptance criteria as checklists during execution of this task.
<!-- SECTION:NOTES:END -->
