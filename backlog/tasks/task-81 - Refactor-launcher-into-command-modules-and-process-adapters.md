---
id: TASK-81
title: Refactor launcher into command modules and process adapters
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - launcher
  - cli
  - architecture
dependencies:
  - TASK-70
  - TASK-73
  - TASK-74
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launcher code is still large and process-control heavy (`launcher/main.ts`, `launcher/mpv.ts`, `launcher/config.ts`). This task refactors launcher flows into command modules plus explicit process adapters to reduce branching complexity and improve testability.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Split by command domain (`doctor`, `config`, `mpv`, `jellyfin`, `play`).
- Isolate shell/process invocations behind adapter interfaces.
- Keep argv parsing independent from command execution side effects.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define launcher command dispatcher and command handler interfaces.
2. Extract command handlers from `launcher/main.ts`.
3. Move spawn/fs/net operations into process + platform adapter modules.
4. Keep CLI UX and exit-code behavior stable.
5. Add unit tests per command module with mocked adapters.
6. Update launcher docs for module structure and extension guidance.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Launcher commands are implemented as focused modules
- [ ] #2 Process side effects are isolated behind adapter interfaces
- [ ] #3 Existing CLI behavior and exit codes remain compatible
- [ ] #4 Launcher testability improves via mocked adapter tests
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Launcher tests and core test gate pass
- [ ] #2 No regression in wrapper command behavior
<!-- DOD:END -->

