---
id: TASK-10
title: Consolidate service naming conventions and barrel exports
status: To Do
assignee: []
created_date: '2026-02-11 08:21'
labels:
  - refactor
  - services
  - naming
milestone: Codebase Clarity & Composability
dependencies: []
references:
  - src/core/services/index.ts
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The service layer has inconsistent naming:
- Some functions end in `Service`: `handleCliCommandService`, `loadSubtitlePositionService`
- Some end in `RuntimeService`: `replayCurrentSubtitleRuntimeService`, `sendMpvCommandRuntimeService`
- Some are plain: `shortcutMatchesInputForLocalFallback`
- Factory functions mix `create*DepsRuntimeService` with `create*Service`

The barrel export (src/core/services/index.ts) re-exports 79 symbols from 28 files through a single surface, which obscures dependency boundaries. Consumers import everything from `./core/services` and can't tell which service file they actually depend on.

Establish consistent naming:
- Exported service functions: `verbNounService` (e.g., `handleCliCommand`)  
- Deps factory functions: `create*Deps`
- Consider whether the barrel re-export is still the right pattern vs direct imports from individual files.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 All service functions follow a consistent naming convention
- [ ] #2 Decision documented on barrel export vs direct imports
- [ ] #3 No functional changes
<!-- AC:END -->
