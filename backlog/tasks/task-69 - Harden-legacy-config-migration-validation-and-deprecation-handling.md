---
id: TASK-69
title: Harden legacy config migration validation and deprecation handling
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - config
  - validation
  - safety
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/config/service.ts` still applies many legacy `ankiConnect` keys via unchecked casts (`as string|number|boolean`) in the `mapLegacy` block. Invalid legacy values can enter resolved runtime config without warnings and break downstream behavior.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Inventory every legacy key currently handled in `mapLegacy`.
2. Replace unchecked assignment with typed validators (`asString`, `asNumber`, `asBoolean`, enum guards).
3. Emit deprecation warnings for accepted legacy keys and validation warnings for invalid values.
4. Preserve backward compatibility for valid legacy configs.
5. Add regression tests for invalid legacy value types and accepted legacy values.
6. Update configuration docs with migration/deprecation notes.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 No legacy key path writes directly to resolved config without type validation
- [ ] #2 Invalid legacy values produce warnings and safe fallback behavior
- [ ] #3 Valid legacy configs still map to equivalent resolved config
- [ ] #4 Config test suite includes legacy-invalid regression coverage
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `bun run test:config:dist` passes
- [ ] #2 Any doc changes for migration/deprecation are committed
<!-- DOD:END -->

