---
id: TASK-69
title: Harden legacy config migration validation and deprecation handling
status: Done
assignee:
  - codex-main
created_date: '2026-02-18 11:35'
updated_date: '2026-02-19 08:27'
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
- [x] #1 No legacy key path writes directly to resolved config without type validation
- [x] #2 Invalid legacy values produce warnings and safe fallback behavior
- [x] #3 Valid legacy configs still map to equivalent resolved config
- [x] #4 Config test suite includes legacy-invalid regression coverage
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Implementation plan saved at `docs/plans/2026-02-19-task-69-legacy-config-migration-hardening.md`.

Execution breakdown:
1. Add failing regression tests for invalid legacy `ankiConnect` migration values and valid-legacy compatibility behavior.
2. Replace unchecked `mapLegacy` casts in `src/config/service.ts` with typed validators using `asString`/`asBoolean`/`asNumber` plus enum/range checks; warn on invalid legacy input and keep deprecation warnings.
3. Preserve modern key precedence and existing `ankiConnect.openRouter` deprecation behavior.
4. Run `bun run build && bun run test:config:dist` to verify.
5. Update `docs/configuration.md` only if migration/deprecation behavior clarification is needed.
6. Append notes and acceptance progress back into TASK-69.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented legacy migration hardening in `src/config/service.ts` by filtering top-level legacy keys out of broad `ankiConnect` merge and replacing unchecked `mapLegacy` casts with parser-backed legacy mapping (`asString`/`asBoolean`/range+enum validators). Invalid legacy values now emit warnings and keep safe fallback values; valid legacy values still map when modern key is absent.

Added regression tests in `src/config/config.test.ts`: `falls back and warns when legacy ankiConnect migration values are invalid` and `maps valid legacy ankiConnect values to equivalent modern config`.

Updated docs in `docs/configuration.md` to clarify that legacy top-level `ankiConnect` migration keys are compatibility-only, validated, and ignored with warnings when invalid.

Verification: `bun run build && bun run test:config:dist` passed (32/32).
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `bun run test:config:dist` passes
- [x] #2 Any doc changes for migration/deprecation are committed
<!-- DOD:END -->
