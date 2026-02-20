---
id: TASK-38
title: Add user-friendly config validation errors on startup
status: Done
assignee:
  - codex-main
created_date: '2026-02-14 02:02'
updated_date: '2026-02-19 23:18'
labels:
  - config
  - developer-experience
  - error-handling
dependencies: []
priority: medium
ordinal: 66000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve config validation to surface clear, actionable error messages at startup when the user's config file has invalid values, missing required fields, or type mismatches.

## Motivation

The project has a config schema with validation, but invalid configs (wrong types, unknown keys after an upgrade, deprecated fields) can cause cryptic failures deep in service initialization rather than being caught and reported clearly at launch.

## Scope

1. Validate the full config against the schema at startup, before any services initialize
2. Collect all validation errors (don't fail on the first one) and present them as a summary
3. Show the specific field path, expected type, and actual value for each error
4. For deprecated or renamed fields, suggest the correct field name
5. Optionally show validation errors in a startup dialog (Electron dialog) rather than just console output
6. Allow the app to start with defaults for non-critical invalid fields, warning the user about the fallback

## Design constraints

- Must not block startup for non-critical warnings (e.g., unknown extra keys)
- Critical errors (e.g., invalid Anki field mappings) should prevent startup with a clear message
- Config file location should be shown in error output so users know what to edit
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All config fields are validated against the schema before service initialization.
- [x] #2 Validation errors show field path, expected type, and actual value.
- [x] #3 Multiple errors are collected and shown together, not one at a time.
- [x] #4 Deprecated or renamed fields produce a helpful migration suggestion.
- [x] #5 Non-critical validation issues allow startup with defaults and a visible warning.
- [x] #6 Critical validation failures prevent startup with a clear dialog or console message.
- [x] #7 Config file path is shown in error output.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Extend config warning coverage in `src/config/service.ts` (+ tests in `src/config/config.test.ts`) to include unknown top-level keys and deprecated `ankiConnect.openRouter` migration suggestions.
2) Add startup critical validation in `src/core/services/startup.ts` (+ tests in `src/core/services/app-ready.test.ts`) so invalid Anki field mappings are treated as startup-blocking and aggregated.
3) Wire end-to-end startup UX in `src/main.ts`: strict startup reload, fatal parse/config dialogs with config path, aggregated warning summary for non-critical issues, and visible warning notification.
4) Verify with targeted test commands and update TASK-38 notes/acceptance criteria in Backlog.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented startup config UX hardening across config resolver + app-ready lifecycle + Electron wiring.

Added non-fatal validation warnings for unknown top-level keys and deprecated `ankiConnect.openRouter` with migration suggestion to `ankiConnect.ai`.

Added startup critical validation for Anki field mappings when `ankiConnect.enabled === true`; app-ready now aborts before service initialization and reports aggregated critical errors.

Startup config reload now uses `reloadConfigStrict()`; parse failures show a clear `dialog.showErrorBox`, include config path, set non-zero exit code, and quit.

Added aggregated startup warning summary (path + expected/message + actual + fallback) with visible desktop notification while continuing startup with defaults.

Verification: `bun run build && node --test dist/config/config.test.js dist/core/services/app-ready.test.js` (pass: 39/39).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented startup config validation UX end-to-end: strict startup parse validation, aggregated warning/error reporting, migration guidance for deprecated config keys, and startup-blocking critical checks for invalid Anki field mappings. Added targeted test coverage for config warnings and app-ready critical validation flow; verification passed with `bun run build && node --test dist/config/config.test.js dist/core/services/app-ready.test.js`.
<!-- SECTION:FINAL_SUMMARY:END -->
