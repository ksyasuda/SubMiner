---
id: TASK-72
title: Make startup config loading strict with clear user-facing errors
status: Done
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-22 07:49'
labels:
  - config
  - ux
  - reliability
dependencies: []
priority: medium
ordinal: 82000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Startup config loading currently falls back to default config when strict parse fails, while hot-reload path reports strict errors. This mismatch can hide broken config at startup and cause surprising runtime behavior.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Align startup loading semantics with strict reload behavior.
2. Surface parse/validation errors to user via clear logs/OSD/notification where appropriate.
3. Define fallback policy explicitly (fail fast vs safe mode) and apply consistently.
4. Add regression tests for malformed JSON/JSONC at startup.
5. Document startup error behavior in troubleshooting/config docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Startup and hot-reload use consistent strictness semantics
- [x] #2 Malformed config surfaces explicit actionable error
- [x] #3 Tests cover malformed startup config path
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: implementing plan Task 1 only in this run (strict constructor startup load + malformed startup config test coverage), no commit.

2026-02-21 Task 1 complete: `ConfigService` constructor now strict-loads startup config and throws `ConfigStartupParseError` (path + parse reason + restart guidance) on malformed parse; added malformed-startup constructor test in `src/config/config.test.ts`; verified with `bun run build && node --test dist/config/config.test.js` (pass).

2026-02-21 Task 2-4 complete: added shared `buildConfigParseErrorDetails` formatter and reused it in startup/hot-reload parse-failure handling (`src/main/config-validation.ts`, `src/main/runtime/startup-config.ts`).

2026-02-21 startup wiring updated in `src/main.ts`: `ConfigService` construction now catches `ConfigStartupParseError` and fails startup via `failStartupFromConfig` with user-facing dialog + actionable details, keeping non-parse failures rethrown.

2026-02-21 docs updated in `docs/configuration.md` to clarify malformed JSON/JSONC is startup-blocking while value-level validation remains warn-and-fallback.

Verification: `bun test src/config/config.test.ts src/main/config-validation.test.ts src/main/runtime/startup-config.test.ts` => pass (46/46). Full `bun run build` currently blocked by unrelated pre-existing TS errors in in-flight TASK-75/TASK-77 files (`src/main/runtime/mpv-osd-log.test.ts`, `src/main/runtime/mpv-osd-runtime-handlers.test.ts`, `src/core/services/tokenizer/parser-selection-stage.test.ts`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Startup config loading is now strict and aligned with hot-reload semantics. `ConfigService` no longer silently falls back on malformed startup config; it throws a typed `ConfigStartupParseError`, and `main.ts` now converts that into the same actionable user-facing startup failure path used elsewhere (dialog/error details include config path, parse reason, and restart guidance). Added malformed-startup constructor regression coverage plus shared parse-error formatter tests, and updated configuration docs to distinguish fatal syntax errors from warn-and-fallback value validation.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Config and startup tests pass
- [x] #2 Docs updated for new startup error behavior
<!-- DOD:END -->
