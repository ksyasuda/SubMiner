---
id: TASK-36
title: Add structured logging with configurable verbosity levels
status: Done
assignee: []
created_date: '2026-02-14 00:59'
updated_date: '2026-02-18 04:11'
labels:
  - infrastructure
  - developer-experience
  - observability
dependencies: []
priority: medium
ordinal: 20000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace ad-hoc console.log/console.error calls throughout the codebase with a lightweight structured logger that supports configurable verbosity levels (debug, info, warn, error).

## Motivation
- TASK-33 (restrict mpv socket logs) is a symptom of a broader problem: no log-level filtering
- Debugging production issues requires grepping through noisy output
- Users report log spam in normal operation
- No way to enable verbose logging for bug reports without code changes

## Scope
1. Create a minimal logger module (no external dependencies needed) with `debug`, `info`, `warn`, `error` levels
2. Add a config option for log verbosity (default: `info`)
3. Add a CLI flag to control logging verbosity (`--log-level`) while keeping `--debug` as app/dev mode.
4. Migrate existing console.log/error calls to use the logger
5. Include context tags (service name, operation) in log output for filterability

## Design constraints
- Zero external dependencies — use a simple wrapper over console methods
- Must not impact hot-path performance (subtitle rendering, tokenization)
- Log level should be changeable at runtime via config hot-reload if that feature exists
- TASK-33 becomes trivial once this lands (mpv socket logs become `debug` level)
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A logger module exists with debug/info/warn/error levels.
- [x] #2 Config option controls default verbosity level.
- [x] #3 CLI `--log-level` override config.
- [x] #4 Existing console.log/error calls in core services are migrated to structured logger.
- [x] #5 MPV socket connection logs use debug level (resolves TASK-33 implicitly).
- [x] #6 Log output includes source context (service/module name).
- [x] #7 No performance regression on hot paths (rendering, tokenization).
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Audit remaining runtime console calls and classify by target (core runtime vs help/UI/test-only).
2) Keep `--debug` scoped to Electron app/dev mode only; `--log-level` controls logging verbosity.
3) Add tests for parsing and startup to keep logging override behavior stable.
4) Migrate remaining non-user-facing `console.*` calls in core paths (especially tokenization, jimaku, config generation, electron-backend/notifications) to logger and include context via child loggers.
5) Ensure mpv-related connection/reconnect messages use debug level; keep user-facing success/failure outputs when intended.
6) Run focused test updates for impacted files, update task notes/acceptance criteria, and finalize task state.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated logging override semantics so `--debug` stays an app/dev-only flag and `--log-level` is the CLI logging control.

Migrated remaining non-user-facing runtime `console.*` calls in core paths (`notification`, `config-gen`, `electron-backend`, `jimaku`, `tokenizer`) to structured logger

Moved MPV socket lifecycle chatter to debug-level in `mpv-service` so default info is less noisy

Updated help text and CLI parsing/tests for logging via `--log-level` while keeping `--debug` app/dev-only.

Validated with focused test run: `node --test dist/cli/args.test.js dist/core/services/startup-bootstrap-service.test.js dist/core/services/cli-command-service.test.js`

Build still passes via `pnpm run build`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
TASK-36 is now complete; structured logging is consistently used in core runtime paths, with CLI log verbosity controlled by `--log-level`, while `--debug` remains an Electron app/dev-mode toggle.

Backlog task moved to Done after verification of build and focused tests.
<!-- SECTION:FINAL_SUMMARY:END -->
