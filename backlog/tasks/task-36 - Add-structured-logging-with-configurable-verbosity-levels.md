---
id: TASK-36
title: Add structured logging with configurable verbosity levels
status: To Do
assignee: []
created_date: '2026-02-14 00:59'
labels:
  - infrastructure
  - developer-experience
  - observability
dependencies: []
priority: medium
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
3. Add a CLI flag `--verbose` / `--debug` to override
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
- [ ] #1 A logger module exists with debug/info/warn/error levels.
- [ ] #2 Config option controls default verbosity level.
- [ ] #3 CLI --verbose/--debug flag overrides config.
- [ ] #4 Existing console.log/error calls in core services are migrated to structured logger.
- [ ] #5 MPV socket connection logs use debug level (resolves TASK-33 implicitly).
- [ ] #6 Log output includes source context (service/module name).
- [ ] #7 No performance regression on hot paths (rendering, tokenization).
<!-- AC:END -->
