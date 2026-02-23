---
id: TASK-111
title: Remove Maintainability Guardrails docs section and related guardrail code
status: Done
assignee: []
created_date: '2026-02-23 03:37'
updated_date: '2026-02-23 03:40'
labels:
  - docs
  - cleanup
  - guardrails
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User requested removing the docs section labeled "Maintainability Guardrails" and deleting associated project code/commands for those guardrails so docs and scripts stay consistent.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Docs no longer contain the Maintainability Guardrails section shown in the request.
- [x] #2 Commands and code associated specifically with that removed guardrails section are removed from scripts/config where applicable.
- [x] #3 Project references remain consistent (no stale mentions of removed guardrails commands).
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed the `## Maintainability Guardrails` section from `docs/development.md` including strict command examples and troubleshooting bullets.

Removed guardrail scripts from `package.json`: `check:main-fanin`, `check:main-fanin:strict`, `check:runtime-cycles`, `check:runtime-cycles:strict`.

Deleted associated script code and fixtures: `scripts/check-main-runtime-fanin.ts`, `scripts/check-runtime-cycles.ts`, `scripts/check-runtime-cycles.test.ts`, and `scripts/fixtures/runtime-cycles/**`.

Removed CI fail-fast guardrail step from `.github/workflows/ci.yml` that invoked strict fan-in/runtime-cycle checks.

Validation: `bun run tsc --noEmit` and `bun run docs:build` passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the Maintainability Guardrails docs section and fully removed the related fan-in/runtime-cycle guardrail implementation from scripts, package commands, CI wiring, and fixtures. The repository now has no active references to those guardrail commands in development docs, package scripts, or CI workflow.
<!-- SECTION:FINAL_SUMMARY:END -->
