---
id: TASK-113
title: Scope make pretty to maintained source files
status: Done
assignee:
  - codex
created_date: '2026-03-08 00:20'
updated_date: '2026-03-08 00:22'
labels:
  - tooling
  - formatting
dependencies: []
references:
  - Makefile
  - package.json
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Change the `make pretty` workflow so it formats only the maintained source/config files we intentionally keep under Prettier, instead of sweeping backlog/docs/generated content across the whole repository.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 `make pretty` formats only the approved maintained source/config paths
- [x] #2 The allowlist is reusable for check/write flows instead of duplicating path logic
- [x] #3 Verification shows the scoped formatting command targets the intended files without touching backlog or vendored content
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Inspect current Prettier config/ignore behavior and keep the broad repo-wide format command unchanged.
2. Add a reusable scoped Prettier script that targets maintained source/config paths only.
3. Update `make pretty` to call the scoped script.
4. Verify the scoped command resolves only intended files and does not traverse backlog or vendor paths.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

User approved the allowlist approach: keep repo-wide `format` intact, make `make pretty` use a maintained-path formatter scope.

Added `scripts/prettier-scope.sh` as the single allowlist for scoped Prettier paths and wired `format:src` / `format:check:src` to it.

Updated `make pretty` to call `bun run format:src`. Verified with `make -n pretty` and shell tracing that the helper only targets the maintained allowlist and does not traverse `backlog/` or `vendor/`.

Excluded `Makefile` and `.prettierignore` from the allowlist after verification showed Prettier cannot infer parsers for them.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Scoped the repo's day-to-day formatting entrypoint without changing the existing broad repo-wide Prettier scripts. Added `scripts/prettier-scope.sh` as the shared allowlist for maintained source/config paths (`.github`, `build`, `launcher`, `scripts`, `src`, plus selected root JSON config files), added `format:src` and `format:check:src` in `package.json`, and updated `make pretty` to run the scoped formatter.

Verification: `make -n pretty` now resolves to `bun run format:src`. `bash -n scripts/prettier-scope.sh` passed, and shell-traced `bash -x scripts/prettier-scope.sh --check` confirmed the exact allowlist passed to Prettier. `bun run format:check:src` fails only because existing files inside the allowed source scope are not currently formatted; it no longer touches `backlog/` or `vendor/`.

<!-- SECTION:FINAL_SUMMARY:END -->
