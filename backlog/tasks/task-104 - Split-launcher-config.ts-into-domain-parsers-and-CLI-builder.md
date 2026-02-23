---
id: TASK-104
title: Split launcher config.ts into domain parsers and CLI builder
status: Done
assignee:
  - codex
created_date: '2026-02-22 07:13'
updated_date: '2026-02-23 02:06'
labels:
  - refactor
  - launcher
  - maintainability
dependencies:
  - TASK-81
  - TASK-102
priority: medium
ordinal: 103000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`launcher/config.ts` is still a large multi-responsibility file (~700 LOC) combining:
- config file reading/parsing for multiple domains,
- plugin runtime config parsing,
- CLI command tree construction,
- root/subcommand arg normalization.

This file remains a cleanup hotspot and makes contract changes (like Jellyfin session migration) expensive to land safely.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract launcher config-file readers into domain loaders (YouTube/Jimaku, Jellyfin, plugin runtime).
2. Extract Commander command-tree setup into a dedicated CLI builder module.
3. Extract post-parse normalization into focused argument-normalization helpers.
4. Remove stale Jellyfin config auth field assumptions from launcher config readers.
5. Add focused tests per extracted module while preserving existing `launcher/config.test.ts` behavior expectations.
6. Keep `parseArgs` API stable for launcher call sites.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `launcher/config.ts` is reduced to thin orchestration over extracted modules.
- [x] #2 Each extracted module has focused tests that assert current behavior.
- [x] #3 Launcher still passes `bun run test:launcher` without CLI behavior regressions.
- [x] #4 Launcher config readers align with current Jellyfin session contract (no config token/userId dependency).
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan recorded at `docs/plans/2026-02-22-task-104-launcher-config-domain-parsers-cli-builder.md`.

Execution phases:
1. Extract shared config readers and domain parser modules (`launcher/config/shared-config-reader.ts`, `launcher/config/youtube-subgen-config.ts`, `launcher/config/jellyfin-config.ts`) and reduce `launcher/config.ts` to orchestration.
2. Extract plugin runtime parser into `launcher/config/plugin-runtime-config.ts` with focused behavior tests.
3. Extract CLI parser builder/normalization into `launcher/config/cli-parser-builder.ts`, `launcher/config/parse-helpers.ts`, and `launcher/config/args-normalizer.ts`; keep `parseArgs` API unchanged.
4. Align Jellyfin config contract by removing launcher config token/userId dependence and update caller/type wiring.
5. Verify with `bun run test:launcher` and `bun run test:fast`, then finalize TASK-104 evidence (AC/DoD checks + summary).

Validation-first loop per phase: add/expand focused tests, run targeted launcher tests, implement minimal refactor to pass, then run full required suites.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented launcher config decomposition by extracting domain-focused modules under `launcher/config/` (`shared-config-reader`, `youtube-subgen-config`, `jellyfin-config`, `plugin-runtime-config`, `cli-parser-builder`, `args-normalizer`) and reducing `launcher/config.ts` to an orchestration facade with unchanged exported API.

Aligned launcher Jellyfin config contract by removing `accessToken`/`userId` from `LauncherJellyfinConfig` and parser output; launcher Jellyfin play now requires `SUBMINER_JELLYFIN_ACCESS_TOKEN` + `SUBMINER_JELLYFIN_USER_ID` env session values instead of config fields.

Added focused parser regression tests in `launcher/config-domain-parsers.test.ts` (youtube domain normalization, jellyfin legacy token/userId omission, plugin socket parsing) and expanded `launcher/parse-args.test.ts` branch coverage for jellyfin/config/mpv command mappings.

Verification: `bun test launcher/config-domain-parsers.test.ts launcher/parse-args.test.ts`, `bun run test:launcher`, and `bun run test:fast` all pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Split `launcher/config.ts` into focused domain parser and CLI parsing modules while preserving the public launcher parsing API (`parseArgs`, config readers, plugin runtime reader). Added focused launcher parser tests, expanded parse-args coverage, and removed launcher config dependency on Jellyfin token/userId fields to match the current session contract. Verified behavior with `bun run test:launcher` and `bun run test:fast` passing.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Public launcher parsing API unchanged for downstream callers.
- [x] #2 Help text and subcommand option behavior remains unchanged.
- [x] #3 `bun run test:launcher` and `bun run test:fast` pass.
<!-- DOD:END -->
