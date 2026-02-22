---
id: TASK-104
title: Split launcher config.ts into domain parsers and CLI builder
status: To Do
assignee: []
created_date: '2026-02-22 07:13'
updated_date: '2026-02-22 07:13'
labels:
  - refactor
  - launcher
  - maintainability
dependencies:
  - TASK-81
  - TASK-102
priority: medium
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
- [ ] #1 `launcher/config.ts` is reduced to thin orchestration over extracted modules.
- [ ] #2 Each extracted module has focused tests that assert current behavior.
- [ ] #3 Launcher still passes `bun run test:launcher` without CLI behavior regressions.
- [ ] #4 Launcher config readers align with current Jellyfin session contract (no config token/userId dependency).
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Public launcher parsing API unchanged for downstream callers.
- [ ] #2 Help text and subcommand option behavior remains unchanged.
- [ ] #3 `bun run test:launcher` and `bun run test:fast` pass.
<!-- DOD:END -->

