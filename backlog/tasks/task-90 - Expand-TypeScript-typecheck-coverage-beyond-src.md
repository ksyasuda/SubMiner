---
id: TASK-90
title: Expand TypeScript typecheck coverage beyond src
status: Done
assignee: []
created_date: '2026-03-06 08:18'
updated_date: '2026-03-06 08:23'
labels:
  - tooling
  - typescript
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/tsconfig.json
  - /home/sudacode/projects/japanese/SubMiner/package.json
  - /home/sudacode/projects/japanese/SubMiner/launcher
  - /home/sudacode/projects/japanese/SubMiner/scripts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Bring all repository TypeScript entrypoints outside src/ into the enforced typecheck gate so CI and local checks cover launcher/ and script files, then resolve any surfaced diagnostics.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 TypeScript typecheck covers repository TypeScript entrypoints outside src/ that should be maintained in this repo, including launcher/ and script files.
- [x] #2 The enforced typecheck command used by CI and local development passes with the expanded coverage.
- [x] #3 Any diagnostics surfaced by the expanded coverage are fixed without weakening existing strictness for src/.
- [x] #4 Relevant documentation or command wiring is updated if the typecheck entrypoint changes.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a dedicated repo-wide typecheck config at tsconfig.typecheck.json and wired package.json/CI to use `bun run typecheck` for launcher and scripts coverage without changing the existing src build config. Fixed the strict-null/indexing diagnostics surfaced in launcher/* and scripts/*, keeping src strictness intact. Verified with `bun run typecheck`, `bun run tsc --noEmit`, and `bun run test:launcher:src` (47 passing, plugin start gate OK).
<!-- SECTION:FINAL_SUMMARY:END -->
