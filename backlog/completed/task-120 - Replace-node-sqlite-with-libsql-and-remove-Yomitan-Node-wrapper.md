---
id: TASK-120
title: 'Replace node:sqlite with libsql and remove Yomitan Node wrapper'
status: Done
assignee: []
created_date: '2026-03-08 04:14'
updated_date: '2026-03-16 05:13'
labels:
  - runtime
  - bun
  - sqlite
  - tech-debt
dependencies: []
priority: medium
ordinal: 65500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove the remaining root Node requirement caused by immersion tracking SQLite usage and the old Yomitan build wrapper by migrating the local SQLite layer off node:sqlite, running the SQLite-backed verification lanes under Bun, and switching the vendored Yomitan build flow to Bun-native scripts.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Immersion tracker runtime no longer imports or requires node:sqlite
- [x] #2 SQLite-backed immersion tracker tests run under Bun without Node --experimental-sqlite
- [x] #3 Root build/test scripts no longer require the Yomitan Node wrapper or Node-based SQLite verification lanes
- [x] #4 README requirements/testing docs reflect the Bun-native workflow
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the immersion tracker SQLite dependency with a local libsql-backed wrapper, updated Bun/runtime compatibility tests to avoid process.exitCode side effects, switched Yomitan builds to run directly inside the vendored Bun-native project, deleted scripts/build-yomitan.mjs, and verified typecheck plus Bun build/test lanes (`build:yomitan`, `test:immersion:sqlite`, `test:runtime:compat`, `test:fast`).
<!-- SECTION:FINAL_SUMMARY:END -->
