---
id: TASK-105
title: Stop local docs artifact writes after docs repo split
status: Done
assignee: []
created_date: '2026-03-07 00:00'
updated_date: '2026-03-07 00:20'
labels: []
dependencies: []
priority: medium
ordinal: 10500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Now that user-facing docs live in `../subminer-docs`, first-party scripts in this repo should not keep writing generated artifacts into the local `docs/` tree.

Scope:

- Audit first-party scripts/automation for writes to `docs/`.
- Keep repo-local outputs only where they are still intentionally owned by this repo.
- Repoint generated docs artifacts to `../subminer-docs` when that is the maintained source of truth.
- Add regression coverage for the config-example generation path contract.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 The config-example generator no longer writes to `docs/public/config.example.jsonc` inside this repo.
- [x] #2 When `../subminer-docs` exists, the generator updates `../subminer-docs/public/config.example.jsonc`.
- [x] #3 Automated coverage guards the output-path contract so local docs writes do not regress.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Removed the first-party local `docs/public` config-example write path from `src/generate-config-example.ts` and replaced it with sibling-docs-repo detection that targets `../subminer-docs/public/config.example.jsonc` only when that repo exists.

Added a project-local regression suite for output-path resolution and artifact writing, wired that suite into the maintained config test lane, and removed the stale generated `docs/public/config.example.jsonc` artifact from the working tree.

<!-- SECTION:FINAL_SUMMARY:END -->
