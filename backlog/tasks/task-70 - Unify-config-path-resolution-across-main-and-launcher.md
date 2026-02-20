---
id: TASK-70
title: Unify config path resolution across main and launcher
status: Done
assignee:
  - codex-main
created_date: '2026-02-18 11:35'
updated_date: '2026-02-19 23:18'
labels:
  - config
  - launcher
  - consistency
dependencies: []
priority: high
ordinal: 63000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Config discovery logic is duplicated and inconsistent between app main process and launcher. Main resolves `XDG_CONFIG_HOME` + case variants (`SubMiner`/`subminer`), while launcher loaders still hardcode `~/.config/SubMiner` in places.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract shared config path discovery utility (candidate roots, case variants, `.jsonc`/`.json` preference).
2. Replace ad-hoc resolution in `src/main.ts`, `launcher/main.ts`, and `launcher/config.ts`.
3. Normalize fallback behavior when config file is absent.
4. Keep existing user-visible behavior for `subminer config path|show`.
5. Add tests for XDG and case-variant path resolution.
6. Update docs if path precedence changes.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Single canonical path-resolution logic used by app and launcher
- [x] #2 `XDG_CONFIG_HOME` and `SubMiner|subminer` precedence covered by tests
- [x] #3 No behavior drift for existing config-path CLI commands
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused precedence tests in `src/config/path-resolution.test.ts` for XDG/home base-dir order, SubMiner/subminer fallbacks, config.jsonc/config.json preference, and fallback path behavior.
2. Create canonical helper module `src/config/path-resolution.ts` exporting shared discovery functions for config dir and config file resolution.
3. Replace duplicated path-resolution logic in `src/main.ts`, `launcher/main.ts`, and launcher config loaders in `launcher/config.ts` to use the canonical helper.
4. Verify no behavior drift with `bun run build`, config/path tests, and launcher bundle build; then update backlog acceptance/DoD checks and execution notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-19T08:52:23Z: Started task execution; gathering code context for unified config-path resolution across main and launcher.

Plan captured in docs/plans/2026-02-19-task-70-unify-config-path-resolution.md and recorded in task. User requested immediate execution after planning.

Implemented canonical config path discovery in `src/config/path-resolution.ts` and switched `src/main.ts`, `launcher/main.ts`, and launcher config loaders in `launcher/config.ts` to use it.

Added precedence regression coverage in `src/config/path-resolution.test.ts` and wired into `test:config:dist`.

Verification passed: `bun run build`, `node --test dist/config/path-resolution.test.js dist/config/config.test.js`, and `bun build ./launcher/main.ts --target=bun --packages=bundle --outfile=/tmp/subminer-task70`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Unified config discovery behind a single canonical utility at `src/config/path-resolution.ts` and replaced duplicated resolvers in `src/main.ts`, `launcher/main.ts`, and launcher config loaders in `launcher/config.ts`. This keeps `XDG_CONFIG_HOME` + `~/.config` base-dir handling, `SubMiner|subminer` case variants, and `config.jsonc` > `config.json` file preference consistent across app and launcher while preserving `subminer config path|show` fallback output behavior.

Added regression tests in `src/config/path-resolution.test.ts` for base-dir trimming/dedup, candidate precedence, lowercase fallback, directory fallback for main config dir resolution, and fallback file-path behavior; wired the new suite into `test:config:dist` in `package.json`.

Verification run: `bun run build`, `node --test dist/config/path-resolution.test.js dist/config/config.test.js`, `bun build ./launcher/main.ts --target=bun --packages=bundle --outfile=/tmp/subminer-task70` (all pass).
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Launcher and config tests pass
- [x] #2 Code no longer duplicates config path candidate logic
<!-- DOD:END -->
