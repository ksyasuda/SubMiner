---
id: TASK-70
title: Unify config path resolution across main and launcher
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - config
  - launcher
  - consistency
dependencies: []
priority: high
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
- [ ] #1 Single canonical path-resolution logic used by app and launcher
- [ ] #2 `XDG_CONFIG_HOME` and `SubMiner|subminer` precedence covered by tests
- [ ] #3 No behavior drift for existing config-path CLI commands
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Launcher and config tests pass
- [ ] #2 Code no longer duplicates config path candidate logic
<!-- DOD:END -->

