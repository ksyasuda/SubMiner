---
id: TASK-69
title: Refactor mpv plugin into modular script components
status: To Do
assignee: []
created_date: '2026-02-24 17:09'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Break plugin/subminer.lua into smaller Lua modules under mpv scripts subdirectory while preserving user-visible behavior and keybindings. Include migration + docs updates for install paths and smoke/regression checks.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Plugin entrypoint stays at scripts/subminer.lua and loads modules from scripts/subminer/
- [ ] #2 No behavior change for keybindings, script messages, auto-start, AniSkip, hover highlight
- [ ] #3 Install/docs updated for recursive plugin copy
- [ ] #4 Add or update regression checks for start/stop + module loading
<!-- AC:END -->
