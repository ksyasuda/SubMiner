---
id: TASK-353
title: Remove Makefile global mpv plugin installer
status: Done
assignee: []
created_date: '2026-05-12 14:07'
updated_date: '2026-05-12 14:07'
labels:
  - launcher
  - mpv-plugin
  - docs
dependencies: []
references:
  - app-managed-mpv-runtime-plugin-plan.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove the legacy Makefile path that installs SubMiner into mpv's global scripts directory, including the Windows config rewrite script hook, because managed playback now injects the bundled runtime plugin.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Makefile no longer exposes an `install-plugin` target or help entry.
- [x] #2 Windows install no longer delegates to global mpv plugin installation.
- [x] #3 The obsolete config rewrite bun script is removed when no longer referenced.
- [x] #4 Docs no longer tell users to run `make install-plugin`.
- [x] #5 Regression coverage prevents reintroducing the target or script hook.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the legacy global mpv plugin install target from the Makefile, removed the obsolete Windows config rewrite script, updated docs to describe bundled runtime plugin loading instead of separate installation, and removed the setup window's legacy install action while keeping legacy plugin removal available.
<!-- SECTION:FINAL_SUMMARY:END -->
