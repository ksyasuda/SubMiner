---
id: TASK-352
title: Inject bundled mpv plugin for managed launches
status: Done
assignee: []
created_date: '2026-05-12 20:06'
updated_date: '2026-05-12 20:15'
labels:
  - launcher
  - mpv-plugin
  - windows
  - setup
dependencies: []
references:
  - app-managed-mpv-runtime-plugin-plan.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement app-managed mpv runtime plugin loading so SubMiner-managed launcher and Windows mpv shortcut launches do not require a globally installed mpv plugin, while installed legacy/global plugins are detected and take precedence until removed.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher-managed mpv launch injects the bundled plugin when no installed SubMiner mpv plugin is detected.
- [x] #2 Launcher idle/Jellyfin mpv launch follows the same bundled-vs-installed plugin policy.
- [x] #3 Windows SubMiner mpv shortcut launch skips bundled injection when an installed plugin is detected and injects bundled plugin otherwise.
- [x] #4 First-run setup no longer requires global mpv plugin installation to finish; plugin install remains optional compatibility action.
- [x] #5 Runtime plugin path resolution is test-covered and reports a clear failure when no bundled plugin path is available and no installed plugin exists.
- [x] #6 Release docs/changelog explain that managed launches no longer require global plugin installation and legacy plugin removal switches to bundled runtime loading.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented app-managed mpv runtime plugin policy. Launcher-managed playback and idle/Jellyfin mpv startup now inject the bundled plugin when no global SubMiner plugin is detected, and keep using/logging the installed plugin when one is present. Windows SubMiner mpv shortcut launches use the same installed-vs-bundled policy while still passing SubMiner script opts. First-run setup no longer requires global plugin installation to finish, keeps legacy install as optional compatibility, and documents/removes legacy global plugin files via OS trash.
<!-- SECTION:FINAL_SUMMARY:END -->
