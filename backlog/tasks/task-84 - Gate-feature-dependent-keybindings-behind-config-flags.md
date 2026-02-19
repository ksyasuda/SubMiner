---
id: TASK-84
title: Gate feature-dependent keybindings behind config flags
status: To Do
assignee: []
created_date: '2026-02-19 08:41'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Ensure feature-specific keybindings only work when their related feature is enabled in configuration, so users do not trigger unavailable behavior and the app avoids loading integrations that are disabled (for example Jellyfin).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Feature-dependent keybindings are effectively disabled when their corresponding feature flag/config is off, with no user-facing error dialogs.
- [ ] #2 When a feature is disabled in config, its related integration code is not loaded or initialized during startup (including Jellyfin as a concrete case).
- [ ] #3 When a feature is enabled, existing keybinding behavior continues to work as expected.
- [ ] #4 Automated tests cover enabled and disabled paths for keybinding gating and disabled-integration loading behavior.
- [ ] #5 User-facing docs/config guidance explain that feature keybindings require the corresponding feature to be enabled.
<!-- AC:END -->
