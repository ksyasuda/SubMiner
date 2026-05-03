---
id: TASK-314
title: Improve Jellyfin setup popup and tray discovery toggle
status: Done
assignee: []
created_date: '2026-05-02 22:45'
updated_date: '2026-05-02 23:11'
labels:
  - jellyfin
dependencies: []
references:
  - src/main/runtime/jellyfin-setup-window.ts
  - src/main/runtime/jellyfin-cli-auth.ts
  - src/main/runtime/tray-runtime.ts
  - src/main/runtime/jellyfin-remote-session-lifecycle.ts
documentation:
  - docs-site/jellyfin-integration.md
  - docs-site/configuration.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve the Jellyfin integration setup experience and remove the need to use command-line discovery mode for normal tray-driven use. The existing `--jellyfin` setup popup should become a frontend for the same auth persistence path used by CLI login, with manual/recent server selection and inline feedback. The tray should expose a runtime-only Jellyfin Discovery checkbox when Jellyfin is configured so users can start or stop cast/discovery mode without changing config.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The Jellyfin setup popup supports config/recent/default server choices, manual URL entry, username/password login, logout when a session exists, done/close, and inline success/error status without persisting passwords.
- [x] #2 CLI login and setup popup login share the same auth persistence behavior, including encrypted token storage, enabled/server/username/client metadata config patching, and recent server updates.
- [x] #3 `jellyfin.recentServers` is parsed, normalized, deduplicated, capped, documented, and included in generated config examples if exposed.
- [x] #4 The tray keeps Configure Jellyfin visible and shows a Jellyfin Discovery checkbox only when Jellyfin is configured with enabled integration, server URL, access token, and user ID.
- [x] #5 The tray Jellyfin Discovery checkbox starts/stops the current remote session at runtime only, announces after start, reports OSD/log status, and does not patch config.
- [x] #6 Startup auto-connect behavior remains governed by existing config, including `remoteControlAutoConnect`; explicit tray start can start discovery without requiring `remoteControlAutoConnect`.
- [x] #7 Focused tests cover setup popup actions/rendering, shared auth persistence, config parsing, tray toggle visibility/state/click behavior, and remote lifecycle auto-connect versus explicit-start behavior.
- [x] #8 Jellyfin docs and changelog fragment are updated.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented Jellyfin setup popup improvements, shared auth persistence for CLI/setup config shape, recent server config support, runtime-only tray Jellyfin Discovery toggle, docs/config examples, and changelog fragment. Verified focused Jellyfin/tray tests, config tests, launcher tests, typecheck, and docs tests.
<!-- SECTION:FINAL_SUMMARY:END -->
