---
id: TASK-84
title: Gate feature-dependent keybindings behind config flags
status: Done
assignee:
  - opencode-task84-keybindings-gating
created_date: '2026-02-19 08:41'
updated_date: '2026-02-22 01:35'
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
- [x] #1 Feature-dependent keybindings are effectively disabled when their corresponding feature flag/config is off, with no user-facing error dialogs.
- [x] #2 When a feature is disabled in config, its related integration code is not loaded or initialized during startup (including Jellyfin as a concrete case).
- [x] #3 When a feature is enabled, existing keybinding behavior continues to work as expected.
- [x] #4 Automated tests cover enabled and disabled paths for keybinding gating and disabled-integration loading behavior.
- [x] #5 User-facing docs/config guidance explain that feature keybindings require the corresponding feature to be enabled.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Add feature-gated shortcut resolution in `src/core/utils/shortcut-config.ts` so Anki-dependent shortcuts resolve to `null` when `ankiConnect.enabled` is false, with focused tests in `src/core/utils/shortcut-config.test.ts`.
2) Gate Jellyfin startup integration in `src/main/runtime/jellyfin-remote-session-lifecycle.ts` and startup warmup wiring in `src/main.ts` so remote session initialization requires `jellyfin.enabled`, `jellyfin.remoteControlEnabled`, and `jellyfin.remoteControlAutoConnect`.
3) Expand automated coverage for enabled/disabled paths (shortcut gating + Jellyfin startup no-op), then include new shortcut gating tests in `test:core:src`.
4) Update `docs/configuration.md` to clarify feature-dependent shortcuts require enabled features and Jellyfin remote autoconnect prerequisites.
5) Run verification (`bun test` focused suites, `bun run test:core:src`, `bun run docs:build`) and finalize TASK-84 notes/AC/final summary (no commit).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Plan captured in docs/plans/2026-02-22-task-84-feature-keybinding-gates.md. Proceeding immediately with execution via executing-plans flow per user request.

Implemented feature-dependent shortcut gating in `src/core/utils/shortcut-config.ts`: when `ankiConnect.enabled` is false, Anki/Kiku-dependent shortcuts resolve to `null` (`updateLastCardFromClipboard`, `triggerFieldGrouping`, `mineSentence`, `mineSentenceMultiple`, `markAudioCard`).

Added focused shortcut gating coverage in `src/core/utils/shortcut-config.test.ts` (disabled + enabled + normalization fallback), and wired this file into `test:core:src` in `package.json`.

Added disabled-integration startup guards for Jellyfin remote initialization: `createStartJellyfinRemoteSessionHandler` now early-returns when `jellyfin.enabled` is false, and startup warmup predicate in `src/main.ts` now requires `enabled && remoteControlEnabled && remoteControlAutoConnect`.

Extended Jellyfin startup tests in `src/main/runtime/jellyfin-remote-session-lifecycle.test.ts` and `src/main/runtime/startup-warmups.test.ts` to cover disabled and fully-enabled paths.

Updated docs in `docs/configuration.md` to clarify feature-dependent shortcuts require enabled integrations and Jellyfin remote autoconnect prerequisites.

Verification passed: `bun test src/core/utils/shortcut-config.test.ts src/main/runtime/jellyfin-remote-session-lifecycle.test.ts src/main/runtime/startup-warmups.test.ts src/core/services/overlay-shortcut-handler.test.ts`, `bun run test:core:src`, and `bun run docs:build`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented config-driven feature gates for shortcuts and integration startup so feature-dependent behavior is disabled cleanly when the feature is off. Anki/Kiku-dependent overlay shortcuts now resolve to null when `ankiConnect.enabled` is false, Jellyfin remote startup now requires `jellyfin.enabled` in addition to remote-control flags, focused automated tests cover enabled/disabled paths, and configuration docs now explicitly describe these feature prerequisites.
<!-- SECTION:FINAL_SUMMARY:END -->
