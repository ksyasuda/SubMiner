---
id: TASK-109
title: Add Discord Rich Presence integration with polished activity card
status: In Progress
assignee: []
created_date: '2026-02-22 19:40'
updated_date: '2026-02-22 22:36'
labels:
  - feature
  - discord
  - presence
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add optional Discord Rich Presence support so SubMiner can publish current activity (show/file context, study/mining status, elapsed session) with a polished, readable Discord activity card.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Choose and add a Discord RPC client strategy compatible with Electron + launcher runtime.
2. Add config toggle and safe defaults so Discord presence is opt-in.
3. Define presence payload model (state/details/timestamps/assets/buttons) from existing runtime session metadata.
4. Wire lifecycle updates for start, pause/resume, episode/file changes, and stop/quit.
5. Design assets/text for a clean, branded activity box that is informative without noisy updates.
6. Add debounce/rate-limit handling and fallback behavior when Discord is unavailable.
7. Add focused tests for payload mapping and lifecycle transitions; document setup in user docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Discord Rich Presence can be enabled via config and remains disabled by default.
- [ ] #2 Activity card shows clear state/details and updates correctly across playback/session transitions.
- [ ] #3 Activity card visuals (assets/text) are polished and consistent with project branding.
- [ ] #4 Runtime handles Discord closed/not installed/disconnected without crashes or noisy logs.
- [ ] #5 Docs include setup steps (app/client id), config keys, and troubleshooting notes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented opt-in `discordPresence` config surface in types/defaults/resolver/option-registry/template sections with clamped timing fields and focused config tests.

Added `src/core/services/discord-presence.ts` with payload mapping, debounce+interval update throttling, duplicate suppression, and resilient login/set/clear error handling.

Wired MPV/runtime lifecycle hooks to refresh Discord presence on subtitle/media/path/time/pause/connection transitions and added cleanup stop hook in app lifecycle cleanup path + focused tests.

Updated docs and generated config examples with Discord Rich Presence setup/config/troubleshooting guidance.

Validation status: focused config/runtime/discord tests pass and docs build passes. Full `bun run build` currently blocked by pre-existing `src/main.ts` duplicate imports/symbol errors unrelated to TASK-109 scope (existing in working tree).
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Focused tests cover presence payload mapping and lifecycle update behavior.
- [ ] #2 Manual validation confirms Discord card appearance/updates for at least one real playback session.
- [ ] #3 Build/test/docs gates pass with no regressions.
<!-- DOD:END -->
