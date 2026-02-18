---
id: TASK-30
title: Enable anime streaming via extension repos with configurable mpv playback flow
status: To Do
assignee: []
created_date: '2026-02-13 18:32'
updated_date: '2026-02-13 18:34'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add a feature to query configured extension repositories for anime titles/episodes from SubMiner, let users select a streamable source, and play it through mpv with minimal friction. The result should be interactive from Electron (and triggered from mpv via existing command bridge), and fully configurable in app config.

The implementation should provide a modular backend resolver and a clear UI flow that mirrors existing modal interaction patterns, while keeping mpv playback unchanged (use loadfile with resolved URL and optional headers/referrer metadata).

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Create a stable config schema for one or more extension source backends (repos/endpoints/flags) and persist in user config + default template.
- [ ] #2 Resolve an anime search term to candidate series from configured sources.
- [ ] #3 Resolve an episode selection to at least one playable stream URL candidate with playback metadata when available.
- [ ] #4 Provide an interactive user flow in the app to search/select episode and choose stream source.
- [ ] #5 Play the selected stream by invoking the existing mpv command path.
- [ ] #6 Support launching the flow from mpv interactions (keyboard/menu) by forwarding a command/event into the renderer modal flow.
- [ ] #7 Add error states for empty results, no playable source, and repository failures, with clear user messages.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Execution sequence for implementation: 1) TASK-30.1 (config/model), 2) TASK-30.2 (resolver), 3) TASK-30.3 (IPC), 4) TASK-30.4 (UI modal), 5) TASK-30.5 (mpv trigger), 6) TASK-30.6 (validation/rollout checklist).

Rollout recommendation: complete TASK-30.6 only after TASK-30.1-30.5 are done and can be verified in combination.

<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 Config schema validated in app startup (bad config surfaces clear error).
- [ ] #2 No hardcoded source/resolver URLs in UI layer; resolver details are backend-driven.
- [ ] #3 Play command path uses existing mpv IPC/runtime helpers.
- [ ] #4 Documentation includes how to configure extension repos and optional auth/headers.
- [ ] #5 Feature is gated behind a config flag and can be disabled cleanly.
<!-- DOD:END -->
