---
id: TASK-8
title: >-
  Reduce MpvIpcClient deps interface and separate protocol from application
  logic
status: To Do
assignee: []
created_date: '2026-02-11 08:20'
labels:
  - refactor
  - mpv
  - architecture
milestone: Codebase Clarity & Composability
dependencies: []
references:
  - src/core/services/mpv-service.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
MpvIpcClient (761 lines) in src/core/services/mpv-service.ts has a 22-property `MpvIpcClientDeps` interface that reaches back into main.ts state for application-level concerns (overlay visibility, subtitle timing, media path updates, OSD display).

The class mixes two responsibilities:
1. **IPC Protocol**: Socket connection, JSON message framing, reconnection, property observation
2. **Application Integration**: Subtitle text broadcasting, overlay visibility sync, timing tracking

Separating these would let the protocol layer be simpler and testable, while application-level reactions to mpv events could be handled by listeners/callbacks registered externally.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 MpvIpcClient deps interface reduced to protocol-level concerns only
- [ ] #2 Application-level reactions (subtitle broadcast, overlay sync, timing) handled via event emitter or external listeners
- [ ] #3 MpvIpcClient is testable without mocking 22 callbacks
- [ ] #4 Existing behavior preserved
<!-- AC:END -->
