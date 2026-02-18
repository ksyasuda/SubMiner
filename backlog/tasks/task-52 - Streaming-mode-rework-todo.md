---
id: TASK-52
title: >-
  Consolidate stream-mode/ani-cli work into one deferred streaming integration
  ticket
status: To Do
assignee: []
created_date: '2026-02-14 08:21'
labels:
  - someday
  - streaming
  - ani-cli
  - feature-flag
  - consolidation
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Replace fragmented streaming-mode items with a single future-oriented task.

Context:

- Previous attempts to support `subminer -s` via ani-cli were removed due instability and blockers.
- Stream URLs resolved through ani-cli often hit Cloudflare anti-bot/403 (`generic:impersonate`) in subminer/mpv/ytdl-hook path.
- URL/title handling from stream selection was inconsistent, breaking robust subtitle discovery/matching.
- Subsync logging/debugging required improvements before deeper stream validation.
- A single cohesive redesign is preferable over piecemeal reintroduction.

Scope (one ticket):

- Reintroduce streaming mode behind a feature flag only.
- Evaluate and redesign stream source resolution architecture (not a direct ani-cli passthrough by default).
- Define robust metadata pipeline so source title/query/episode are reliable for Jimaku/subsync matching.
- Define a single end-to-end flow for stream subtitle source selection + fallback, including explicit failure handling.
- Add verification plan for playback and subtitle matching before implementation.

Notes:

- This ticket supersedes and consolidates the old fragmented streaming tasks:
  - TASK-48, TASK-48.1, TASK-48.2, TASK-48.3, TASK-48.4, TASK-48.5, TASK-48.6
  - TASK-51

Acceptance Criteria:

- One documented design + implementation plan exists for future reintroduction.
- A feature-flagged streaming mode path is specified with clear boundaries and testability.
- Stream source resolution and subtitle matching rules are deterministic and validated.
- No partial reintroduction until full path is implemented and tested.
<!-- SECTION:DESCRIPTION:END -->
