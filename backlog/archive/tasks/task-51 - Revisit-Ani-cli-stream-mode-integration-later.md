---
id: TASK-51
title: Revisit Ani-cli stream mode integration later
status: To Do
assignee: []
created_date: '2026-02-14 08:19'
labels:
  - someday
  - streaming
  - ani-cli
  - feature-flag
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current codebase has removed ani-cli integration and stream-mode from subminer temporarily. Keep a deferred design task to reintroduce streaming mode in a future cycle.

Findings from prior attempts:
- `subminer -s <query>` path relied on `ani-cli` resolving stream URLs, but returned stream URLs that are Cloudflare-protected (`tools.fast4speed.rsvp`) and often returned 403 from mpv/ytdl-hook (generic anti-bot/Forbidden).
- Even after passing `ytdl` extractor args, stream playback via subminer still failed because URL/anti-bot handling differed from direct ani-cli execution context.
- We also observed stream title resolution issues: selected titles from ani-cli menu were unreliable/random and broke downstream Jimaku matching behavior.
- ffsubsync failures were difficult to debug initially due to OSD-only error visibility; logging was added to mpv log path.
- Based on these findings and instability, stream mode should be explicitly deferred rather than partially reintroduced.

Proposal:
- Reintroduce behind a feature flag / future milestone only.
- Re-design around a dedicated stream source resolver with robust URL acquisition and source metadata preservation (query/episode/title) before subtitle sync flows.
<!-- SECTION:DESCRIPTION:END -->
