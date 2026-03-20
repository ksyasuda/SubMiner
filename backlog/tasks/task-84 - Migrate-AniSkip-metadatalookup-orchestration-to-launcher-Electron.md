---
id: TASK-84
title: Migrate AniSkip metadata+lookup orchestration to launcher/Electron
status: Done
assignee:
  - Codex
created_date: '2026-03-03 08:31'
updated_date: '2026-03-16 05:13'
labels:
  - enhancement
  - aniskip
  - launcher
  - mpv-plugin
dependencies: []
references:
  - launcher/aniskip-metadata.ts
  - launcher/mpv.ts
  - plugin/subminer/aniskip.lua
  - plugin/subminer/options.lua
  - plugin/subminer/state.lua
  - plugin/subminer/lifecycle.lua
  - plugin/subminer/messages.lua
  - plugin/subminer.conf
  - launcher/aniskip-metadata.test.ts
documentation:
  - docs/mpv-plugin.md
  - launcher/aniskip-metadata.ts
  - plugin/subminer/aniskip.lua
  - docs/architecture.md
priority: medium
ordinal: 97500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Move AniSkip MAL/title-to-MAL lookup and intro payload resolution from mpv Lua to launcher Electron flow, while keeping mpv-side intro skip UX and chapter/chapter prompt behavior in plugin. Launcher should infer/analyze file metadata, fetch AniSkip payload when launching files, and pass resolved skip window via script options; plugin should trust launcher payload and fall back only when absent.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher infers AniSkip metadata for file targets using existing guessit/fallback logic and performs AniSkip MAL + payload resolution during mpv startup.
- [x] #2 Launcher injects script options containing resolved MAL id and intro window fields (or explicit lookup-failure status) into mpv startup.
- [x] #3 Lua plugin consumes launcher-provided AniSkip intro data and skips all network lookups when payload is present.
- [x] #4 Standalone mpv/plugin usage without launcher payload continues to function using existing async in-plugin lookup path.
- [x] #5 Docs and defaults are updated to document new script-option contract.
- [x] #6 Launcher tests cover payload generation contract and fallback behavior where metadata is unavailable.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add launcher-side AniSkip payload resolution helpers in launcher/aniskip-metadata.ts (MAL prefix lookup + AniSkip payload fetch + result normalization).
2. Wire launcher/mpv.ts + buildSubminerScriptOpts to pass resolved AniSkip fields/mode in --script-opts for file playback.
3. Update plugin/subminer/aniskip.lua plus options/state to consume injected payload: if intro_start/end present, apply immediately and skip network lookup; otherwise retain existing async behavior.
4. Ensure fallback for standalone mpv usage remains intact for no-launcher/manual refresh.
5. Add/update tests/docs/config references for new script-opt contract and edge cases.
<!-- SECTION:PLAN:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Executed end-to-end migration so launcher resolves AniSkip title/MAL/payload before mpv start and injects it via --script-opts. Plugin now parses and consumes launcher payload (JSON/url/base64), applies OP intro from payload, tracks payload metadata in state, and keeps legacy async lookup path for non-launcher/absent payload playback. Added launcher config key aniskip_payload and updated launcher/aniskip-metadata tests for resolve/payload behavior and contract validation.
<!-- SECTION:FINAL_SUMMARY:END -->
