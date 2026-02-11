---
id: TASK-5
title: Eliminate type duplication between renderer.ts and types.ts
status: To Do
assignee: []
created_date: '2026-02-11 08:20'
labels:
  - refactor
  - types
  - renderer
milestone: Codebase Clarity & Composability
dependencies: []
references:
  - src/renderer/renderer.ts
  - src/types.ts
  - src/main.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
renderer.ts locally redefines 20+ interfaces/types that already exist in types.ts: MergedToken, SubtitleData, MpvSubtitleRenderMetrics, Keybinding, SubtitlePosition, SecondarySubMode, all Jimaku/Kiku types, RuntimeOption types, SubsyncSourceTrack, SubsyncManualPayload, etc.

This creates divergence risk — changes in types.ts don't automatically propagate to the renderer's local copies.

Additionally, `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` and `sanitizeMpvSubtitleRenderMetrics()` exist only in renderer.ts despite being shared concerns (main.ts also has DEFAULT_MPV_SUBTITLE_RENDER_METRICS).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 All shared types are imported from types.ts — no local redefinitions in renderer.ts
- [ ] #2 DEFAULT_MPV_SUBTITLE_RENDER_METRICS lives in one canonical location (types.ts or a shared module)
- [ ] #3 sanitizeMpvSubtitleRenderMetrics moved to a shared module importable by both main and renderer
- [ ] #4 TypeScript compiles cleanly with no type errors
- [ ] #5 Renderer-only types (ChordAction, KikuModalStep, KikuPreviewMode) can stay local
<!-- AC:END -->
