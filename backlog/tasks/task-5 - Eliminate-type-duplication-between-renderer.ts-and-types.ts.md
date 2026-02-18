---
id: TASK-5
title: Eliminate type duplication between renderer.ts and types.ts
status: Done
assignee:
  - codex
created_date: '2026-02-11 08:20'
updated_date: '2026-02-18 04:11'
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
ordinal: 53000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
renderer.ts locally redefines 20+ interfaces/types that already exist in types.ts: MergedToken, SubtitleData, MpvSubtitleRenderMetrics, Keybinding, SubtitlePosition, SecondarySubMode, all Jimaku/Kiku types, RuntimeOption types, SubsyncSourceTrack, SubsyncManualPayload, etc.

This creates divergence risk — changes in types.ts don't automatically propagate to the renderer's local copies.

Additionally, `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` and `sanitizeMpvSubtitleRenderMetrics()` exist only in renderer.ts despite being shared concerns (main.ts also has DEFAULT_MPV_SUBTITLE_RENDER_METRICS).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All shared types are imported from types.ts — no local redefinitions in renderer.ts
- [x] #2 DEFAULT_MPV_SUBTITLE_RENDER_METRICS lives in one canonical location (types.ts or a shared module)
- [x] #3 sanitizeMpvSubtitleRenderMetrics moved to a shared module importable by both main and renderer
- [x] #4 TypeScript compiles cleanly with no type errors
- [x] #5 Renderer-only types (ChordAction, KikuModalStep, KikuPreviewMode) can stay local
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Audit `src/renderer/renderer.ts`, `src/types.ts`, and `src/main.ts` to identify every duplicated type and both subtitle-metrics helpers/constants.
2) Remove duplicated shared type declarations from `renderer.ts` and replace them with direct imports from `types.ts`; keep renderer-only types (`ChordAction`, `KikuModalStep`, `KikuPreviewMode`) local.
3) Create a single canonical home for `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` in a shared location and update all call sites to import it from that canonical module.
4) Move `sanitizeMpvSubtitleRenderMetrics` to a shared module importable by both main and renderer, then switch both files to consume that shared implementation.
5) Run TypeScript compile/check, fix any fallout, and verify no local shared-type redefinitions remain in `renderer.ts`.
6) Update task notes and check off acceptance criteria as each item is validated.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Replaced renderer-local shared type declarations with `import type` from `src/types.ts`; only `KikuModalStep`, `KikuPreviewMode`, and `ChordAction` remain local in renderer.

Moved canonical MPV subtitle metrics defaults to `src/core/services/mpv-render-metrics-service.ts` as `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` and switched `main.ts` to consume it.

Added shared `sanitizeMpvSubtitleRenderMetrics` export in `src/core/services/mpv-render-metrics-service.ts` and re-exported it from `src/core/services/index.ts`.

Removed renderer-local `DEFAULT_MPV_SUBTITLE_RENDER_METRICS`, `coerceFiniteNumber`, and `sanitizeMpvSubtitleRenderMetrics`; renderer now consumes full sanitized metrics from main via IPC and keeps nullable local state until startup metrics are loaded.

Validation: `pnpm run build` passed; `node --test dist/core/services/mpv-render-metrics-service.test.js` passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented TASK-5 by eliminating duplicate shared type/model definitions from the renderer and centralizing MPV subtitle render metrics primitives.

What changed:
- `src/renderer/renderer.ts`
  - Removed local redefinitions of shared interfaces/types (subtitle, keybinding, Jimaku, Kiku, runtime options, subsync, render metrics).
  - Added `import type` usage from `src/types.ts` for shared contracts.
  - Kept renderer-only local types (`KikuModalStep`, `KikuPreviewMode`, `ChordAction`).
  - Removed renderer-local metrics default/sanitization helpers and switched invisible overlay resize behavior to rely on already-synced metrics state from main.
- `src/core/services/mpv-render-metrics-service.ts`
  - Added canonical `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` export.
  - Added shared `sanitizeMpvSubtitleRenderMetrics` export.
- `src/core/services/index.ts`
  - Re-exported `DEFAULT_MPV_SUBTITLE_RENDER_METRICS` and `sanitizeMpvSubtitleRenderMetrics`.
- `src/main.ts`
  - Removed local default metrics constant and imported the canonical default from core services.
- `src/core/services/mpv-render-metrics-service.test.ts`
  - Updated base fixture to derive from canonical default metrics constant.

Why:
- Prevent type drift between renderer and shared contracts.
- Establish a single source of truth for MPV subtitle render metric defaults and sanitization utilities.

Validation:
- `pnpm run build`
- `node --test dist/core/services/mpv-render-metrics-service.test.js`

Result:
- Shared renderer types now come from `types.ts`.
- MPV subtitle render metrics defaults are canonicalized in one shared module.
- TypeScript compiles cleanly and relevant metrics service tests pass.
<!-- SECTION:FINAL_SUMMARY:END -->
