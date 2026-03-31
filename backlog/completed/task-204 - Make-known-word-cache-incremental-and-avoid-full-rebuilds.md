---
id: TASK-204
title: Make known-word cache incremental and avoid full rebuilds
status: Done
assignee:
  - Codex
created_date: '2026-03-19 19:05'
updated_date: '2026-03-19 19:12'
labels:
  - anki
  - cache
  - performance
dependencies: []
references:
  - src/anki-integration/known-word-cache.ts
  - src/anki-integration.ts
  - src/config/resolve/anki-connect.ts
  - src/config/definitions/defaults-integrations.ts
priority: high
ordinal: 105722
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Replace the known-word cache rebuild behavior with incremental synchronization. Startup should load existing cache state without immediately pulling all tracked Anki notes. Config-timed sync should reconcile adds, deletes, and in-place field edits against cached per-note state. Mined cards should optionally append their extracted words immediately after mining, enabled by default. Full rebuild should remain available only through explicit doctor tooling.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Known-word cache startup no longer performs an automatic full rebuild.
- [x] #2 Config-timed sync incrementally reconciles note additions, deletions, and edited word fields for the tracked known-word deck scope.
- [x] #3 Newly mined cards update the known-word cache immediately when the new config flag is enabled, and skip that fast path when disabled.
- [x] #4 Persisted cache state remains usable by stats endpoints that read the `words` set from disk.
- [x] #5 Regression tests cover startup behavior, incremental sync diffs, and the new config flag.
<!-- AC:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->

Known-word cache startup now loads persisted state and schedules sync based on refresh timing instead of wiping and rebuilding immediately. Persisted cache state now includes per-note word snapshots so timed refreshes can remove deleted notes, update edited notes, and keep the global `words` set stable for stats consumers. Added `ankiConnect.knownWords.addMinedWordsImmediately`, default `true`, so newly mined cards can update the cache immediately without waiting for the next timed sync.

Verification:

- `bun test src/anki-integration/known-word-cache.test.ts`
- `bun test src/config/resolve/anki-connect.test.ts src/config/config.test.ts`
- `bun test src/anki-integration.test.ts src/anki-integration/runtime.test.ts src/core/services/__tests__/stats-server.test.ts`
- `bun run test:config:src`
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run test:smoke:dist`

<!-- SECTION:OUTCOME:END -->
