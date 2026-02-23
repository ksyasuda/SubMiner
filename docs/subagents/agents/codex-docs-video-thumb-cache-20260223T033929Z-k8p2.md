# Agent: `codex-docs-video-thumb-cache-20260223T033929Z-k8p2`

- alias: `codex-docs-video-thumb-cache`
- mission: `Fix docs landing page demo video thumbnail staleness after direct asset replacement`
- status: `handoff`
- branch: `main`
- started_at: `2026-02-23T03:40:48Z`
- heartbeat_minutes: `5`

## Current Work
- [2026-02-23T03:40:48Z] intent: confirm root cause for stale docs video thumbnail and apply minimal cache-busting/docs fix.
- [2026-02-23T03:40:48Z] plan: touch `docs/index.md`; add regression coverage if lightweight; validate with docs build/targeted test.
- [2026-02-23T03:40:48Z] assumption: stale thumbnail caused by static poster/source URLs reusing identical path after direct file replacement.
- [2026-02-23T03:42:25Z] progress: added failing regression test `docs/index.assets.test.ts` asserting shared cache-busting token across demo media URLs.
- [2026-02-23T03:42:58Z] progress: generated fresh poster from `docs/public/assets/minecard.webm` (`minecard-poster.jpg`) and wired docs video block to shared `demoAssetVersion` tokenized URLs.
- [2026-02-23T03:43:20Z] test: `bun test docs/index.assets.test.ts` passes (1/1).
- [2026-02-23T03:43:35Z] test: `bun run docs:build` passes.
- [2026-02-23T03:43:55Z] handoff: TASK-112 marked Done with implementation notes and validation commands.
- [2026-02-23T03:44:04Z] progress: regenerated `minecard-poster.jpg` from `minecard.webm` at exactly `00:00:12` and bumped `demoAssetVersion` to `20260223-2` to force refresh.
- [2026-02-23T03:44:04Z] test: reran `bun test docs/index.assets.test.ts` (pass).

## Files Touched
- `docs/subagents/agents/codex-docs-video-thumb-cache-20260223T033929Z-k8p2.md`
- `docs/subagents/INDEX.md`
- `docs/subagents/collaboration.md`
- `backlog/tasks/task-112 - Fix-docs-demo-video-thumbnail-cache-staleness.md`
- `docs/index.assets.test.ts`
- `docs/index.md`
- `docs/public/assets/minecard-poster.jpg`

## Assumptions
- Browser/CDN cache can keep prior media when URL path remains unchanged.
- Docs home video section in `docs/index.md` is the only requested scope.

## Open Questions / Blockers
- None.

## Next Step
- Share fix summary and ask user to hard-refresh docs page once (`Cmd+Shift+R`) to bypass local cache.
