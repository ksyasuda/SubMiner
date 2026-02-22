# Agent: `codex-aniskip-intro-skip-20260222T080257Z-51fx`

- alias: `codex-aniskip-intro-skip`
- mission: `Port intro skip to AniSkip API in mpv plugin with OSD skip button`
- status: `handoff`
- branch: `main`
- started_at: `2026-02-22T08:02:57Z`
- heartbeat_minutes: `5`

## Current Work (newest first)
- [2026-02-22T08:03:20Z] intent: inspect plugin flow + ani-skip script behavior; design direct API integration.
- [2026-02-22T08:13:40Z] progress: added launcher guessit metadata extraction + script-opts passthrough (`aniskip_title`, `aniskip_season`, `aniskip_episode`) for file playback.
- [2026-02-22T08:13:40Z] progress: updated plugin intro hint behavior to show `You can skip by pressing y-k` for first 3 seconds from intro start.
- [2026-02-22T08:13:40Z] test: `bun test launcher/aniskip-metadata.test.ts`, `bun test launcher/mpv.test.ts`, `luac -p plugin/subminer.lua`, `bun run tsc --noEmit`.

## Files Touched
- `plugin/subminer.lua`
- `plugin/subminer.conf`
- `docs/mpv-plugin.md`
- `launcher/aniskip-metadata.ts`
- `launcher/aniskip-metadata.test.ts`
- `launcher/mpv.ts`

## Assumptions
- mpv build supports `subprocess` curl fallback for HTTP.
- intro skip should be OP-only by default.

## Open Questions / Blockers
- none

## Next Step
- manual mpv runtime validation still needed (network + real media metadata).
- [2026-02-22T09:25:57Z] progress: fixed intro skip keypath reliability by always binding fallback `y-k` alias; added explicit OSD feedback when skip not available/outside intro window.
- [2026-02-22T09:25:57Z] test: `luac -p plugin/subminer.lua`; `bun run tsc --noEmit`.
- [2026-02-22T09:30:27Z] progress: added explicit AniSkip query logging in plugin (`title/season/episode` context, MAL lookup query, resolved MAL id, AniSkip URL) for season mismatch diagnosis.
- [2026-02-22T09:30:27Z] test: `luac -p plugin/subminer.lua`.
- [2026-02-22T09:33:22Z] progress: improved MAL resolution heuristic for AniSkip (title overlap + season-aware ranking) to avoid sequel mis-selection; logs chosen MAL candidate title/id/score.
- [2026-02-22T09:33:22Z] test: `luac -p plugin/subminer.lua`.
- [2026-02-22T09:36:05Z] progress: hardened MAL candidate scoring (token overlap + stopword filtering + coverage + low-confidence rejection threshold) to prevent unrelated anime selection; syntax check passed.
- [2026-02-22T09:42:00Z] progress: added MAL lookup fallback attempts across multiple title sources (`aniskip_title`, `media-title`, filename, path) with per-attempt logs; keeps season/episode fixed.
- [2026-02-22T09:57:40Z] progress: tightened title overlap scoring for MAL candidate selection (strong penalty for partial multi-token matches) to prevent wrong-ID picks like "Shadow Skill" from "The Eminence in Shadow".
- [2026-02-22T09:57:40Z] test: `luac -p plugin/subminer.lua`.
- [2026-02-22T10:04:20Z] progress: launcher AniSkip metadata now prioritizes guessit `series` over episode `title`; fallback title extraction now prefers anime directory name before season dir (`.../<Show>/Season-1/<file>`).
- [2026-02-22T10:04:20Z] test: `bun test launcher/aniskip-metadata.test.ts`, `bun test launcher/mpv.test.ts`, `bun run tsc --noEmit`, `luac -p plugin/subminer.lua`.
- [2026-02-22T10:10:30Z] progress: plugin now extracts show title from media path (`.../<Show>/Season-*`) and prioritizes it over `aniskip_title`, preventing episode-title queries when launcher metadata is stale/wrong.
- [2026-02-22T10:10:30Z] test: `luac -p plugin/subminer.lua`.
- [2026-02-22T10:14:40Z] progress: per user request removed MAL scoring/rejection; resolver now selects first MAL search result with an id and logs selected first result.
- [2026-02-22T10:14:40Z] test: `luac -p plugin/subminer.lua`.
