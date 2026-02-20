# Agent: codex-main

- alias: planner-exec
- mission: Fix frequency/N+1 regression in plugin --start flow
- status: in_progress
- branch: main
- started_at: 2026-02-19T08:06:28Z
- heartbeat_minutes: 20

## Current Work (newest first)

- [2026-02-19T19:36:46Z] progress: config confirmed frequency enabled (`~/.config/SubMiner/config.jsonc`); likely mode-latch issue after initial `--texthooker`.
- [2026-02-19T19:36:46Z] change: in `src/main.ts`, `handleCliCommand` now disables `texthookerOnlyMode` on follow-up `--start`/overlay commands and triggers background warmups.
- [2026-02-19T19:36:46Z] test: `bun run build` pass; `bun test src/core/services/cli-command.test.ts src/cli/args.test.ts src/core/services/startup-bootstrap.test.ts` pass (28/28).
- [2026-02-19T19:29:00Z] progress: found second-instance toggle gap: MPV connect was gated to initial-source toggles only; plugin toggle handoff could show overlay without subtitle stream/tokenization.
- [2026-02-19T19:29:00Z] change: updated `src/core/services/cli-command.ts` so toggle/toggle-visible/toggle-invisible trigger MPV connect regardless of source when app is already running.
- [2026-02-19T19:29:00Z] test: added `handleCliCommand connects MPV for toggle on second-instance`; `bun test src/core/services/cli-command.test.ts` pass (19/19).
- [2026-02-19T19:24:18Z] progress: root cause confirmed: subtitle tokenization is skipped when no overlay windows (`src/main.ts`), and `--start` command path did not initialize overlay runtime.
- [2026-02-19T19:24:18Z] change: `src/core/services/cli-command.ts` now initializes overlay runtime for `--start`; second-instance `--start` is ignored only when overlay runtime is already initialized.
- [2026-02-19T19:24:18Z] test: `src/core/services/cli-command.test.ts` covers start-driven init + second-instance behaviors; `bun test src/core/services/cli-command.test.ts` pass (18/18).
- [2026-02-19T19:04:58Z] intent: investigate report that frequency tracking resolves from app config location instead of active app binary root during just-built binary tests; add regression test first, then patch runtime path resolution.
- [2026-02-19T19:04:58Z] planned files: `src/main/frequency-dictionary-runtime.ts`, `src/core/services/frequency-dictionary.ts`, `src/core/services/frequency-dictionary.test.ts`, `backlog/tasks/*` (ticket link).
- [2026-02-19T19:04:58Z] assumptions: runtime should prefer binary-adjacent frequency assets when `SUBMINER_BINARY_PATH` points to a built app binary; config-path-based defaults should remain fallback.
- [2026-02-19T19:05:00Z] backlog: linked work to `TASK-87` (`backlog/tasks/task-87 - Resolve-frequency-storage-path-relative-to-active-app-binary.md`).
- [2026-02-19T16:58:00Z] intent: investigate AniList login regression (`unsupported_grant_type`), add callback token-consumption tests first, then wire `openAnilistSetupWindow` navigation handlers to persist OAuth token from callback URL/hash.
- [2026-02-19T16:54:18Z] handoff: implemented inline same-line option comments in config template output, including enum/boolean value lists where known; regenerated `config.example.jsonc` and `docs/public/config.example.jsonc`.
- [2026-02-19T16:54:18Z] test: `bun run build && node --test dist/config/config.test.js` -> pass (33/33); macOS helper compile sandbox-denied cache path fallback unchanged.
- [2026-02-19T16:51:29Z] intent: user requested inline same-line comments for each example config option with short purpose + allowed values where enum-like; target `config.example.jsonc` (+ `docs/public/config.example.jsonc` parity if needed).
- [2026-02-19T10:24:42Z] progress: continued TASK-85 decomposition slices in `src/main.ts`; extracted config-derived runtime wrappers, AniList setup helpers, clipboard queue runtime, AniList state runtime, immersion media runtime, startup config runtime, and main subsync runtime under `src/main/runtime/`; updated call sites.
- [2026-02-19T10:24:42Z] test: `bun run build` pass (sandbox Swift cache fallback unchanged), `node --test dist/main/runtime/anilist-setup.test.js dist/main/runtime/clipboard-queue.test.js dist/main/runtime/anilist-state.test.js dist/main/runtime/immersion-media.test.js dist/main/runtime/immersion-startup.test.js dist/main/runtime/startup-config.test.js dist/main/config-validation.test.js dist/core/services/app-ready.test.js dist/core/services/startup-bootstrap.test.js` -> pass (35/35).
- [2026-02-19T10:05:25Z] progress: continued TASK-85 refactor execution; extracted app-ready config runtime handlers to `src/main/runtime/startup-config.ts` + tests and immersion tracker startup handler to `src/main/runtime/immersion-startup.ts` + tests; wired `src/main.ts` call sites to factories; `src/main.ts` now 3250 LOC.
- [2026-02-19T10:05:25Z] test: `bun run build` pass (same macOS helper sandbox fallback), `node --test dist/main/runtime/immersion-startup.test.js dist/main/runtime/startup-config.test.js dist/main/config-validation.test.js dist/core/services/app-ready.test.js dist/core/services/startup-bootstrap.test.js` -> pass (24/24).
- [2026-02-19T09:55:47Z] handoff: initialized Backlog non-interactive (`backlog init --defaults --integration-mode mcp --auto-open-browser false`), moved TASK-85 to In Progress, completed Task 2 guardrail slice (`scripts/check-file-budgets.ts`, docs, package scripts), and started Task 3 with first `src/main.ts` extraction (`src/main/config-validation.ts` + tests); build/tests pass.
- [2026-02-19T09:55:47Z] test: `bun run check:file-budgets` (warn list), `bun run check:file-budgets:strict` (expected fail), `bun run build` (pass; macOS helper falls back to source due sandbox cache permission), `node --test dist/main/config-validation.test.js dist/core/services/app-ready.test.js dist/core/services/startup-bootstrap.test.js` (15/15 pass).
- [2026-02-19T09:47:52Z] handoff: delivered concrete maintainability plan at `docs/plans/2026-02-19-repo-maintainability-refactor-plan.md` plus umbrella ticket `TASK-85`; ready for execution-mode choice.
- [2026-02-19T09:45:26Z] intent: user requested concrete maintainability/readability plan via writing-plans + refactor skills; focus hotspots `src/main.ts`, `src/anki-integration.ts`, `plugin/subminer.lua`, generated `subminer` launcher artifact strategy.
- [2026-02-19T09:45:26Z] progress: scanned LOC hotspots and module graph; confirmed `subminer` is generated/ignored build artifact (`make build-launcher`) and major live hotspots are `src/main.ts` (3316 LOC), `src/anki-integration.ts` (1703 LOC), `src/config/service.ts` (1565 LOC), `src/core/services/immersion-tracker-service.ts` (1470 LOC).
- [2026-02-19T09:05:26Z] handoff: completed TASK-70 (Done); unified config path resolution into shared `src/config/path-resolution.ts`, wired app+launcher call sites, added precedence tests, verified build/tests/launcher bundle, and checked AC/DoD in Backlog.
- [2026-02-19T09:05:26Z] test: `bun run build && node --test dist/config/path-resolution.test.js dist/config/config.test.js && bun build ./launcher/main.ts --target=bun --packages=bundle --outfile=/tmp/subminer-task70` -> pass.
- [2026-02-19T08:57:36Z] progress: plan saved to `docs/plans/2026-02-19-task-70-unify-config-path-resolution.md`, mirrored to backlog TASK-70 `planSet`; moving to edit/test execution.
- [2026-02-19T08:52:23Z] intent: execute TASK-70 by reading full backlog context, writing implementation plan via writing-plans skill, then executing via executing-plans skill with parallel subagents where possible.
- [2026-02-19T08:41:44Z] handoff: created backlog TASK-84 for config-gated keybindings + disabled-feature integration non-loading (Jellyfin example) with tests/docs acceptance criteria.
- [2026-02-19T08:41:22Z] intent: create backlog task for config-gated keybindings so disabled features become no-ops and feature modules (example: Jellyfin) are not loaded when disabled.
- [2026-02-19T08:40:53Z] handoff: completed TASK-83 simplification; removed configurable sentence/audio field overrides for Lapis sentence cards and verified.
- [2026-02-19T08:40:53Z] test: `bun run build && bun run test:config:dist` -> pass (33/33).
- [2026-02-19T08:38:30Z] intent: implement TASK-83 to remove `isLapis.sentenceCardSentenceField` + `isLapis.sentenceCardAudioField` from config/types/docs and force fixed `Sentence`/`SentenceAudio` in integration.
- [2026-02-19T08:27:23Z] handoff: completed TASK-69 implementation + verification; backlog TASK-69 set Done with AC/DoD checked.
- [2026-02-19T08:27:23Z] test: `bun run build && bun run test:config:dist` -> pass (32/32).
- [2026-02-19T08:27:23Z] progress: hardened legacy ankiConnect migration mapping with validators + warnings in `src/config/service.ts`; added invalid/valid legacy regression tests in `src/config/config.test.ts`; updated migration doc note in `docs/configuration.md`.
- [2026-02-19T08:23:30Z] intent: user approved TASK-69 plan; start implementation (tests first, then validated legacy mapping, then verify + backlog updates).
- [2026-02-19T08:21:11Z] handoff: completed TASK-38 implementation + verification; updated Backlog acceptance checks and notes.
- [2026-02-19T08:21:11Z] test: `bun run build && node --test dist/config/config.test.js dist/core/services/app-ready.test.js` -> pass (39/39).
- [2026-02-19T08:21:11Z] progress: wired strict startup parse failure handling + warning summary + critical validation dialog path in `src/main.ts` and threaded new app-ready dependency in `src/main/app-lifecycle.ts`.
- [2026-02-19T08:20:00Z] intent: read TASK-69 details, write end-to-end implementation plan via writing-plans skill, then execute via executing-plans skill with parallel subagents where possible.
- [2026-02-19T08:18:47Z] progress: implemented unknown top-level config warning + `ankiConnect.openRouter` deprecation warning; added coverage in `src/config/config.test.ts`; verification passed (`bun run build && node --test dist/config/config.test.js`).
- [2026-02-19T08:10:40Z] intent: implement TASK-38 plan Task 1 scope (`src/config/service.ts`, `src/config/config.test.ts`) for unknown top-level key warnings + `ankiConnect.openRouter` deprecation warning, then run required build+test command.
- [2026-02-19T08:09:46Z] progress: wrote plan `docs/plans/2026-02-19-task-38-config-validation-startup.md`, set TASK-38 In Progress, and stored plan in Backlog.
- [2026-02-19T08:06:28Z] intent: read backlog/task context, draft implementation plan for TASK-38, then execute with parallel subagents where possible.

## Files Touched

- `docs/subagents/INDEX.md`
- `docs/subagents/agents/codex-main.md`
- `docs/plans/2026-02-19-task-70-unify-config-path-resolution.md`
- `src/config/path-resolution.ts`
- `src/config/path-resolution.test.ts`
- `launcher/main.ts`
- `launcher/config.ts`
- `src/main.ts`
- `package.json`
- `backlog/tasks/task-70 - Unify-config-path-resolution-across-main-and-launcher.md`
- `backlog/tasks/task-84 - Gate-feature-dependent-keybindings-behind-config-flags.md`
- `src/config/service.ts`
- `src/config/config.test.ts`
- `src/core/services/startup.ts`
- `src/core/services/app-ready.test.ts`
- `src/main.ts`
- `src/main/app-lifecycle.ts`
- `docs/plans/2026-02-19-task-38-config-validation-startup.md`
- `docs/plans/2026-02-19-task-69-legacy-config-migration-hardening.md`
- `docs/configuration.md`
- `src/types.ts`
- `src/config/definitions.ts`
- `src/anki-integration.ts`
- `config.example.jsonc`
- `docs/public/config.example.jsonc`
- `docs/anki-integration.md`
- `docs/mining-workflow.md`
- `backlog/tasks/task-83 - Simplify-isLapis-sentence-card-field-config-to-fixed-field-names.md`
- `backlog/tasks/task-38 - Add-user-friendly-config-validation-errors-on-startup.md`
- `docs/plans/2026-02-19-repo-maintainability-refactor-plan.md`
- `backlog/tasks/task-85 - Refactor-large-files-for-maintainability-and-readability.md`
- `backlog/config.yml`
- `scripts/check-file-budgets.ts`
- `docs/file-size-budgets.md`
- `src/main/config-validation.ts`
- `src/main/config-validation.test.ts`

## Assumptions

- AGENT_ID/AGENT_ALIAS env vars are unset; using stable fallback `codex-main` / `planner-exec` for this run.

## Open Questions / Blockers

- none

## Next Step

- Continue Task 3: extract next `src/main.ts` runtime slice (startup reload/logging branch) into `src/main/runtime/*` with seam tests, then re-run build + targeted core tests.
- [2026-02-19T17:20:00Z] intent: improve launcher help output so top-level help auto-lists available subcommands from command registry; add/adjust tests if present.

- [2026-02-19T16:56:43Z] intent: improve launcher help output so top-level help auto-lists available subcommands from command registry; add/adjust tests if present.

- [2026-02-19T16:59:23Z] progress: added auto-generated root help subcommand section by rendering Commander subcommand registry (`launcher/config.ts`) and covered with launcher help regression test (`launcher/config.test.ts`).
- [2026-02-19T16:59:23Z] test: `bun test launcher/config.test.ts` pass; `make build-launcher && ./subminer -h` shows Commands list with jellyfin/yt/doctor/config/mpv/texthooker.
- [2026-02-19T16:59:23Z] handoff: completed requested launcher help improvement; top-level `-h` now includes available subcommands without hard-coded help string updates.

- [2026-02-19T17:07:25Z] progress: added launcher passthrough subcommand `app|bin` that forwards raw args directly to SubMiner binary (`launcher/config.ts`, `launcher/main.ts`, `launcher/types.ts`).
- [2026-02-19T17:07:25Z] test: `bun test launcher/config.test.ts launcher/parse-args.test.ts` pass; passthrough smoke via stub binary env (`SUBMINER_APPIMAGE_PATH=/tmp/subminer-app-stub.sh bun run launcher/main.ts app --anilist --anilist-status`) forwards raw args; `make build-launcher && ./subminer -h` shows `app|bin`.
- [2026-02-19T17:07:25Z] handoff: user-requested direct app/binary passthrough implemented; supports AniList trigger via `subminer app --anilist`.
- [2026-02-19T17:34:30Z] progress: hardened `app|bin` passthrough to bypass Commander parsing and forward raw argv suffix verbatim after subcommand token (`launcher/config.ts`).
- [2026-02-19T17:34:30Z] test: `bun test launcher/config.test.ts launcher/parse-args.test.ts` pass; stub verification confirms exact forwarding for `app --start --anilist-setup -h`; rebuilt wrapper via `make build-launcher`.
