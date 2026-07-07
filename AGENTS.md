# AGENTS.MD

## Internal Docs

Start here, then leave this file.

- Internal system of record: [`docs/README.md`](./docs/README.md)
- Architecture map: [`docs/architecture/README.md`](./docs/architecture/README.md)
- Workflow map: [`docs/workflow/README.md`](./docs/workflow/README.md)
- Verification lanes: [`docs/workflow/verification.md`](./docs/workflow/verification.md)
- Knowledge-base rules: [`docs/knowledge-base/README.md`](./docs/knowledge-base/README.md)
- Release guide: [`docs/RELEASING.md`](./docs/RELEASING.md)

`docs-site/` is user-facing. Do not treat it as the canonical internal source of truth.

`CLAUDE.md` is a symlink to this file — there is one project instruction file, not two.

## Quick Start

- Init workspace: `git submodule update --init --recursive`
- Install deps: `make deps` or `bun install` plus `(cd vendor/texthooker-ui && bun install --frozen-lockfile)`
- Fast dev loop: `make dev-watch`
- Full local run: `bun run dev`
- Verbose Electron debug: `electron . --start --dev --log-level debug`

## Build / Test

- Runtime/package manager: Bun (`packageManager: bun@1.3.5`)
- Default handoff gate:
  `bun run typecheck`
  `bun run test:fast`
  `bun run test:env`
  `bun run build`
  `bun run test:smoke:dist`
- If `docs-site/` changed, also run:
  `bun run docs:test`
  `bun run docs:build`
- Prefer `make pretty` and `bun run format:check:src`

## Change-Specific Checks

- Config/schema/defaults: `bun run test:config`; if template/defaults changed, `bun run generate:config-example`
- Launcher/plugin: `bun run test:launcher` or `bun run test:env`
- Runtime-compat / dist-sensitive: `bun run test:runtime:compat`
- Stats dashboard UI (`stats/`): `bun run test:stats`
- Build/release scripts (`scripts/**`): `bun run test:scripts`
- Docs-only: `bun run docs:test`, then `bun run docs:build`
- Test lanes are directory-discovered via `scripts/test-lanes.ts`; never hand-list test files in `package.json`

## Docs Upkeep

- Docs ship with the change, not after. If a change alters behavior, defaults, flags, shortcuts, ports, or APIs, update the matching docs in the same PR. Touching code without reconciling its docs is an incomplete change.
- Source of truth for config defaults is the generated `config.example.jsonc`. Never write a default value into prose you didn't read from it — and don't restate the same default across multiple docs; cite/link to one place so there's a single thing to update.
- Trigger map (touch left → update right):
  - `src/config/definitions/**` (schema/defaults/template) → `bun run generate:config-example`, then reconcile `docs-site/configuration.md` + any feature doc that cites that default
  - shortcuts/keybindings (`shortcuts.*`, `keybindings`, `stats.*Key`, `subtitleSidebar.toggleKey`, controller bindings) → `docs-site/shortcuts.md`
  - CLI flags/subcommands (`src/cli/args.ts`, `launcher/**`) → `docs-site/usage.md` + relevant integration doc
  - feature behavior (anki / jellyfin / jimaku / anilist / youtube / immersion / stats / websocket / sidebar / character-dictionary / annotations) → matching `docs-site/<feature>.md`
  - architecture / IPC / workflow / internal process → internal `docs/` (system of record)
  - feature set / requirements / install flow → `README.md`
- Removing or renaming a config key: grep `docs-site/` and `docs/` for the old key and any value it documented; legacy/hidden keys (`LEGACY_HIDDEN_CONFIG_PATHS`) should not appear in user docs as current settings.
- Verify after doc edits: `bun run verify:config-example` (if config touched), `bun run docs:test`, `bun run docs:build`.

## Sensitive Files

- Launcher source of truth: `launcher/*.ts`
- Generated launcher artifact: `dist/launcher/subminer`; never hand-edit it
- Repo-root `./subminer` is stale; do not revive it
- `bun run build` rebuilds bundled Yomitan from `vendor/subminer-yomitan`
- Do not change signing/packaging identifiers unless the task explicitly requires it

## Release / PR Notes

- User-visible PRs need reconciled current-outcome fragment(s) in `changes/*.md` — format and rules in [`changes/README.md`](./changes/README.md) (`type` + `area` keys required; inspect existing same-PR fragments, then update/remove stale bullets or add only genuinely separate outcomes; apply the `skip-changelog` label to opt out)
- User-visible docs changes get a `type: docs` fragment
- CI enforces `bun run changelog:lint` and `bun run changelog:pr-check`
- PR review helpers:
  - `gh pr view --json number,title,url --jq '"PR #\\(.number): \\(.title)\\n\\(.url)"'`
  - `gh api repos/:owner/:repo/pulls/<num>/comments --paginate`

## Runtime Notes

- Use Codex background for long jobs; tmux only when persistence/interaction is required
- CI red: `gh run list/view`, rerun, fix, repeat until green
- TypeScript: keep files small; follow existing patterns
- Only Swift is the `scripts/get-mpv-window-macos.swift` helper (macOS mpv window detection); validate via `bun test scripts/get-mpv-window-macos.test.ts`
