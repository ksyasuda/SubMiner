# AGENTS.MD

## Quick Start

- Read [`docs-site/development.md`](./docs-site/development.md) and [`docs-site/architecture.md`](./docs-site/architecture.md) before substantial changes; follow them unless task requires deviation.
- Init workspace: `git submodule update --init --recursive`.
- Install deps: `make deps` or `bun install` plus `(cd vendor/texthooker-ui && bun install --frozen-lockfile)`.
- Fast dev loop: `make dev-watch`.
- Full local run: `bun run dev`.
- Verbose Electron debug: `electron . --start --dev --log-level debug`.

## Build / Test

- Use repo package manager/runtime only: Bun (`packageManager: bun@1.3.5`).
- Default handoff gate:
  `bun run typecheck`
  `bun run test:fast`
  `bun run test:env`
  `bun run build`
  `bun run test:smoke:dist`
- If `docs-site/` changed, also run:
  `bun run docs:test`
  `bun run docs:build`
- Formatting: prefer `make pretty` and `bun run format:check:src`; use `bun run format` only intentionally.
- Keep verification observable; capture failing command + exact error in notes/handoff.

## Change-Specific Checks

- Config/schema/defaults changes: run `bun run test:config`; if config template/defaults changed, run `bun run generate:config-example`.
- Launcher/plugin changes: run `bun run test:launcher` or `bun run test:env`; use `bun run test:launcher:smoke:src` for focused launcher e2e checks.
- Runtime-compat or compiled/dist-sensitive changes: run `bun run test:runtime:compat`.
- Docs-only changes: at least `bun run docs:test` if docs behavior/assertions changed; `bun run docs:build` before handoff.

## Generated / Sensitive Files

- Launcher source of truth: `launcher/*.ts`.
- Generated launcher artifact: `dist/launcher/subminer`; never hand-edit it.
- Repo-root `./subminer` is stale artifact path; do not revive/use it.
- `bun run build` rebuilds bundled Yomitan from `vendor/subminer-yomitan`; check submodules before debugging build failures.
- Avoid changing packaging/signing identifiers (`build.appId`, mac entitlements, signing-related settings) unless task explicitly requires it.

## Docs

- Docs site lives in-repo under [`docs-site/`](./docs-site/).
- Update docs for new/breaking behavior; no ship with stale docs.
- Make sure [`docs-site/changelog.md`](./docs-site/changelog.md) is updated on each release.

## PR Feedback

- Active PR: `gh pr view --json number,title,url --jq '"PR #\\(.number): \\(.title)\\n\\(.url)"'`.
- PR comments: `gh pr view …` + `gh api …/comments --paginate`.
- Replies: cite fix + file/line; resolve threads only after fix lands.

## Changelog

- User-visible PRs: add one fragment in `changes/*.md`.
- Fragment format:
  `type: added|changed|fixed|docs|internal`
  `area: <short-area>`
  blank line
  `- bullet`
- `changes/README.md`: instructions only; generator ignores it.
- No release-note entry wanted: use PR label `skip-changelog`.
- CI runs `bun run changelog:lint` + `bun run changelog:pr-check` on PRs.
- Release prep: `bun run changelog:build`, review `CHANGELOG.md` + `release/release-notes.md`, commit generated changelog + fragment deletions, then tag.
- Release CI expects committed changelog entry already present; do not rely on tag job to invent notes.

## Flow & Runtime

- Use Codex background for long jobs; tmux only for interactive/persistent (debugger/server).
- CI red: `gh run list/view`, rerun, fix, push, repeat til green.

## Language/Stack Notes

- Swift: use workspace helper/daemon; validate `swift build` + tests; keep concurrency attrs right.
- TypeScript: use repo PM; keep files small; follow existing patterns.

<!-- BACKLOG.MD MCP GUIDELINES START -->

<CRITICAL_INSTRUCTION>

## BACKLOG WORKFLOW INSTRUCTIONS

This project uses Backlog.md MCP for all task and project management activities.

**CRITICAL GUIDANCE**

- If your client supports MCP resources, read `backlog://workflow/overview` to understand when and how to use Backlog for this project.
- If your client only supports tools or the above request fails, call `backlog.get_workflow_overview()` tool to load the tool-oriented overview (it lists the matching guide tools).

- **First time working here?** Read the overview resource IMMEDIATELY to learn the workflow
- **Already familiar?** You should have the overview cached ("## Backlog.md Overview (MCP)")
- **When to read it**: BEFORE creating tasks, or when you're unsure whether to track work

These guides cover:

- Decision framework for when to create tasks
- Search-first workflow to avoid duplicates
- Links to detailed guides for task creation, execution, and finalization
- MCP tools reference

You MUST read the overview resource to understand the complete workflow. The information is NOT summarized here.

</CRITICAL_INSTRUCTION>

<!-- BACKLOG.MD MCP GUIDELINES END -->
