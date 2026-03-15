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
- Docs-only: `bun run docs:test`, then `bun run docs:build`

## Sensitive Files

- Launcher source of truth: `launcher/*.ts`
- Generated launcher artifact: `dist/launcher/subminer`; never hand-edit it
- Repo-root `./subminer` is stale; do not revive it
- `bun run build` rebuilds bundled Yomitan from `vendor/subminer-yomitan`
- Do not change signing/packaging identifiers unless the task explicitly requires it

## Release / PR Notes

- User-visible PRs need one fragment in `changes/*.md`
- CI enforces `bun run changelog:lint` and `bun run changelog:pr-check`
- PR review helpers:
  - `gh pr view --json number,title,url --jq '"PR #\\(.number): \\(.title)\\n\\(.url)"'`
  - `gh api repos/:owner/:repo/pulls/<num>/comments --paginate`

## Runtime Notes

- Use Codex background for long jobs; tmux only when persistence/interaction is required
- CI red: `gh run list/view`, rerun, fix, repeat until green
- TypeScript: keep files small; follow existing patterns
- Swift: use workspace helper/daemon; validate `swift build` + tests

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
