<!-- read_when: using or modifying repo-local agent plugins -->

# Agent Plugins

Status: active
Last verified: 2026-03-26
Owner: Kyle Yasuda
Read when: packaging or migrating repo-local agent workflow skills into plugins

## SubMiner Workflow Plugin

- Canonical plugin path: `plugins/subminer-workflow/`
- Marketplace catalog: `.agents/plugins/marketplace.json`
- Canonical skill sources:
  - `plugins/subminer-workflow/skills/subminer-scrum-master/`
  - `plugins/subminer-workflow/skills/subminer-change-verification/`

## Migration Rule

- Plugin-owned skills are the source of truth.
- `.agents/skills/subminer-*` remain only as compatibility shims.
- Existing script entrypoints under `.agents/skills/subminer-change-verification/scripts/` stay as wrappers so historical commands do not break.

## Backlog

- Prefer Backlog.md MCP when the host session exposes it.
- If MCP is unavailable, use repo-local `backlog/` files and record that fallback.

## Verification

- For plugin/docs-only changes, start with `bun run test:docs:kb`.
- Use the plugin-owned verifier when the change crosses from docs into scripts or workflow logic.
