---
name: subminer-change-verification
description: Verify SubMiner changes with repo-native cheap-first test lanes. Use after code, config, launcher, plugin, runtime, stats, documentation, or workflow changes; do not use for read-only questions.
---

# SubMiner Change Verification

Verify the behavior claimed by a change without running unrelated expensive checks by default.

## Workflow

1. Inspect the requested scope and changed paths with `git status --short` and `git diff`.
2. Read `docs/workflow/verification.md` as the source of truth for maintained lanes.
3. Run the cheapest lane or lanes that cover the changed behavior.
4. Escalate to the full handoff gate only for substantial or cross-boundary changes.
5. Report exact commands, results, skipped checks, blockers, and remaining risk.

Do not use hidden wrapper commands. Verification commands are owned by `package.json` and the workflow documentation.

## Lane Selection

- Internal docs, `AGENTS.md`, or `.agents/skills/**`: `bun run test:docs:kb`
- User-facing `docs-site/**`: `bun run docs:test`, then `bun run docs:build`
- Config/schema/defaults: `bun run test:config`
  - If defaults or templates changed, also run `bun run generate:config-example` and `bun run verify:config-example`.
- General TypeScript source: `bun run typecheck`, then `bun run test:fast`
- Launcher or mpv plugin: `bun run test:launcher` or `bun run test:env`, based on the behavior changed
- Runtime compatibility or dist-sensitive wiring: `bun run test:runtime:compat`
- Stats dashboard: `bun run test:stats`
- Build/release scripts: `bun run test:scripts`

For substantial changes, use the full gate documented in `AGENTS.md` and `docs/workflow/verification.md`.

## Runtime Escalation

Real runtime checks are required when the claim depends on actual Electron, mpv, overlay, focus, window tracking, launch, or socket behavior. Run the relevant application flow when the environment supports it. Otherwise, report the missing runtime dependency and do not present cheaper checks as authoritative runtime validation.

## Pre-Handoff Checks

Before handoff, reconcile both questions:

1. Do behavior, defaults, flags, shortcuts, ports, APIs, architecture, or workflow changes require documentation updates?
2. Does the change require a current-outcome fragment under `changes/` according to `changes/README.md`?

Complete required updates before handoff or report the blocker.
