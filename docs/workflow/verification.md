<!-- read_when: choosing what tests/build steps to run before handoff -->

# Verification

Status: active  
Last verified: 2026-07-06  
Owner: Kyle Yasuda  
Read when: selecting the right verification lane for a change

## Lane Infrastructure

- Lane membership is defined once in `scripts/test-lanes.ts` and discovered by
  directory — new test files join their lane automatically; never hand-list test
  files in `package.json`.
- `scripts/run-test-lane.mjs` runs each test file in its own `bun test` process
  (per-file isolation with a wall timeout) so a hanging test or leaked global in
  one file cannot cascade into the rest of the lane. `--jobs N` parallelizes;
  `--single-process` restores the shared-process mode for debugging.
- `bun run test:fast` is the full source gate: discovered `src/**`, launcher
  unit, `scripts/**`, and the compiled runtime-compat slice.
- `.github/workflows/quality-gate.yml` is the reusable `workflow_call` gate for
  pull requests, stable tags, and prerelease tags. Keep common quality steps
  there instead of copying them into caller workflows.
- The reusable gate installs Lua and runs `bun run test:env`, so the shipped mpv
  plugin tests run for every pull request and tagged release.

## Default Handoff Gate

```bash
bun run typecheck
bun run test:fast
bun run test:env
bun run build
bun run test:smoke:dist
```

If `docs-site/` changed, also run:

```bash
bun run docs:test
bun run docs:build
```

## Cheap-First Lane Selection

- Docs-only boundary/content changes: `bun run docs:test`, `bun run docs:build`
- Internal KB / `AGENTS.md` changes: `bun run test:docs:kb`
- Config/schema/defaults: `bun run test:config`, then `bun run generate:config-example` if template/defaults changed
- Launcher/plugin: `bun run test:launcher` or `bun run test:env`
- Runtime-compat / compiled behavior: `bun run test:runtime:compat`
- Stats dashboard UI: `bun run test:stats`
- Build/release scripts (`scripts/**`): `bun run test:scripts`
- Coverage for the maintained source lane: `bun run test:coverage:src`
- Deep/local full gate: default handoff gate above

## Coverage Reporting

- `bun run test:coverage:src` runs the maintained `test:src` lane through a sharded coverage runner: one Bun coverage process per test file, then merged LCOV output.
- Machine-readable output lands at `coverage/test-src/lcov.info`.
- Every reusable quality-gate run uploads that LCOV file as the
  `coverage-test-src` artifact.

## Dependency Audit Policy

- `bun audit --audit-level high` blocks the reusable quality gate.
- Keep security overrides and dependency patches at the minimum fixed version.
  Remove them after the owning package ships and adopts a compatible fix.

## Rules

- Capture exact failing command and error when verification breaks.
- Prefer the cheapest sufficient lane first.
- Escalate when the change crosses boundaries or touches release-sensitive behavior.
- Never hand-edit `dist/launcher/subminer`; validate it through build/test flow instead.
