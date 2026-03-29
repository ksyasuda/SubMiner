<!-- read_when: choosing what tests/build steps to run before handoff -->

# Verification

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: selecting the right verification lane for a change

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
- Coverage for the maintained source lane: `bun run test:coverage:src`
- Deep/local full gate: default handoff gate above

## Coverage Reporting

- `bun run test:coverage:src` runs the maintained `test:src` lane through a sharded coverage runner: one Bun coverage process per test file, then merged LCOV output.
- Machine-readable output lands at `coverage/test-src/lcov.info`.
- CI and release quality-gate runs upload that LCOV file as the `coverage-test-src` artifact.

## Rules

- Capture exact failing command and error when verification breaks.
- Prefer the cheapest sufficient lane first.
- Escalate when the change crosses boundaries or touches release-sensitive behavior.
- Never hand-edit `dist/launcher/subminer`; validate it through build/test flow instead.
