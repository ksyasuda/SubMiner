# Maintainability Guardrails

Purpose: keep runtime composition boundaries maintainable.

## Current Checks

- Main runtime fan-in (warning): `bun run check:main-fanin`
- Main runtime fan-in (strict): `bun run check:main-fanin:strict`
- Default main fan-in thresholds: `import lines <= 110`, `unique runtime paths <= 11`
- Runtime cycle check (warning): `bun run check:runtime-cycles`
- Runtime cycle check (strict): `bun run check:runtime-cycles:strict`

## Policy

- Keep composition/orchestration files focused on wiring.
- Do not hand-edit generated artifacts; refactor source modules.
- Keep `src/main.ts` runtime import fan-in low by routing domain imports through `src/main/runtime/domains/*` and `src/main/runtime/registry.ts`.
- Avoid runtime import cycles by extracting shared helpers into leaf modules.
