# File Size Budgets

Purpose: keep large modules from becoming maintenance bottlenecks.

## Current Budget

- TypeScript source files in `src/` and `launcher/`
- Soft budget: `500` LOC
- Excludes generated bundle artifacts (for example `subminer`)

## Commands

- Warning mode (non-blocking): `bun run check:file-budgets`
- Strict mode (CI/local gate): `bun run check:file-budgets:strict`
- Custom limit: `bun run scripts/check-file-budgets.ts --limit 650`
- Main runtime fan-in (warning): `bun run check:main-fanin`
- Main runtime fan-in (strict): `bun run check:main-fanin:strict`
- Default main fan-in thresholds: `import lines <= 110`, `unique runtime paths <= 11`

## Policy

- If file exceeds budget, prefer extracting domain module(s) first.
- Keep composition/orchestration files focused on wiring.
- Do not hand-edit generated artifacts; refactor source modules.
- Keep `src/main.ts` runtime import fan-in low by routing domain imports through `src/main/runtime/domains/*` and `src/main/runtime/registry.ts`.
