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

## Policy

- If file exceeds budget, prefer extracting domain module(s) first.
- Keep composition/orchestration files focused on wiring.
- Do not hand-edit generated artifacts; refactor source modules.
