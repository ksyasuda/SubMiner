type: internal
area: testing

- Test lanes are now defined once in `scripts/test-lanes.ts` and discovered by directory instead of hand-maintained file lists in `package.json`; the unused `test:core:*`/`test:config:dist`/`test:full` scripts were removed.
- `scripts/run-test-lane.mjs` runs each test file in an isolated `bun test` process with a wall timeout, so a hanging test or leaked global can no longer cascade failures across the lane.
- Previously orphaned suites now run in CI: the stats dashboard tests (`bun run test:stats`), the `scripts/**` tests (`bun run test:scripts`, including the change-verification skill tests), the `test-plugin-process-start-retries.lua` plugin test, and the runtime-compat dist slice (now part of `bun run test:fast`).
- The change-verification skill gained a `stats` lane for `stats/` edits.
