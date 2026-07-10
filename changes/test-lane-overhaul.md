type: internal
area: testing

- Test lanes now live in `scripts/test-lanes.ts`, discover tests by directory, and run each Bun test file in an isolated process with a wall timeout.
- CI now covers the previously orphaned stats, scripts, plugin process-retry, and runtime-compat suites. The SubMiner change-verification workflow also gained a stats lane.
