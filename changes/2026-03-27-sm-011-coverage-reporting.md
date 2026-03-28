type: internal
area: release

- Added a maintained source coverage lane that shards Bun coverage one test file at a time and merges LCOV output into `coverage/test-src/lcov.info`.
- CI and release quality-gate now upload the merged source-lane LCOV artifact for inspection.
