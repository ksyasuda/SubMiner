---
id: TASK-1.5
title: 'Phase 5: Add critical behavior tests for untested services'
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:46'
updated_date: '2026-02-11 03:35'
labels: []
dependencies:
  - TASK-1.4
references:
  - plan.md
  - src/core/services/mpv-runtime-service.ts
  - src/core/services/subsync-runtime-service.ts
  - src/core/services/tokenizer-service.ts
  - src/core/services/cli-command-service.ts
parent_task_id: TASK-1
ordinal: 4000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add meaningful behavior tests for high-risk services called out in plan.md: mpv, subsync, tokenizer, and expanded CLI command coverage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `mpv` service has focused tests for protocol parsing, event dispatch, request/response matching, reconnection, and subtitle extraction behavior.
- [x] #2 `subsync` service has focused tests for engine path resolution, command construction, timeout/error handling, and result parsing.
- [x] #3 `tokenizer` service has focused tests for parser readiness, token extraction, fallback behavior, and edge-case inputs.
- [x] #4 CLI command service tests cover all dispatch paths, async error propagation, and second-instance forwarding behavior.
- [x] #5 `pnpm run test:core` passes with all new tests green.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused tests for `tokenizer-service.ts` behavior (normalization, Yomitan-unavailable fallback, mecab fallback success/error paths, empty input handling).
2. Add focused tests for `subsync-service.ts` command/engine selection and failure handling using mocked command utilities where feasible.
3. Add focused tests for `mpv-service.ts` protocol handling (line parsing, request-response routing, property-change dispatch) with lightweight socket stubs.
4. Expand `cli-command-service.ts` tests for dispatch/error/second-instance forwarding edge paths not currently covered.
5. Run `pnpm run test:core` iteratively and update acceptance criteria as each service reaches meaningful coverage.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added new tokenizer behavior tests in `src/core/services/tokenizer-service.test.ts` covering empty normalized input, newline normalization, mecab fallback success, and mecab error fallback-to-null.

Added new mpv protocol tests in `src/core/services/mpv-service.test.ts` covering JSON line-buffer parsing, property-change subtitle dispatch behavior, and request/response resolution by `request_id`.

Added new subsync workflow tests in `src/core/services/subsync-service.test.ts` covering already-running guard, manual-mode picker flow, and error propagation to OSD when MPV is unavailable.

Expanded `src/core/services/cli-command-service.test.ts` to cover socket/start dispatch, texthooker port override warning path, help-without-window shutdown, and async trigger-subsync error reporting.

Updated `package.json` `test:core` to include new/renamed test files; verification remains green with `pnpm run test:core` (17 tests total).

Expanded `mpv-service` tests with request rejection when disconnected, `requestProperty` error propagation, and pending-request disconnect resolution behavior.

Expanded `subsync-service` tests with manual alass source-track validation and auto-mode executable-path failure handling while ensuring in-progress state cleanup.

All updated tests remain green via `pnpm run test:core` after these additions.

Added Yomitan parser token-extraction coverage in `tokenizer-service.test.ts` (parser-available success path) in addition to fallback/edge-case tests.

Added MPV reconnection/request robustness tests (`scheduleReconnect`, disconnected request rejection, pending request disconnect resolution) to complement protocol/event/request-id tests in `mpv-service.test.ts`.

Added subsync command-construction tests using executable stubs for both engines (`ffsubsync` and `alass`) and validated success/failure result behavior; added timeout behavior coverage in `src/subsync/utils.test.ts` for child-process timeout handling used by subsync.

Expanded CLI dispatch tests with broad branch coverage for visibility/settings/copy/multi-copy/mining/open-runtime-options/stop/help/second-instance behaviors and async error propagation.

Verification: `pnpm run test:core` passes with 18 green tests including newly added `dist/subsync/utils.test.js`.
<!-- SECTION:NOTES:END -->
