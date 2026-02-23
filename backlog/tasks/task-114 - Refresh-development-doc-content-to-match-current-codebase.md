---
id: TASK-114
title: Refresh development doc content to match current codebase
status: Done
assignee:
  - codex-development-docs-review
created_date: '2026-02-23 03:46'
updated_date: '2026-02-23 03:49'
labels:
  - docs
  - developer-experience
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User requested a thorough review of the current codebase and `docs/development.md`, then updates so the page reflects current scripts, architecture, and contributor workflow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `docs/development.md` commands/tooling sections match current `package.json`, `Makefile`, and active CI lanes.
- [x] #2 Runtime/composition guidance reflects current module ownership and does not reference removed guardrail steps.
- [x] #3 Updated doc builds cleanly via `bun run docs:build`.
<!-- AC:END -->

## Notes

- Reviewed `docs/development.md` against current `package.json` scripts, `Makefile` targets, launcher/runtime env overrides, and `.github/workflows/ci.yml` gate order.
- Updated setup instructions to use Bun for `vendor/texthooker-ui` and include submodule initialization.
- Added CI-equivalent local test/build/docs command sequence and clarified that subtitle dist tests are currently placeholders.
- Expanded environment-variable coverage (`SUBMINER_ROFI_THEME`, Jimaku/Jellyfin overrides, `SUBMINER_LOG_LEVEL`, `SUBMINER_MPV_LOG`, macOS helper skip flag).
- Verification: `bun run docs:build` passed.
