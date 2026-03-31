---
id: TASK-256
title: Fix texthooker page live websocket connect/send regression
status: Done
assignee:
  - codex
created_date: '2026-03-30 06:04'
updated_date: '2026-03-31 19:37'
labels:
  - bug
  - texthooker
  - websocket
dependencies: []
priority: medium
ordinal: 160500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate why the bundled texthooker page loads at the local HTTP endpoint but does not reliably connect to the configured websocket feed or receive/display live subtitle lines. Identify the regression in the SubMiner startup/bootstrap or vendored texthooker client path, restore live line delivery, and cover the fix with focused regression tests and any required docs updates.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Bundled texthooker connects to the intended websocket endpoint on launch using the configured/default SubMiner startup path.
- [x] #2 Incoming subtitle or annotation websocket messages are accepted by the bundled texthooker and rendered as live lines.
- [x] #3 Regression coverage fails before the fix and passes after the fix for the identified breakage.
- [x] #4 Relevant docs/config notes are updated if user-facing behavior or troubleshooting guidance changes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused CLI regression test covering `--texthooker` startup when the runtime has a resolved websocket URL, proving the handler currently starts texthooker without that URL.
2. Extend CLI texthooker dependencies/runtime wiring so the handler can retrieve the resolved texthooker websocket URL from current config/runtime state.
3. Update the CLI texthooker flow to pass the resolved websocket URL into texthooker startup instead of starting the HTTP server with only a port.
4. Run focused tests for CLI command handling and texthooker bootstrap behavior; update task notes/final summary with the verified root cause and fix.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: the CLI `--texthooker` path started the HTTP texthooker server with only the port, so the served page never received `bannou-texthooker-websocketUrl` and fell back to the vendored default `ws://localhost:6677`. In environments where the regular websocket was skipped or the annotation websocket should have been used, the page stayed on `Connecting...` and never received lines.

Fix: added a shared `resolveTexthookerWebsocketUrl(...)` helper for websocket selection, reused it in both app-ready startup and CLI texthooker context wiring, and threaded the resolved websocket URL through `handleCliCommand` into `Texthooker.start(...)`.

Verification: `bun run typecheck`; focused Bun tests for texthooker bootstrap, startup, CLI command handling, and CLI context wiring; browser-level repro against a throwaway source-backed texthooker server confirmed the page bootstraps `ws://127.0.0.1:6678`, connects successfully, and renders live sample lines (`テスト一`, `テスト二`).

Docs: no user-facing behavior change beyond restoring the intended existing behavior, so no docs update was required.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored texthooker live line delivery for the CLI/startup path that launched the page without a resolved websocket URL. Shared websocket URL resolution between app-ready startup and CLI texthooker context, forwarded that URL into `Texthooker.start(...)`, added regression coverage for the CLI path, and verified both by focused tests and a browser-level throwaway server that connected on `ws://127.0.0.1:6678` and rendered live sample lines.
<!-- SECTION:FINAL_SUMMARY:END -->
