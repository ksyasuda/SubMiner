---
id: TASK-33
title: Restrict mpv socket connection logs to debug mode in Electron
status: Done
assignee: []
created_date: '2026-02-13 19:39'
updated_date: '2026-02-18 04:11'
labels:
  - electron
  - logging
  - mpv
  - frontend
  - quality
dependencies: []
priority: high
ordinal: 31000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

In normal operation, Electron should not spam logs while waiting for mpv socket connection. Emit mpv socket connection logs only when app logging level is debug. In regular use, keep connection attempts silent while waiting, and log exactly once when the socket connects successfully.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 When logging level is not debug, do not emit logs for mpv socket connection attempts, retries, or wait loops.
- [ ] #2 In non-debug mode, the app should silently wait for mpv socket readiness instead of printing connection-loop noise.
- [ ] #3 Log exactly one concise INFO log entry when a mpv socket connection succeeds (e.g., per lifecycle attempt/session).
- [ ] #4 In debug mode, keep existing detailed connection attempt/retry logs to aid diagnosis.
- [ ] #5 No functional change to socket connection/retry behavior besides logging level gating.
- [ ] #6 If connection fails and retries continue, keep a debug-only or final-error log policy consistent with existing logging severity conventions.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented debug-gated MPV socket noise suppression in normal mode while preserving reconnect behavior. Added an internal check for SUBMINER_LOG_LEVEL==='debug' and wrapped socket reconnect-attempt, close, and error logs so regular startup/reconnect loops no longer emit per-interval messages unless debug logging is enabled. Kept existing connection-success flow unchanged. Validation: `pnpm run build`, `node --test dist/core/services/mpv-service.test.js`.

<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 No connection-attempt/connect-wait logs are produced in non-debug mode.
- [ ] #2 Exactly one success log is produced when connection is established in non-debug mode.
- [ ] #3 Debug mode continues to emit detailed connection logs as before.
- [ ] #4 Behavior is validated across cold-start and reconnect attempts (single success log per successful connect event).
<!-- DOD:END -->
