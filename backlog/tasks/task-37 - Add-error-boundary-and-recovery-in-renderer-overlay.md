---
id: TASK-37
title: Add error boundary and recovery in renderer overlay
status: To Do
assignee: []
created_date: '2026-02-14 01:01'
labels:
  - renderer
  - reliability
  - error-handling
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add a top-level error boundary in the renderer orchestrator that catches unhandled errors in modals and rendering logic, displays a user-friendly error message, and recovers without crashing the overlay.

## Motivation

If a renderer modal throws (e.g., jimaku API timeout, DOM manipulation error, malformed subtitle data), the entire overlay can become unresponsive. Since the overlay is transparent and positioned over mpv, a crashed renderer is invisible but blocks interaction.

## Scope

1. Wrap the renderer orchestrator's modal dispatch and subtitle rendering in try/catch boundaries
2. On error: log the error, dismiss the active modal, show a brief toast/notification in the overlay
3. Ensure the overlay returns to a usable state (subtitle display, click-through, shortcuts all work)
4. Add a global `window.onerror` / `unhandledrejection` handler as a last-resort safety net
5. Consider a "renderer health" heartbeat from main process that can force-reload the renderer if it becomes unresponsive

## Design constraints

- Error recovery must not disrupt mpv playback
- Toast notifications should auto-dismiss and not interfere with subtitle layout
- Errors should be logged with enough context for debugging (stack trace, active modal, last subtitle state)
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Unhandled errors in modal flows are caught and do not crash the overlay.
- [ ] #2 After an error, the overlay returns to a functional state (subtitles render, shortcuts work).
- [ ] #3 A brief toast/notification informs the user that an error occurred.
- [ ] #4 Global unhandledrejection and onerror handlers are registered as safety nets.
- [ ] #5 Error details are logged with context (stack trace, active modal, subtitle state).
- [ ] #6 mpv playback is never interrupted by renderer errors.
<!-- AC:END -->
