---
id: TASK-20
title: Implement content-bounded overlay windows and decoupled secondary top bar
status: To Do
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-12 09:42'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement the overlay sizing redesign documented in `overlay_window.md`: move visible/invisible overlays from fullscreen bounds to content-bounded sizing, and decouple secondary subtitle rendering into an independent top bar window/lifecycle.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Per-layer bounds ownership is implemented for overlay windows (no shared full-bounds setter for all layers).
- [ ] #2 Renderer-to-main IPC contract exists for measured overlay content bounds with layer identity and safe validation.
- [ ] #3 Visible and invisible overlays use content-bounded sizing with padding/clamp/jitter protections and full-bounds fallback when measurements are unavailable.
- [ ] #4 Secondary subtitle top bar is decoupled from primary overlay visibility and follows mode-specific behavior.
- [ ] #5 Automated tests and manual validation matrix cover wrapping, style changes, monitor moves, tracker churn, and simultaneous overlay states.
<!-- AC:END -->
