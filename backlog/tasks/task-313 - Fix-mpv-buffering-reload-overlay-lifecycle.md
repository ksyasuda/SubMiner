---
id: TASK-313
title: Fix mpv buffering reload overlay lifecycle
status: To Do
assignee: []
created_date: '2026-05-02 22:12'
labels:
  - bug
  - mpv
  - overlay
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
macOS local playback can emit an mpv reload/end-file/file-loaded sequence during buffering. SubMiner should treat same-media reload churn as a continuation, not a fresh playback session, so the visible overlay remains available and startup-only tokenization/AniSkip work is not repeated unnecessarily.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Same-media mpv reload buffering does not hide the visible overlay.
- [ ] #2 Same-media mpv reload buffering does not re-arm the pause-until-ready startup gate or wait for a second tokenization-ready signal.
- [ ] #3 Same-media mpv reload buffering does not repeat AniSkip lookup work for the already-loaded media.
- [ ] #4 Normal new-file playback still clears per-media state, applies managed subtitle defaults, auto-starts/updates the overlay, and runs needed startup checks.
- [ ] #5 Regression coverage exercises the buffering reload/end-file/file-loaded sequence in the mpv plugin lifecycle.
<!-- AC:END -->
