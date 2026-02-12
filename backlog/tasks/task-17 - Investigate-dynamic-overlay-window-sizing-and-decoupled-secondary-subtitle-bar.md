---
id: TASK-17
title: Investigate dynamic overlay window sizing and decoupled secondary subtitle bar
status: To Do
assignee: []
created_date: '2026-02-12 02:27'
labels:
  - overlay
  - ux
  - investigation
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current visible and invisible overlays can both be enabled at once, but both windows occupy full mpv bounds so hover/input targets conflict. Investigate feasibility of sizing each overlay window to subtitle content bounds (including secondary subtitles) instead of fullscreen. Also investigate decoupling secondary subtitle rendering from overlay visibility by introducing a dedicated top bar region (full-width or text-width) that remains at top and respects secondary subtitle display mode config (hover/always/never).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Document current overlap/input conflict behavior when both overlays are enabled.
- [ ] #2 Prototype or design approach for content-bounded overlay window sizing for visible and invisible overlays.
- [ ] #3 Evaluate interaction model and technical constraints for a dedicated top secondary-subtitle bar independent from overlay visibility.
- [ ] #4 Define implementation plan with tradeoffs, edge cases (line wrapping, long lines, resize, multi-monitor, mpv style sync), and recommended path forward.
<!-- AC:END -->
