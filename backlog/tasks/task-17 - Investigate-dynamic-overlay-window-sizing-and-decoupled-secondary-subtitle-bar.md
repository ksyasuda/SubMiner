---
id: TASK-17
title: Investigate dynamic overlay window sizing and decoupled secondary subtitle bar
status: Done
assignee:
  - codex
created_date: '2026-02-12 02:27'
updated_date: '2026-02-12 02:56'
labels:
  - overlay
  - ux
  - investigation
dependencies: []
documentation:
  - docs/overlay-window-sizing-investigation.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current visible and invisible overlays can both be enabled at once, but both windows occupy full mpv bounds so hover/input targets conflict. Investigate feasibility of sizing each overlay window to subtitle content bounds (including secondary subtitles) instead of fullscreen. Also investigate decoupling secondary subtitle rendering from overlay visibility by introducing a dedicated top bar region (full-width or text-width) that remains at top and respects secondary subtitle display mode config (hover/always/never).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Document current overlap/input conflict behavior when both overlays are enabled.
- [x] #2 Prototype or design approach for content-bounded overlay window sizing for visible and invisible overlays.
- [x] #3 Evaluate interaction model and technical constraints for a dedicated top secondary-subtitle bar independent from overlay visibility.
- [x] #4 Define implementation plan with tradeoffs, edge cases (line wrapping, long lines, resize, multi-monitor, mpv style sync), and recommended path forward.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Baseline current behavior
- Trace current visible/invisible overlay window creation and sizing behavior in main process and renderer paths.
- Document why enabling both overlays causes overlap and input/hover conflict.

2. Design content-bounded overlay sizing
- Evaluate measurement sources (renderer DOM bounds vs main process approximations).
- Define sizing algorithm (content bounds + padding + min/max + screen clamping) and update triggers.
- Identify required IPC contracts and ownership boundaries.

3. Design decoupled top secondary-subtitle bar
- Define a dedicated top region/window independent from primary subtitle visibility.
- Specify behavior for secondary subtitle display modes (hover/always/never) and pointer-event policy.

4. Recommend phased implementation
- Produce concrete file-level rollout steps, risks, edge cases (wrapping, long lines, resize, multi-monitor, style sync), and validation approach.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed code-path investigation across overlay window creation/sizing, visibility gating, IPC mouse passthrough, and renderer interaction layers.

Documented current conflict mechanism and architecture limitation: shared fullscreen bounds ownership across both overlay windows blocks safe simultaneous activation.

Proposed renderer-driven content-bounds IPC model, per-window bounds ownership, and a dedicated secondary top-bar window with mode-specific behavior.

Included phased implementation plan with risks and edge-case handling in docs/overlay-window-sizing-investigation.md.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed investigation and design for dynamic overlay window sizing and decoupled secondary subtitle rendering.

What changed:
- Added `docs/overlay-window-sizing-investigation.md` with a code-referenced analysis of current overlay behavior.
- Documented why overlap/input conflicts emerge under simultaneous overlay activation: both windows are full-size and currently share bounds ownership semantics.
- Designed a content-bounded sizing approach based on renderer-measured DOM bounds reported to main via IPC, with clamping, jitter guards, and fallback behavior.
- Evaluated top-bar decoupling options and recommended a dedicated secondary overlay window for independent lifecycle, z-order, and pointer policy.
- Defined a phased file-level implementation path and explicit edge-case handling (wrapping/long lines, resize churn, multi-monitor transitions, mpv style/render sync).

Validation:
- Investigation validated by tracing current runtime/services/renderer code paths; no runtime behavior changes were applied in this task.
- No automated tests were run because this task produced design documentation only (no executable code changes).
<!-- SECTION:FINAL_SUMMARY:END -->
