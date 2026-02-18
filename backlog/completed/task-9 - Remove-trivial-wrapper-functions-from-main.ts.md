---
id: TASK-9
title: Remove trivial wrapper functions from main.ts
status: Done
assignee: []
created_date: '2026-02-11 08:21'
updated_date: '2026-02-15 07:00'
labels:
  - refactor
  - main
  - simplicity
milestone: Codebase Clarity & Composability
dependencies:
  - TASK-7
references:
  - src/main.ts
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

main.ts contains many trivial single-line wrapper functions that add indirection without value:

```typescript
function getOverlayWindows(): BrowserWindow[] {
  return overlayManager.getOverlayWindows();
}
function updateOverlayBounds(geometry: WindowGeometry): void {
  updateOverlayBoundsService(geometry, () => getOverlayWindows());
}
function ensureOverlayWindowLevel(window: BrowserWindow): void {
  ensureOverlayWindowLevelService(window);
}
```

Similarly, config accessor wrappers like `getJimakuLanguagePreference()`, `getJimakuMaxEntryResults()`, `resolveJimakuApiKey()` are pure boilerplate.

After TASK-7 (AppState container), many of these can be eliminated by having services access the state container directly, or by using the service functions directly at call sites without local wrappers.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Trivial pass-through wrappers eliminated (call service/manager directly)
- [ ] #2 Config accessor wrappers replaced with direct calls or a config accessor helper
- [ ] #3 main.ts line count reduced
- [ ] #4 No functional changes
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Priority changed from medium to low: this work is largely subsumed by TASK-27.2 (split main.ts). When main.ts is decomposed into composition-root modules, trivial wrappers will naturally be eliminated or inlined. Recommend folding remaining wrapper cleanup into TASK-27.2 rather than tracking separately. Keep this ticket as a checklist reference but don't execute independently.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Subsumed by TASK-27.2 (main.ts split). Trivial wrappers were eliminated or inlined as composition-root modules were extracted. main.ts reduced from ~2000+ LOC to 1384 with state routed through appState container. Standalone wrapper removal pass no longer needed.

<!-- SECTION:FINAL_SUMMARY:END -->
