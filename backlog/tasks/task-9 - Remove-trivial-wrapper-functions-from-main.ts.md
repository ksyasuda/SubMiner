---
id: TASK-9
title: Remove trivial wrapper functions from main.ts
status: To Do
assignee: []
created_date: '2026-02-11 08:21'
labels:
  - refactor
  - main
  - simplicity
milestone: Codebase Clarity & Composability
dependencies:
  - TASK-7
references:
  - src/main.ts
priority: medium
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
