---
id: TASK-238.8
title: Refactor src/main.ts composition root into domain runtimes
status: In Progress
assignee: []
created_date: '2026-03-31 06:28'
updated_date: '2026-04-01 07:07'
labels:
  - tech-debt
  - runtime
  - maintainability
  - composition-root
dependencies: []
references:
  - src/main.ts
  - src/main/boot/services
  - src/main/runtime/composers
  - docs/architecture/README.md
parent_task_id: TASK-238
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Refactor `src/main.ts` so it becomes a thin composition root and the domain-specific runtime wiring moves into short wrapper modules under `src/main/`. Preserve all current behavior, IPC contracts, and config/schema semantics while reducing the entrypoint to boot services, grouped runtime instantiation, startup execution, and process-level quit handling.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `src/main.ts` is bootstrap/composition only: platform preflight, boot services, runtime creation, startup execution, and top-level quit/signal handling.
- [ ] #2 `src/main.ts` no longer imports `src/main/runtime/*-main-deps.ts` directly.
- [ ] #3 `src/main.ts` has no local names like `build*MainDepsHandler`, `*MainDeps`, or trivial `*Handler` pass-through wrappers.
- [ ] #4 New wrapper files stay under ~500 LOC each; if one exceeds that, split before merge.
- [ ] #5 Cross-domain coordination stays in `main.ts`; wrapper modules stay acyclic and communicate via injected callbacks.
- [ ] #6 No user-facing behavior, config fields, or IPC channel names change.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
CI follow-up: typecheck failed after the runtime split because playlist-browser IPC deps were not threaded through the new bootstrap/composer surfaces. Wiring the missing open action and registration deps now.
<!-- SECTION:NOTES:END -->
