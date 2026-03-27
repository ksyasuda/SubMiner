---
id: TASK-238.7
title: Split src/main.ts into boot-phase services, runtimes, and handlers
status: To Do
assignee: []
created_date: '2026-03-27 00:00'
labels:
  - tech-debt
  - runtime
  - maintainability
  - architecture
milestone: m-0
dependencies:
  - TASK-238.6
references:
  - src/main.ts
  - src/main/boot/services.ts
  - src/main/boot/runtimes.ts
  - src/main/boot/handlers.ts
  - src/main/runtime/composers
parent_task_id: TASK-238
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After the remaining inline runtime logic and composer gaps are extracted, `src/main.ts` should be split along boot-phase boundaries so the entrypoint stops mixing service construction, domain runtime composition, and handler wiring in one file. This task tracks that structural split: move service instantiation, runtime composition, and handler orchestration into dedicated boot modules, then leave `src/main.ts` as a thin lifecycle coordinator with clear startup-path selection.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Service instantiation lives in a dedicated boot module instead of a large inline setup block in `src/main.ts`.
- [ ] #2 Domain runtime composition lives in a dedicated boot module, separate from lifecycle and handler dispatch.
- [ ] #3 Handler/composer invocation lives in a dedicated boot module, with `src/main.ts` reduced to app lifecycle and startup-path selection.
- [ ] #4 Existing startup behavior remains unchanged across desktop and headless flows.
- [ ] #5 Focused tests cover the split surfaces, and the relevant runtime/typecheck gate passes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Recommended sequence:

1. Re-scan `src/main.ts` after TASK-238.6 lands and mark the remaining boot-phase seams by responsibility.
2. Extract service instantiation into `src/main/boot/services.ts` or equivalent.
3. Extract runtime composition into `src/main/boot/runtimes.ts` or equivalent.
4. Extract handler/composer orchestration into `src/main/boot/handlers.ts` or equivalent.
5. Shrink `src/main.ts` to startup-path selection, app lifecycle hooks, and minimal boot wiring.
6. Verify the split with focused entrypoint/runtime tests first, then run the broader runtime gate if the refactor crosses startup boundaries.

Guardrails:

- Keep the split behavior-preserving.
- Prefer small boot modules with narrow ownership over a new monolithic bootstrap layer.
- Do not reopen the inline logic work already tracked by TASK-238.6 unless a remaining seam truly belongs here.
<!-- SECTION:PLAN:END -->
