---
id: TASK-238
title: Codebase health follow-up: decompose remaining oversized runtime surfaces
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - tech-debt
  - maintainability
  - runtime
milestone: m-0
dependencies: []
references:
  - src/main.ts
  - src/types.ts
  - src/main/character-dictionary-runtime.ts
  - src/core/services/immersion-tracker/query.ts
  - backlog/tasks/task-87 - Codebase-health-harden-verification-and-retire-dead-architecture-identified-in-the-March-2026-review.md
  - backlog/completed/task-87.4 - Runtime-composition-root-remove-dead-symbols-and-tighten-module-boundaries-in-src-main.ts.md
  - backlog/completed/task-87.6 - Anki-integration-maintainability-continue-decomposing-the-oversized-orchestration-layer.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow up the March 2026 codebase-health work with a narrower pass over the biggest remaining production hotspots. The latest review correctly flags `src/main.ts` and `src/types.ts` as maintainability pressure, but it also misses the next real large surfaces that will keep slowing future work: `src/main/character-dictionary-runtime.ts` and `src/core/services/immersion-tracker/query.ts`. This parent task should track focused decomposition work that preserves behavior, avoids redoing already-completed dead-architecture cleanup, and keeps each slice small enough for isolated implementation.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Child tasks exist for each focused cleanup slice instead of one broad “split the monoliths” effort.
- [ ] #2 The parent task records sequencing so agents do not overlap on `src/main.ts` and other shared surfaces.
- [ ] #3 The selected follow-up tasks target still-live pressure points, not already-completed work like TASK-87.4, TASK-87.5, or TASK-87.6.
- [ ] #4 Completion of the child tasks leaves runtime wiring, shared types, character-dictionary orchestration, and immersion-tracker queries materially easier to review and extend.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Recommended sequencing:

1. Start TASK-238.3 first. A compatibility-first type split reduces churn risk for the later runtime/query refactors.
2. Run TASK-238.4 and TASK-238.5 in parallel after TASK-238.3 if desired; they touch different domains.
3. Run TASK-238.1 after or alongside the domain refactors, but keep it focused on window/bootstrap composition only.
4. Run TASK-238.2 after TASK-238.1 because both touch `src/main.ts` and the CLI/headless flow should build on the cleaner composition root.

Shared guardrails:

- Do not reopen already-completed dead-module cleanup from TASK-87.5 unless new evidence appears.
- Keep `src/types.ts` migration compatibility-first; avoid a repo-wide import churn bomb.
- Prefer extracting named runtime/domain modules over moving code into new giant helper files.
- Verify each slice with the cheapest sufficient lane, then escalate when a task crosses runtime/build boundaries.
<!-- SECTION:PLAN:END -->
