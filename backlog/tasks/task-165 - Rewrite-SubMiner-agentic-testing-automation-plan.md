---
id: TASK-165
title: Rewrite SubMiner agentic testing automation plan
status: Done
assignee: []
created_date: '2026-03-13 04:45'
updated_date: '2026-03-13 04:47'
labels:
  - planning
  - testing
  - agents
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/testing-plan.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-change-verification/SKILL.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/.agents/skills/subminer-scrum-master/SKILL.md
documentation:
  - /home/sudacode/projects/japanese/SubMiner/docs-site/development.md
  - /home/sudacode/projects/japanese/SubMiner/docs-site/architecture.md
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the current generic Electron/mpv testing plan with a SubMiner-specific plan that uses the existing skills as the source of truth, treats real launcher/plugin/mpv runtime verification as primary, and defines a non-interference contract for parallel agent work.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `testing-plan.md` is rewritten for SubMiner rather than a generic Electron+mpv app
- [x] #2 The plan keeps `subminer-scrum-master` and `subminer-change-verification` as the primary orchestration and verification entrypoints
- [x] #3 The plan defines real launcher/plugin/mpv runtime verification as the authoritative lane for runtime bug claims
- [x] #4 The plan defines explicit session isolation and non-interference rules for parallel agent work
- [x] #5 The plan defines artifact/reporting expectations and phased rollout, with synthetic/headless verification clearly secondary to real-runtime verification
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Review the existing testing plan and compare it against current SubMiner architecture, verification lanes, and skills.
2. Replace the generic Electron/mpv harness framing with a SubMiner-specific control plane centered on existing skills.
3. Define the authoritative real-runtime lane, session isolation rules, concurrency classes, and reporting contract.
4. Sanity-check the rewritten document against current repo docs and skill contracts before handoff.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Rewrote `testing-plan.md` around existing `subminer-scrum-master` and `subminer-change-verification` responsibilities instead of proposing a competing new top-level testing skill.

Set real launcher/plugin/mpv/runtime verification as the authoritative lane for runtime bug claims and made synthetic/headless verification explicitly secondary.

Defined session-scoped paths, unique mutable resources, concurrency classes, and an exclusive lease for conflicting real-runtime verification to prevent parallel interference.

Sanity-checked the final document by inspecting the rewritten file content and diff.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Rewrote `testing-plan.md` into a SubMiner-specific agentic verification plan. The new document keeps `subminer-scrum-master` and `subminer-change-verification` as the primary orchestration and verification entrypoints, treats the real launcher/plugin/mpv/runtime path as authoritative for runtime bug claims, and defines a hard non-interference contract for parallel work through session isolation and an exclusive real-runtime lease. The plan now also includes an explicit reporting schema, capture policy, phased rollout, and a clear statement that true parallel full-app instances are not a phase-1 requirement. Verification for this task was a document sanity pass against the current repo docs, skills, and the resulting file diff.
<!-- SECTION:FINAL_SUMMARY:END -->
