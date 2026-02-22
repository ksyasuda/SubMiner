---
id: TASK-101
title: Consolidate architecture docs and archive task noise
status: Done
assignee:
  - codex-task101-docs-archive
created_date: '2026-02-21 07:15'
updated_date: '2026-02-22 03:01'
labels:
  - documentation
  - maintainability
dependencies:
  - TASK-96
  - TASK-97
  - TASK-98
  - TASK-99
  - TASK-100
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Architecture guidance is fragmented across long-lived docs and task notes. Consolidate canonical runtime architecture docs and move stale/noisy task-level detail to archive.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Audit architecture and development docs for duplicated runtime composition guidance.
2. Select one canonical architecture section for runtime composition and dependency boundaries.
3. Refactor duplicate sections to brief pointers to the canonical section.
4. Move stale task-evidence prose from persistent docs to backlog/archive where appropriate.
5. Add a compact architecture diagram and update links in contributor docs.
6. Validate docs build and link integrity.
7. Record what moved and why in task notes for traceability.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Runtime composition guidance exists in one canonical location.
- [x] #2 Duplicated/stale architecture notes removed from long-lived docs.
- [x] #3 Task evidence retained in backlog/archive, not lost.
- [x] #4 Docs build and links pass after consolidation.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Canonicalize runtime architecture guidance in `docs/architecture.md` and remove task-ID provenance wording so architecture guidance remains timeless.
2) Trim duplicated runtime architecture bullets from `docs/development.md` Contributor Notes and point contributors to `/architecture` as canonical guidance.
3) Convert `docs/structure-roadmap.md` from a full task-history roadmap into a short archival note that points to canonical architecture docs and states historical detail is retained in backlog task records.
4) Record a changelog of moved/removed sections in TASK-101 notes/final summary to preserve evidence that was removed from long-lived docs.
5) Run `bun run docs:build` and finalize AC/DoD checks in Backlog once docs build/link validation passes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Plan captured in docs/plans/2026-02-22-task-101-architecture-doc-consolidation.md before edits.

Initial removal/move inventory: (a) TASK-tagged provenance bullets in docs/architecture.md Why This Design, (b) duplicated runtime-composer/domain-ownership bullets in docs/development.md Contributor Notes, (c) TASK-27 split-sequence/migration-risk body in docs/structure-roadmap.md.

Doc consolidation changelog (pass 1):
- docs/architecture.md: removed task-ID provenance bullets from `Why This Design` and replaced with timeless architecture wording.
- docs/development.md: removed duplicated runtime-composer/domain-registry/MPV-split architecture bullets from `Contributor Notes`; added single canonical pointer to `/architecture`.
- docs/structure-roadmap.md: replaced full TASK-27 roadmap/migration-risk/split-sequence body with archival notice that redirects to canonical architecture docs and backlog historical records.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Consolidated runtime architecture guidance into canonical long-lived docs and removed task-history noise:
- `docs/architecture.md`: removed task-ID provenance bullets from `Why This Design` and kept architecture rationale timeless.
- `docs/development.md`: trimmed duplicated architecture implementation bullets in `Contributor Notes`; retained a direct canonical link to `/architecture`.
- `docs/structure-roadmap.md`: replaced full TASK-27 roadmap body with archival notice pointing to canonical architecture guidance.

Task-noise evidence is preserved in Backlog records (TASK-101 implementation notes/final summary plus existing backlog archive/task history) rather than long-lived docs. Verification: `bun run docs:build` passed (VitePress build complete, no link errors).
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Change log of moved/removed doc sections included in task notes.
- [x] #2 `bun run docs:build` passes.
- [x] #3 Contributor-facing entry points link to canonical architecture section.
<!-- DOD:END -->
