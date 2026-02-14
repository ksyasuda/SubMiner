---
id: TASK-27.6
title: Add structural quality gates for file size and complexity
status: To Do
assignee:
  - architect
created_date: '2026-02-13 17:13'
updated_date: '2026-02-13 21:19'
labels:
  - 'owner:architect'
  - 'owner:backend'
  - 'owner:frontend'
dependencies:
  - TASK-27.1
  - TASK-27.2
  - TASK-27.3
  - TASK-27.4
  - TASK-27.5
parent_task_id: TASK-27
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add automated safeguards so oversized/complex files are caught early and refactor progress is measurable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Extend check-main-lines gate script to accept any file path and apply it to: src/main.ts, src/anki-integration.ts (or src/anki-integration/index.ts), src/core/services/mpv-service.ts, src/renderer/positioning.ts, src/config/service.ts.
- [ ] #2 Define per-file thresholds (suggested: 400 LOC default, 600 for config/service.ts, justified exceptions documented in the script).
- [ ] #3 Add ESLint complexity rule (or lightweight proxy) with per-directory thresholds — at minimum for src/core/services/ and src/anki-integration/.
- [ ] #4 Create a clear exception process for justified threshold breaks: comment in code with expiration date and owner.
- [ ] #5 Document thresholds in docs/structure-roadmap.md.
- [ ] #6 Clarify enforcement: local-only (npm script) or CI-enforced. If CI, add to the CI pipeline config.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Review Additions

Original task omitted anki-integration.ts from the gated file list — it's the largest file at 2,679 LOC and the primary target of TASK-27.3. Added to AC#1.

The existing check-main-lines.sh is a simple `wc -l` check. Consider augmenting with:
- ESLint `complexity` rule for cyclomatic complexity
- Method count per file (proxy for cohesion)
- Import count per file (proxy for coupling)

Raw line count is better than nothing but doesn't catch files that are long because of well-structured, low-complexity code (like config/definitions.ts at 479 LOC which is just defaults).
<!-- SECTION:NOTES:END -->
