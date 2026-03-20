---
id: TASK-160
title: Create repo-local scrum master orchestration skill
status: Done
assignee:
  - codex
created_date: '2026-03-11 06:32'
updated_date: '2026-03-16 05:13'
labels:
  - skills
  - workflow
  - backlog
  - subagents
  - automation
dependencies: []
priority: high
ordinal: 28500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Design and implement a repo-local skill that can turn incoming requests/issues into well-scoped Backlog tasks, then coordinate one or more subagents to implement and verify the resulting work using the repo's established skills and verification workflow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Define the trigger surface, responsibilities, and limits of the scrum master skill for request intake, backlog updates, and execution orchestration.
- [x] #2 Implement the skill package with concise guidance for intake, task search/create/update, planning, and subagent dispatch.
- [x] #3 Specify how the skill coordinates with existing repo workflows, including Backlog.md requirements and SubMiner change verification.
- [x] #4 Verify the skill is usable in this repo context for at least one representative request flow.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a design doc at `docs/plans/2026-03-10-subminer-scrum-master-design.md` capturing the approved contract, backlog decision rules, orchestration policy, and verification handoff.
2. Create the repo-local skill at `.agents/skills/subminer-scrum-master/` with a concise `SKILL.md`.
3. Keep v1 instruction-heavy: backlog-or-not decision first, search/create/update task structure when needed, require planning before dispatch, use parent + subtasks for multi-part work, dispatch conservatively with explicit ownership, and require `subminer-change-verification` before handoff.
4. Include representative flows for trivial no-ticket work, single-task implementation, and multi-part parent+subtask execution.
5. Verify the skill in-repo with one representative request flow and update `TASK-160` with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
V1 will be instruction-heavy rather than script-heavy. The skill will decide whether backlog is warranted, record plans before dispatch, create parent+subtasks when needed, dispatch conservatively with explicit ownership, and require `subminer-change-verification` before handoff.

2026-03-10: Added design doc at docs/plans/2026-03-10-subminer-scrum-master-design.md and created the repo-local skill at .agents/skills/subminer-scrum-master/.

2026-03-10: Verified the skill against a representative request flow ('Fix launcher auto-start pause bug and make sure it is verified') by checking that the instructions cover backlog decisioning, planning-before-dispatch, single-task execution, explicit worker ownership, and verification via subminer-change-verification.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented a repo-local `subminer-scrum-master` skill that assesses whether backlog tracking is needed, manages task structure and planning when appropriate, dispatches subagents conservatively, and requires verification before handoff. Verified the skill in-repo against a representative launcher bug workflow.
<!-- SECTION:FINAL_SUMMARY:END -->
