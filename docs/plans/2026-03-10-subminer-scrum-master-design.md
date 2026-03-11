# SubMiner Scrum Master Skill Design

**Date:** 2026-03-10
**Status:** Approved

## Goal

Create a repo-local skill that can take incoming requests, bugs, or issues in the SubMiner repo, decide whether backlog tracking is warranted, create or update backlog work when appropriate, plan the implementation, dispatch one or more subagents, and ensure verification happens before handoff.

## Skill Contract

- **Name:** `subminer-scrum-master`
- **Location:** `.agents/skills/subminer-scrum-master/`
- **Use when:** the user gives a feature request, bug report, issue, refactor, or implementation ask in the SubMiner repo and the agent should own intake, planning, backlog hygiene, dispatch, and completion flow.
- **Responsibilities:**
  - assess whether backlog tracking is warranted
  - if needed, search/update/create proper backlog structure
  - write a plan before dispatching coding work
  - choose sequential vs parallel execution
  - assign explicit ownership to workers
  - require verification before final handoff
- **Limits:**
  - not the default code implementer unless delegation would be wasteful
  - no overlapping parallel write scopes
  - no skipping planning before dispatch
  - no skipping verification
  - must pause for ambiguous, risky, or external side-effect work

## Backlog Decision Rules

Backlog use is conditional, not mandatory.

- **Skip backlog when:**
  - question only
  - obvious mechanical edit
  - tiny isolated change with no real planning
- **Use backlog when:**
  - implementation requires planning
  - scope/approach needs decisions
  - multiple phases or subsystems
  - likely subagent dispatch
  - work should remain traceable

When backlog is used:
- search first
- update existing matching work when appropriate
- otherwise create standalone task or parent task
- use parent + subtasks for multi-part work
- record the implementation plan before coding starts

## Orchestration Policy

The skill orchestrates; workers implement.

- **No dispatch** for trivial/mechanical work
- **Single worker** for focused single-scope work
- **Parallel workers** only for clearly disjoint scopes
- **Sequential flow** for shared files, runtime coupling, or unclear boundaries

Every worker should receive:
- owned files/modules
- explicit reminder not to revert unrelated edits
- requirement to report changed files, tests run, and blockers

## Verification Policy

Every nontrivial code task gets verification.

- prefer `subminer-change-verification`
- use cheap-first lanes
- escalate only when needed
- accept worker-run verification only if it is clearly relevant and sufficient
- run a consolidating final verification pass when the scrum master needs stronger evidence

## Representative Flows

### Trivial fix, no backlog

1. assess request as mechanical or narrowly reversible
2. skip backlog
3. keep a short internal plan
4. implement directly or use one worker if helpful
5. run targeted verification
6. report concise summary

### Single-task implementation

1. search backlog
2. create/update one task
3. record plan
4. dispatch one worker
5. integrate result
6. run verification
7. update task and report outcome

### Multi-part feature

1. search backlog
2. create/update parent task
3. create subtasks for distinct deliverables/phases
4. record sequencing in the plan
5. dispatch workers only where write scopes do not overlap
6. integrate
7. run consolidated verification
8. update task state and report outcome

## V1 Scope

- instruction-heavy `SKILL.md`
- no helper scripts unless orchestration becomes too repetitive
- strong coordination with existing Backlog workflow and `subminer-change-verification`
