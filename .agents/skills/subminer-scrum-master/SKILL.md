---
name: "subminer-scrum-master"
description: "Use in the SubMiner repo when a request should be turned into planned work and driven through execution. Assesses whether backlog tracking is warranted, creates or updates tasks when needed, records a plan, dispatches one or more subagents, and requires verification before handoff."
---

# SubMiner Scrum Master

Own workflow, not code by default.

Use this skill when the user gives a feature request, bug report, issue, refactor, or implementation ask and the agent should manage intake, planning, backlog hygiene, worker dispatch, and verification through completion.

## Core Rules

1. Decide first whether backlog tracking is warranted.
2. If backlog is needed, search first. Update existing work when it clearly matches.
3. If backlog is not needed, keep the process light. Do not invent ticket ceremony.
4. Record a plan before dispatching coding work.
5. Use parent + subtasks for multi-part work when backlog is used.
6. Dispatch conservatively. Parallelize only disjoint write scopes.
7. Require verification before handoff, typically via `subminer-change-verification`.
8. Report backlog actions, dispatched workers, verification, blockers, and remaining risks.

## Backlog Decision

Skip backlog when the request is:
- question only
- obvious mechanical edit
- tiny isolated change with no real planning

Use backlog when the work:
- needs planning or scope decisions
- spans multiple phases or subsystems
- is likely to need subagent dispatch
- should remain traceable for handoff/resume

If backlog is used:
- search existing tasks first
- create/update a standalone task for one focused deliverable
- create/update a parent task plus subtasks for multi-part work
- record the implementation plan in the task before implementation begins

## Intake Workflow

1. Parse the request.
   Classify it as question, mechanical edit, bugfix, feature, refactor, investigation, or follow-up.
2. Decide whether backlog is needed.
3. If backlog is needed:
   - search first
   - update existing task if clearly relevant
   - otherwise create the right structure
   - write the implementation plan before dispatch
4. If backlog is skipped:
   - write a short working plan in-thread
   - proceed without fake ticketing
5. Choose execution mode:
   - no subagents for trivial work
   - one worker for focused work
   - parallel workers only for disjoint scopes
6. Run verification before handoff.

## Dispatch Rules

The scrum master orchestrates. Workers implement.

- Do not become the default implementer unless delegation is unnecessary.
- Do not parallelize overlapping files or tightly coupled runtime work.
- Give every worker explicit ownership of files/modules.
- Tell every worker other agents may be active and they must not revert unrelated edits.
- Require each worker to report:
  - changed files
  - tests run
  - blockers

Use worker agents for implementation and explorer agents only for bounded codebase questions.

## Verification

Every nontrivial code task gets verification.

Preferred flow:
1. use `subminer-change-verification`
2. start with the cheapest sufficient lane
3. escalate only when needed
4. if worker verification is sufficient, accept it or run one final consolidating pass

Never hand off nontrivial work without stating what was verified and what was skipped.

## Pre-Handoff Policy Checks (Required)

Before handoff, always ask and answer both of these questions explicitly:

1. **Docs update required?**
2. **Changelog fragment required?**

Rules:
- Do not assume silence implies "no." Record an explicit yes/no decision for each item.
- If the answer is yes, either complete the update or report the blocker before handoff.
- Include the final answers in the handoff summary even when both answers are "no."

## Failure / Scope Handling

- If a worker hits ambiguity, pause and ask the user.
- If verification fails, either:
  - send the worker back with exact failure context, or
  - fix it directly if it is tiny and clearly in scope
- If new scope appears, revisit backlog structure before silently expanding work.

## Representative Flows

### Trivial no-ticket work

- decide backlog is unnecessary
- keep a short plan
- implement directly or with one worker if helpful
- run targeted verification
- report outcome concisely

### Single-task implementation

- search/create/update one task
- record plan
- dispatch one worker
- integrate
- verify
- update task and report outcome

### Parent + subtasks execution

- search/create/update parent task
- create subtasks for distinct deliverables/phases
- record sequencing in the plan
- dispatch workers only where scopes are disjoint
- integrate
- run consolidated verification
- update task state and report outcome

## Output Expectations

At the end, report:
- whether backlog was used and what changed
- which workers were dispatched and what they owned
- what verification ran
- explicit answers to:
  - docs update required?
  - changelog fragment required?
- blockers, skips, and risks
