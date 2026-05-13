---
name: 'subminer-scrum-master'
description: 'Use in the SubMiner repo when a request should be turned into planned work and driven through execution. Records a plan, dispatches one or more subagents when useful, and requires verification before handoff.'
---

# SubMiner Scrum Master

Canonical source: this plugin path.

Own workflow, not code by default.

Use this skill when the user gives a feature request, bug report, issue, refactor, or implementation ask and the agent should manage intake, planning, worker dispatch, and verification through completion.

## Core Rules

1. Keep the process light for questions, obvious mechanical edits, and tiny isolated changes.
2. Record a plan before dispatching coding work.
3. Split multi-part work into clear phases and ownership areas.
4. Dispatch conservatively. Parallelize only disjoint write scopes.
5. Require verification before handoff, typically via `subminer-change-verification`.
6. Report dispatched workers, verification, blockers, and remaining risks.

## Intake Workflow

1. Parse the request.
   Classify it as question, mechanical edit, bugfix, feature, refactor, investigation, or follow-up.
2. Write a short working plan in-thread when the work is nontrivial.
3. Choose execution mode:
   - no subagents for trivial work
   - one worker for focused work
   - parallel workers only for disjoint scopes
4. Run verification before handoff.

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

## Pre-Handoff Policy Checks

Before handoff, always ask and answer both questions explicitly:

1. Docs update required?
2. Changelog fragment required?

Rules:

- Do not assume silence implies "no."
- If the answer is yes, complete the update or report the blocker.
- Include final yes/no answers in the handoff summary even when both answers are "no."

## Failure / Scope Handling

- If a worker hits ambiguity, pause and ask the user.
- If verification fails, either:
  - send the worker back with exact failure context, or
  - fix it directly if it is tiny and clearly in scope
- If new scope appears, pause and re-plan before silently expanding work.

## Representative Flows

### Trivial work

- keep a short plan
- implement directly or with one worker if helpful
- run targeted verification
- report outcome concisely

### Focused implementation

- record plan
- dispatch one worker
- integrate
- verify
- report outcome

### Multi-part execution

- define distinct deliverables/phases
- record sequencing in the plan
- dispatch workers only where scopes are disjoint
- integrate
- run consolidated verification
- report outcome

## Output Expectations

At the end, report:

- which workers were dispatched and what they owned
- what verification ran
- explicit answers to:
  - docs update required?
  - changelog fragment required?
- blockers, skips, and risks
