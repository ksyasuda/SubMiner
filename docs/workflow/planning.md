<!-- read_when: deciding whether work needs a plan or writing one -->

# Planning

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: the task spans multiple files, subsystems, or verification lanes

## Plan Types

- Lightweight plan: small change, a few reversible steps, minimal coordination
- Execution plan: nontrivial feature/refactor/debugging effort with multiple phases or important decisions

## Use a Lightweight Plan When

- one subsystem
- obvious change shape
- low risk
- easy to verify

## Use an Execution Plan When

- multiple subsystems or runtimes
- architectural tradeoffs matter
- staged verification is needed
- the work should be resumable by another agent or human

## Plan Location

- active design and implementation docs live in `docs/plans/`
- keep names date-prefixed and task-specific
- remove or archive old plans deliberately; do not leave mystery artifacts

## Plan Contents

- problem / goal
- non-goals
- file ownership or edit scope
- verification plan
- decisions made during execution
