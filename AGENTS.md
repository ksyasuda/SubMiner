<!-- BACKLOG.MD MCP GUIDELINES START -->

<CRITICAL_INSTRUCTION>

## BACKLOG WORKFLOW INSTRUCTIONS

This project uses Backlog.md MCP for all task and project management activities.

**CRITICAL GUIDANCE**

- If your client supports MCP resources, read `backlog://workflow/overview` to understand when and how to use Backlog for this project.
- If your client only supports tools or the above request fails, call `backlog.get_workflow_overview()` tool to load the tool-oriented overview (it lists the matching guide tools).

- **First time working here?** Read the overview resource IMMEDIATELY to learn the workflow
- **Already familiar?** You should have the overview cached ("## Backlog.md Overview (MCP)")
- **When to read it**: BEFORE creating tasks, or when you're unsure whether to track work

These guides cover:

- Decision framework for when to create tasks
- Search-first workflow to avoid duplicates
- Links to detailed guides for task creation, execution, and finalization
- MCP tools reference

You MUST read the overview resource to understand the complete workflow. The information is NOT summarized here.

</CRITICAL_INSTRUCTION>

<!-- BACKLOG.MD MCP GUIDELINES END -->

## Subagent Coordination Protocol (`docs/subagents/`)

Purpose: multi-agent coordination across runs; single-agent continuity during long runs.

Layout:
- `docs/subagents/INDEX.md` (active agents table)
- `docs/subagents/collaboration.md` (shared notes)
- `docs/subagents/agents/<agent_id>.md` (one file per agent)
- `docs/subagents/archive/<yyyy-mm>/` (archived histories)

Required behavior (all agents):

1. At run start, read in order:
   - `docs/subagents/INDEX.md`
   - `docs/subagents/collaboration.md`
   - your own file: `docs/subagents/agents/<agent_id>.md`
2. Identify self by stable `agent_id` (runner/env-provided). If missing, create own file from template.
3. Maintain `alias` (short human-readable label) + `mission` (one-line focus).
4. Before coding:
   - record intent, planned files, assumptions in your own file.
5. During run:
   - update on phase changes (plan -> edit -> test -> handoff),
   - heartbeat at least every `HEARTBEAT_MINUTES` (default 5),
   - update your own row in `INDEX.md` (`status`, `last_update_utc`),
   - append cross-agent notes in `collaboration.md` when needed.
6. Write limits:
   - MAY edit own file.
   - MAY append to `collaboration.md`.
   - MAY edit only own row in `INDEX.md`.
   - MUST NOT edit other agent files.
7. At run end:
   - record files touched, key decisions, assumptions, blockers, next step for handoff.
8. Conflict handling:
   - if another agent touched your target files, add conflict note in `collaboration.md` before continuing.
9. Brevity:
   - terse bullets; factual; no long prose.

Suggested env vars:

- `AGENT_ID` (required)
- `AGENT_ALIAS` (required)
- `HEARTBEAT_MINUTES` (optional, default 20)
