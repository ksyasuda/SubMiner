<!-- read_when: migrating or using the repo-local SubMiner workflow plugin -->

# SubMiner Workflow Plugin

Status: active
Last verified: 2026-03-26
Owner: Kyle Yasuda
Read when: using or updating the repo-local plugin that owns SubMiner agent workflow skills

This plugin is the canonical source of truth for the SubMiner agent workflow packaging.

## Contents

- `skills/subminer-scrum-master/`
  - backlog-first intake, planning, dispatch, and handoff workflow
- `skills/subminer-change-verification/`
  - cheap-first verification workflow plus helper scripts

## Backlog MCP

- This plugin assumes Backlog.md MCP is available in the host environment when the client exposes it.
- Canonical backlog behavior remains:
  - read `backlog://workflow/overview` when resources are available
  - otherwise use the matching backlog tool overview
- If backlog MCP is unavailable in the current session, fall back to direct repo-local `backlog/` edits and record that blocker in the task or handoff.

## Compatibility

- `.agents/skills/subminer-scrum-master/` is a compatibility shim that redirects to the plugin-owned skill.
- `.agents/skills/subminer-change-verification/` is a compatibility shim.
- `.agents/skills/subminer-change-verification/scripts/*.sh` remain as wrapper entrypoints so existing docs, backlog tasks, and shell history keep working.

## Verification

For plugin/doc/shim changes, prefer:

```bash
bun run test:docs:kb
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane docs --lane core \
  plugins/subminer-workflow \
  .agents/skills/subminer-scrum-master/SKILL.md \
  .agents/skills/subminer-change-verification/SKILL.md \
  .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh \
  .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh \
  .agents/plugins/marketplace.json \
  docs/workflow/README.md \
  docs/workflow/agent-plugins.md \
  backlog/tasks/task-240\ -\ Migrate-SubMiner-agent-skills-into-a-repo-local-plugin-workflow.md
```
