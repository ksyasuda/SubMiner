---
id: TASK-240
title: Migrate SubMiner agent skills into a repo-local plugin workflow
status: Done
assignee:
  - codex
created_date: '2026-03-26 00:00'
updated_date: '2026-03-26 23:23'
labels:
  - skills
  - plugin
  - workflow
  - backlog
  - tooling
dependencies:
  - TASK-159
  - TASK-160
priority: high
ordinal: 24000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Turn the current SubMiner-specific repo skills into a reproducible repo-local plugin workflow. The plugin should become the canonical source of truth for the SubMiner scrum-master and change-verification skills, bundle the scripts and metadata needed to test and validate changes, and preserve compatibility for existing repo references through thin `.agents/skills/` shims while the migration settles.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 A repo-local plugin scaffold exists for the SubMiner workflow, with manifest and marketplace metadata wired according to the repo-local plugin layout.
- [x] #2 `subminer-scrum-master` and `subminer-change-verification` live under the plugin as the canonical skill sources, along with any helper scripts or supporting files needed for reproducible use.
- [x] #3 Existing repo-level `.agents/skills/` entrypoints are reduced to compatibility shims or redirects instead of remaining as duplicate sources of truth.
- [x] #4 The plugin-owned workflow explicitly documents backlog-first orchestration and change verification expectations, including how the skills work together.
- [x] #5 The migration is validated with the cheapest sufficient repo-native verification lane and the task records the exact commands and any skips/blockers.
<!-- SECTION:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Inspect the plugin-creator contract and current repo skill/script layout, then choose the plugin name, directory structure, and migration boundaries.
2. Scaffold a repo-local plugin plus marketplace entry, keeping the plugin payload under `plugins/<name>/` and the catalog entry under `.agents/plugins/marketplace.json`.
3. Move the two SubMiner-specific skills and their helper scripts into the plugin as the canonical source, adding any plugin docs or supporting metadata needed for reproducible testing/validation.
4. Replace the existing `.agents/skills/subminer-*` surfaces with minimal compatibility shims that point agents at the plugin-owned sources without duplicating logic.
5. Update internal docs or references that should now describe the plugin-first workflow.
6. Run the cheapest sufficient verification lane for plugin/internal-doc changes and record the results in this task.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

2026-03-26: User approved the migration shape where the plugin becomes the canonical source of truth and `.agents/skills/` stays only as compatibility shims. Repo-local plugin chosen over home-local plugin.

2026-03-26: Backlog MCP resources/tools are not available in this Codex session (`MCP startup failed`), so this task is being initialized directly in the repo-local `backlog/` files instead of through the live Backlog MCP interface.

2026-03-26: Scaffolded `plugins/subminer-workflow/` plus `.agents/plugins/marketplace.json`, moved the scrum-master and change-verification skill definitions into the plugin as the canonical sources, and converted the old `.agents/skills/` surfaces into compatibility shims. Preserved the old verifier script entrypoints as wrappers because backlog/docs history already calls them directly.

2026-03-26: Verification passed.

- `bash -n plugins/subminer-workflow/skills/subminer-change-verification/scripts/classify_subminer_diff.sh`
- `bash -n plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh`
- `bash -n .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh`
- `bash -n .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh plugins/subminer-workflow/.codex-plugin/plugin.json docs/workflow/agent-plugins.md .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane docs plugins/subminer-workflow .agents/skills/subminer-scrum-master/SKILL.md .agents/skills/subminer-change-verification/SKILL.md .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh .agents/plugins/marketplace.json docs/workflow/README.md docs/workflow/agent-plugins.md 'backlog/tasks/task-240 - Migrate-SubMiner-agent-skills-into-a-repo-local-plugin-workflow.md'`
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260326-232300-E2NQVX/`

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Created a repo-local `subminer-workflow` plugin as the canonical packaging for the SubMiner scrum-master and change-verification workflow. The plugin now owns both skills, the verifier helper scripts, plugin metadata, and workflow docs. The old `.agents/skills/` surfaces remain only as compatibility shims, and the old verifier script paths now forward to the plugin-owned scripts so existing docs and backlog commands continue to work. Targeted plugin/docs verification passed, including wrapper-script syntax checks and a real verifier run through the legacy entrypoint.

<!-- SECTION:FINAL_SUMMARY:END -->
