# SubMiner Change Verification Skill Design

**Date:** 2026-03-10
**Status:** Approved

## Goal

Create a SubMiner-specific skill that agents can use to verify code changes with automated checks. The skill must support both targeted regression testing during debugging and pre-handoff verification before final response.

## Skill Contract

- **Name:** `subminer-change-verification`
- **Trigger:** Use when working in the SubMiner repo and you need to verify code changes actually work, especially for launcher, mpv, plugin, overlay, runtime, Electron, or env-sensitive behavior.
- **Default posture:** cheap-first; prefer repo-native tests and narrow lanes before broader or GUI-dependent verification.
- **Outputs:**
  - verification summary
  - exact commands run
  - artifact paths for logs, captured summaries, and preserved temp state on failures
  - skipped lanes and blockers
- **Non-goals:**
  - replacing the repo's native tests
  - launching real GUI apps for every change
  - default visual regression or pixel-diff workflows

## Lane Selection

The skill chooses lanes from the diff or explicit file list.

- **`docs`**
  - For `docs-site/`, `docs/`, and similar documentation-only changes.
  - Prefer `bun run docs:test` and `bun run docs:build`.
- **`config`**
  - For `src/config/`, config example generation/verification paths, and config-template-sensitive changes.
  - Prefer `bun run test:config`.
- **`core`**
  - For general source-level changes where type safety and the fast maintained lane are the best cheap signal.
  - Prefer `bun run typecheck` and `bun run test:fast`.
- **`launcher-plugin`**
  - For `launcher/`, `plugin/subminer/`, plugin gating scripts, and wrapper/mpv routing work.
  - Prefer `bun run test:launcher:smoke:src` and `bun run test:plugin:src`.
- **`runtime-compat`**
  - For runtime/composition/bundled behavior where dist-sensitive validation matters.
  - Prefer `bun run build`, `bun run test:runtime:compat`, and `bun run test:smoke:dist`.
- **`real-gui`**
  - Reserved for cases where actual Electron/mpv/window behavior must be validated.
  - Not part of the default lane set; the classifier marks these changes as candidates so the agent can escalate deliberately.

## Escalation Rules

1. Start with the narrowest lane that credibly exercises the changed behavior.
2. If a narrow lane fails in a way that suggests broader fallout, expand once.
3. If a change touches launcher/mpv/plugin/runtime/overlay/window tracking paths, include the relevant specialized lanes before falling back to broad suites.
4. Treat real GUI/mpv verification as opt-in escalation:
   - use only when cheaper evidence is insufficient
   - allow for platform/display/permission blockers
   - report skipped/blocker states explicitly

## Helper Script Design

The skill uses two small shell helpers:

- **`scripts/classify_subminer_diff.sh`**
  - Accepts explicit paths or discovers local changes from git.
  - Emits lane suggestions and flags in a simple line-oriented format.
  - Marks real GUI-sensitive paths as `flag:real-gui-candidate` instead of forcing GUI execution.
- **`scripts/verify_subminer_change.sh`**
  - Creates an artifact directory under `.tmp/skill-verification/<timestamp>/`.
  - Selects lanes from the classifier unless lanes are supplied explicitly.
  - Runs repo-native commands in a stable order and captures stdout/stderr per step.
  - Writes a compact `summary.json` and a human-readable `summary.txt`.
  - Skips real GUI verification unless explicitly enabled.

## Artifact Contract

Each invocation should create:

- `summary.json`
- `summary.txt`
- `classification.txt`
- `env.txt`
- `lanes.txt`
- `steps.tsv`
- `steps/<step>.stdout.log`
- `steps/<step>.stderr.log`

Failures should preserve the artifact directory and identify the exact failing command and log paths.

## Agent Workflow

1. Inspect changed files or requested area.
2. Classify the change into verification lanes.
3. Run the cheapest sufficient lane set.
4. Escalate only if evidence is insufficient.
5. Escalate to real GUI/mpv only for actual Electron/mpv/window behavior claims.
6. Return a short report with:
   - pass/fail/skipped per lane
   - exact commands run
   - artifact paths
   - blockers/gaps

## Initial Implementation Scope

- Ship the skill entrypoint plus the classifier/verifier helpers.
- Make real GUI verification an explicit future hook rather than a default workflow.
- Verify the new skill locally with representative classifier output and artifact generation.
