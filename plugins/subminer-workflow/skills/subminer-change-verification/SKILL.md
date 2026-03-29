---
name: 'subminer-change-verification'
description: 'Use when working in the SubMiner repo and you need to verify code changes actually work. Covers targeted regression checks during debugging and pre-handoff verification, with cheap-first lane selection for config, docs, launcher/plugin, runtime-compat, and optional real-runtime escalation.'
---

# SubMiner Change Verification

Canonical source: this plugin path.

Use this skill for SubMiner code changes. Default to cheap, repo-native verification first. Escalate only when the changed behavior actually depends on Electron, mpv, overlay/window tracking, or other GUI-sensitive runtime behavior.

## Scripts

- `scripts/classify_subminer_diff.sh`
  - Emits suggested lanes and flags from explicit paths or current git changes.
- `scripts/verify_subminer_change.sh`
  - Runs selected lanes, captures artifacts, and writes a compact summary.

If you need an explicit installed path, use the directory that contains this `SKILL.md`. The helper scripts live under:

```bash
export SUBMINER_VERIFY_SKILL="<path-to-plugin-skill>"
```

## Default workflow

1. Inspect the changed files or user-requested area.
2. Run the classifier unless you already know the right lane.
3. Run the verifier with the cheapest sufficient lane set.
4. If the classifier emits `flag:real-runtime-candidate`, do not jump straight to runtime verification. First run the non-runtime lanes.
5. Escalate to explicit `--lane real-runtime --allow-real-runtime` only when cheaper lanes cannot validate the behavior claim.
6. Return:
   - verification summary
   - exact commands run
   - artifact paths
   - skipped lanes and blockers

## Quick start

Plugin-source quick start:

```bash
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/classify_subminer_diff.sh
```

Installed-skill quick start:

```bash
bash "$SUBMINER_VERIFY_SKILL/scripts/classify_subminer_diff.sh"
```

Compatibility entrypoint:

```bash
bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh
```

Classify explicit files:

```bash
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/classify_subminer_diff.sh \
  launcher/main.ts \
  plugin/subminer/lifecycle.lua \
  src/main/runtime/mpv-client-runtime-service.ts
```

Run automatic lane selection:

```bash
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh
```

Installed-skill form:

```bash
bash "$SUBMINER_VERIFY_SKILL/scripts/verify_subminer_change.sh"
```

Compatibility entrypoint:

```bash
bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh
```

Run targeted lanes:

```bash
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh \
  --lane launcher-plugin \
  --lane runtime-compat
```

Dry-run to inspect planned commands and artifact layout:

```bash
bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh \
  --dry-run \
  launcher/main.ts \
  src/main.ts
```

## Lane guidance

- `docs`
  - For `docs-site/`, `docs/`, and doc-only edits.
- `config`
  - For `src/config/` and config-template-sensitive edits.
- `core`
  - For general source changes where `typecheck` + `test:fast` is the best cheap signal.
- `launcher-plugin`
  - For `launcher/`, `plugin/subminer/`, plugin gating scripts, and wrapper/mpv routing work.
- `runtime-compat`
  - For `src/main*`, runtime/composer wiring, mpv/overlay services, window trackers, and dist-sensitive behavior.
- `real-runtime`
  - Only after deliberate escalation.

## Real Runtime Escalation

Escalate only when the change claim depends on actual runtime behavior, for example:

- overlay appears, hides, or tracks a real mpv window
- mpv launch flags or pause-until-ready behavior
- plugin/socket/auto-start handshake under a real player
- macOS/window-tracker/focus-sensitive behavior

If the environment cannot support authoritative runtime verification, report the blocker explicitly. Do not silently downgrade a runtime-required claim to a pass.

## Artifact contract

The verifier writes under `.tmp/skill-verification/<timestamp>/`:

- `summary.json`
- `summary.txt`
- `classification.txt`
- `env.txt`
- `lanes.txt`
- `steps.tsv`
- `steps/*.stdout.log`
- `steps/*.stderr.log`

On failure, quote the exact failing command and point at the artifact directory.
