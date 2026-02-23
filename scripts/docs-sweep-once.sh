#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-$HOME/projects/japanese/SubMiner}"
LOCK_FILE="${LOCK_FILE:-/tmp/subminer-doc-sweep.lock}"
STATE_FILE="${STATE_FILE:-/tmp/subminer-doc-sweep.state}"
LOG_FILE="${LOG_FILE:-$REPO/.codex-doc-sweep.log}"
TIMEOUT_SECONDS="${TIMEOUT_SECONDS:-240}"
SUBAGENT_ROOT="${SUBAGENT_ROOT:-$REPO/docs/subagents}"
SUBAGENT_INDEX_FILE="${SUBAGENT_INDEX_FILE:-$SUBAGENT_ROOT/INDEX.md}"
SUBAGENT_COLLAB_FILE="${SUBAGENT_COLLAB_FILE:-$SUBAGENT_ROOT/collaboration.md}"
SUBAGENT_AGENTS_DIR="${SUBAGENT_AGENTS_DIR:-$SUBAGENT_ROOT/agents}"
LEGACY_SUBAGENT_FILE="${LEGACY_SUBAGENT_FILE:-$REPO/docs/subagent.md}"
AGENT_ID="${AGENT_ID:-docs-sweep}"
AGENT_ALIAS="${AGENT_ALIAS:-Docs Sweep}"
AGENT_MISSION="${AGENT_MISSION:-Docs drift cleanup and coordination updates}"

# Non-interactive agent command used to run the prompt.
# Example:
#   AGENT_CMD='codex exec'
#   AGENT_CMD='opencode run'
AGENT_CMD="${AGENT_CMD:-codex exec}"
AGENT_ID_SAFE="$(printf '%s' "$AGENT_ID" | tr -c 'A-Za-z0-9._-' '_')"
AGENT_FILE="${SUBAGENT_AGENTS_DIR}/${AGENT_ID_SAFE}.md"

mkdir -p "$(dirname "$LOCK_FILE")"
mkdir -p "$(dirname "$STATE_FILE")"
mkdir -p "$SUBAGENT_ROOT" "$SUBAGENT_AGENTS_DIR" "$SUBAGENT_ROOT/archive"

exec 9> "$LOCK_FILE"
if ! flock -n 9; then
	exit 0
fi

cd "$REPO"

current_state="$({
	git status --porcelain=v1
	git ls-files --others --exclude-standard
} | sha256sum | cut -d' ' -f1)"

previous_state="$(cat "$STATE_FILE" 2> /dev/null || true)"
if [[ "$current_state" == "$previous_state" ]]; then
	exit 0
fi

printf '%s' "$current_state" > "$STATE_FILE"

run_started_at="$(date -Is)"
run_started_utc="$(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "[RUN] [$run_started_at] docs sweep running (agent_id=$AGENT_ID alias=$AGENT_ALIAS)"
echo "[$run_started_at] state changed; starting docs sweep (agent_id=$AGENT_ID alias=$AGENT_ALIAS)" >> "$LOG_FILE"

if [[ ! -f "$SUBAGENT_INDEX_FILE" ]]; then
	cat > "$SUBAGENT_INDEX_FILE" << 'EOF'
# Subagents Index

Read first. Keep concise.

| agent_id | alias | mission | status | file | last_update_utc |
| --- | --- | --- | --- | --- | --- |
EOF
fi

if [[ ! -f "$SUBAGENT_COLLAB_FILE" ]]; then
	cat > "$SUBAGENT_COLLAB_FILE" << 'EOF'
# Subagents Collaboration

Shared notes. Append-only.

- [YYYY-MM-DDTHH:MM:SSZ] [agent_id|alias] note, question, dependency, conflict, decision.
EOF
fi

if [[ ! -f "$AGENT_FILE" ]]; then
	cat > "$AGENT_FILE" << EOF
# Agent: $AGENT_ID

- alias: $AGENT_ALIAS
- mission: $AGENT_MISSION
- status: planning
- branch: unknown
- started_at: $run_started_utc
- heartbeat_minutes: 20

## Current Work (newest first)
- [$run_started_utc] intent: initialize section

## Files Touched
- none yet

## Assumptions
- none yet

## Open Questions / Blockers
- none

## Next Step
- continue run
EOF
fi

if [[ -f "$LEGACY_SUBAGENT_FILE" ]]; then
	echo "[WARN] [$run_started_at] legacy file exists; prefer sharded layout: $LEGACY_SUBAGENT_FILE" | tee -a "$LOG_FILE"
fi

read -r -d '' PROMPT << EOF || true
Watch for in-flight refactors. If repo changes introduced drift, update only:
- README.md
- AGENTS.md
- docs/**/*.md
- config.example.jsonc
- docs/public/config.example.jsonc <-- generated automatically with make generate-example-config / bun run generate:config-example
- package.json scripts/config references (only if needed)

Coordination protocol:
- Read in order before edits:
  1) \`$SUBAGENT_INDEX_FILE\`
  2) \`$SUBAGENT_COLLAB_FILE\`
  3) \`$AGENT_FILE\`
- Edit scope:
  - MAY edit own file: \`$AGENT_FILE\`
  - MAY append to collaboration: \`$SUBAGENT_COLLAB_FILE\`
  - MAY update own row in index: \`$SUBAGENT_INDEX_FILE\`
  - MUST NOT edit other agent files in \`$SUBAGENT_AGENTS_DIR\`
- Ensure own file has updated: alias, mission, status, branch, started_at, heartbeat_minutes.
- Add UTC ISO entries in "Current Work (newest first)" for intent/progress/handoff for this run.
- Keep own file sections current: Files Touched, assumptions, blockers, next step.
- Ensure index row for \`$AGENT_ID\` reflects alias/mission/status/file/last_update_utc.
- If file conflict/dependency seen, append note in collaboration.

Run metadata:
- run_started_at_utc: $run_started_utc
- repo: $REPO
- agent_id: $AGENT_ID
- agent_alias: $AGENT_ALIAS
- agent_file: $AGENT_FILE

Rules:
- Keep edits minimal and accurate to current code.
- Do not commit.
- Do not push.
- If ambiguous, do not guess; skip and report uncertainty.
- Print concise summary with:
  1) files changed + why
  2) coordination updates made (\`$SUBAGENT_INDEX_FILE\`, \`$SUBAGENT_COLLAB_FILE\`, \`$AGENT_FILE\`)
  3) open questions/blockers
EOF

quoted_prompt="$(printf '%q' "$PROMPT")"

job_status=0
if timeout "${TIMEOUT_SECONDS}s" bash -lc "$AGENT_CMD $quoted_prompt" >> "$LOG_FILE" 2>&1; then
	run_finished_at="$(date -Is)"
	echo "[OK]  [$run_finished_at] docs sweep complete (agent_id=$AGENT_ID)"
	echo "[$run_finished_at] docs sweep complete (agent_id=$AGENT_ID)" >> "$LOG_FILE"
else
	run_failed_at="$(date -Is)"
	exit_code=$?
	job_status=$exit_code
	echo "[FAIL] [$run_failed_at] docs sweep failed (exit $exit_code, agent_id=$AGENT_ID)"
	echo "[$run_failed_at] docs sweep failed (exit $exit_code, agent_id=$AGENT_ID)" >> "$LOG_FILE"
fi

exit "$job_status"
