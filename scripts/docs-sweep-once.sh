#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-$HOME/projects/japanese/SubMiner}"
LOCK_FILE="${LOCK_FILE:-/tmp/subminer-doc-sweep.lock}"
STATE_FILE="${STATE_FILE:-/tmp/subminer-doc-sweep.state}"
LOG_FILE="${LOG_FILE:-$REPO/.codex-doc-sweep.log}"
TIMEOUT_SECONDS="${TIMEOUT_SECONDS:-240}"

# Non-interactive agent command used to run the prompt.
# Example:
#   AGENT_CMD='codex exec'
#   AGENT_CMD='opencode run'
AGENT_CMD="${AGENT_CMD:-codex exec}"

read -r -d '' PROMPT << 'EOF' || true
Watch for in-flight refactors. If repo changes introduced drift, update only:
- README.md
- docs/**/*.md
- config.example.jsonc
- docs/public/config.example.jsonc <-- generated automatically with make pnpm run generate:config-example
- package.json scripts/config references (only if needed)

Rules:
- Keep edits minimal and accurate to current code.
- Do not commit.
- Do not push.
- If ambiguous, do not guess; skip and report uncertainty.
- Print concise summary of files changed and why.
EOF

mkdir -p "$(dirname "$LOCK_FILE")"
mkdir -p "$(dirname "$STATE_FILE")"

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

quoted_prompt="$(printf '%q' "$PROMPT")"

run_started_at="$(date -Is)"
echo "[RUN] [$run_started_at] docs sweep running"
echo "[$run_started_at] state changed; starting docs sweep" >> "$LOG_FILE"

job_status=0
if timeout "${TIMEOUT_SECONDS}s" bash -lc "$AGENT_CMD $quoted_prompt" >> "$LOG_FILE" 2>&1; then
	run_finished_at="$(date -Is)"
	echo "[OK]  [$run_finished_at] docs sweep complete"
	echo "[$run_finished_at] docs sweep complete" >> "$LOG_FILE"
else
	run_failed_at="$(date -Is)"
	exit_code=$?
	job_status=$exit_code
	echo "[FAIL] [$run_failed_at] docs sweep failed (exit $exit_code)"
	echo "[$run_failed_at] docs sweep failed (exit $exit_code)" >> "$LOG_FILE"
fi

exit "$job_status"
