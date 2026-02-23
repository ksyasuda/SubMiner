#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
RUN_ONCE_SCRIPT="$SCRIPT_DIR/docs-sweep-once.sh"
INTERVAL_SECONDS="${INTERVAL_SECONDS:-300}"
REPO="${REPO:-$HOME/projects/japanese/SubMiner}"
LOG_FILE="${LOG_FILE:-$REPO/.codex-doc-sweep.log}"
SUBAGENT_ROOT="${SUBAGENT_ROOT:-$REPO/docs/subagents}"
SUBAGENT_INDEX_FILE="${SUBAGENT_INDEX_FILE:-$SUBAGENT_ROOT/INDEX.md}"
SUBAGENT_COLLAB_FILE="${SUBAGENT_COLLAB_FILE:-$SUBAGENT_ROOT/collaboration.md}"
SUBAGENT_AGENTS_DIR="${SUBAGENT_AGENTS_DIR:-$SUBAGENT_ROOT/agents}"
AGENT_ID="${AGENT_ID:-docs-sweep}"
AGENT_ID_SAFE="$(printf '%s' "$AGENT_ID" | tr -c 'A-Za-z0-9._-' '_')"
AGENT_FILE="${AGENT_FILE:-$SUBAGENT_AGENTS_DIR/${AGENT_ID_SAFE}.md}"
REPORT_WITH_CODEX=false
REPORT_TIMEOUT_SECONDS="${REPORT_TIMEOUT_SECONDS:-120}"
REPORT_AGENT_CMD="${REPORT_AGENT_CMD:-codex exec}"

if [[ ! -x "$RUN_ONCE_SCRIPT" ]]; then
	echo "Missing executable: $RUN_ONCE_SCRIPT"
	echo "Run: chmod +x scripts/docs-sweep-once.sh"
	exit 1
fi

usage() {
	cat << 'EOF'
Usage: scripts/docs-sweep-watch.sh [options]

Options:
  -r, --report  One-off: summarize current log with Codex and exit.
  -h, --help    Show this help message.

Environment:
  AGENT_ID            Stable agent id (default: docs-sweep)
  AGENT_ALIAS         Human label shown in logs/coordination (default: Docs Sweep)
  AGENT_MISSION       One-line focus for this run
  SUBAGENT_ROOT       Coordination root (default: docs/subagents)
EOF
}

trim_log_runs() {
	# Keep only the last 50 docs-sweep runs in the shared log file.
	if [[ ! -f "$LOG_FILE" ]]; then
		return
	fi

	local keep_runs=50
	local start_line
	start_line="$(
		awk -v max="$keep_runs" '
			/state changed; starting docs sweep/ { lines[++count] = NR }
			END {
				if (count > max) print lines[count - max + 1]
				else print 0
			}
		' "$LOG_FILE"
	)"

	if [[ "$start_line" =~ ^[0-9]+$ ]] && (( start_line > 0 )); then
		local tmp_file
		tmp_file="$(mktemp "${LOG_FILE}.XXXXXX")"
		tail -n +"$start_line" "$LOG_FILE" > "$tmp_file"
		mv "$tmp_file" "$LOG_FILE"
	fi
}

run_report() {
	local has_log=false
	local has_index=false
	local has_collab=false
	local has_agent_file=false
	if [[ -s "$LOG_FILE" ]]; then
		has_log=true
	fi
	if [[ -s "$SUBAGENT_INDEX_FILE" ]]; then
		has_index=true
	fi
	if [[ -s "$SUBAGENT_COLLAB_FILE" ]]; then
		has_collab=true
	fi
	if [[ -s "$AGENT_FILE" ]]; then
		has_agent_file=true
	fi
	if [[ "$has_log" != "true" && "$has_index" != "true" && "$has_collab" != "true" && "$has_agent_file" != "true" ]]; then
		echo "[REPORT] no inputs; missing/empty files:"
		echo "[REPORT] - $LOG_FILE"
		echo "[REPORT] - $SUBAGENT_INDEX_FILE"
		echo "[REPORT] - $SUBAGENT_COLLAB_FILE"
		echo "[REPORT] - $AGENT_FILE"
		return
	fi

	local report_prompt
	read -r -d '' report_prompt << EOF || true
Summarize docs sweep state. Output:
- Changes made (short bullets; file-focused when possible)
- Agent coordination updates from sharded docs/subagents files
- Open questions / uncertainty
- Left undone / follow-up items

Constraints:
- Be concise.
- If uncertain, say uncertain.
Read these files directly if present:
$LOG_FILE
$SUBAGENT_INDEX_FILE
$SUBAGENT_COLLAB_FILE
$AGENT_FILE
EOF

	echo "[REPORT] codex summary start"
	local report_file
	local report_stderr
	report_file="$(mktemp /tmp/docs-sweep-report.XXXXXX)"
	report_stderr="$(mktemp /tmp/docs-sweep-report-stderr.XXXXXX)"
	(
		cd "$REPO"
		timeout "${REPORT_TIMEOUT_SECONDS}s" bash -lc "$REPORT_AGENT_CMD -o $(printf '%q' "$report_file") $(printf '%q' "$report_prompt")" > /dev/null 2> "$report_stderr"
	)
	local report_exit=$?
	if (( report_exit != 0 )); then
		echo "[REPORT] codex summary failed (exit $report_exit)"
		cat "$report_stderr"
		echo
		echo "[REPORT] codex summary end"
		return
	fi
	if [[ -s "$report_file" ]]; then
		cat "$report_file"
	else
		echo "[REPORT] codex produced no final message"
	fi
	echo
	echo "[REPORT] codex summary end"
}

while (( $# > 0 )); do
	case "$1" in
	-r|--report)
		REPORT_WITH_CODEX=true
		shift
		;;
	-h|--help)
		usage
		exit 0
		;;
	*)
		echo "Unknown option: $1"
		usage
		exit 1
		;;
	esac
done

if [[ "$REPORT_WITH_CODEX" == "true" ]]; then
	trim_log_runs
	run_report
	exit 0
fi

stop_requested=false
trap 'stop_requested=true' INT TERM

echo "Starting docs sweep watcher (interval: ${INTERVAL_SECONDS}s, subagent_root: ${SUBAGENT_ROOT}). Press Ctrl+C to stop."

while true; do
	run_started_at="$(date -Is)"
	echo "[RUN] [$run_started_at] docs sweep cycle running"
	if "$RUN_ONCE_SCRIPT"; then
		run_finished_at="$(date -Is)"
		echo "[OK]  [$run_finished_at] docs sweep cycle complete"
	else
		run_failed_at="$(date -Is)"
		exit_code=$?
		echo "[FAIL] [$run_failed_at] docs sweep cycle failed (exit $exit_code)"
	fi
	trim_log_runs

	if [[ "$stop_requested" == "true" ]]; then
		break
	fi

	sleep "$INTERVAL_SECONDS" &
	wait $!

	if [[ "$stop_requested" == "true" ]]; then
		break
	fi
done

echo "Docs sweep watcher stopped."
