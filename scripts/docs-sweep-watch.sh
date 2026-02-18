#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
RUN_ONCE_SCRIPT="$SCRIPT_DIR/docs-sweep-once.sh"
INTERVAL_SECONDS="${INTERVAL_SECONDS:-300}"
REPO="${REPO:-$HOME/projects/japanese/SubMiner}"
LOG_FILE="${LOG_FILE:-$REPO/.codex-doc-sweep.log}"
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
	if [[ ! -f "$LOG_FILE" ]]; then
		echo "[REPORT] log file not found: $LOG_FILE"
		return
	fi

	if [[ ! -s "$LOG_FILE" ]]; then
		echo "[REPORT] log file empty: $LOG_FILE"
		return
	fi

	local report_prompt
	read -r -d '' report_prompt << EOF || true
Summarize docs sweep log. Output:
- Changes made (short bullets; file-focused when possible)
- Open questions / uncertainty
- Left undone / follow-up items

Constraints:
- Be concise.
- If uncertain, say uncertain.
Read this file directly:
$LOG_FILE
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

echo "Starting docs sweep watcher (interval: ${INTERVAL_SECONDS}s). Press Ctrl+C to stop."

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
