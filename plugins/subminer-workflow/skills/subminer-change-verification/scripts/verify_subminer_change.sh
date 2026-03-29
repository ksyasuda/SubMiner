#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage: verify_subminer_change.sh [options] [path ...]

Options:
  --lane <name>             Force a verification lane. Repeatable.
  --artifact-dir <dir>      Use an explicit artifact directory.
  --allow-real-runtime      Allow explicit real-runtime execution.
  --allow-real-gui          Deprecated alias for --allow-real-runtime.
  --dry-run                 Record planned steps without executing commands.
  --help                    Show this help text.

If no lanes are supplied, the script classifies the provided paths. If no paths are
provided, it classifies the current local git changes.

Authoritative real-runtime verification should be requested with explicit path
arguments instead of relying on inferred local git changes.
EOF
}

timestamp() {
  date +%Y%m%d-%H%M%S
}

timestamp_iso() {
  date -u +%Y-%m-%dT%H:%M:%SZ
}

generate_session_id() {
  local tmp_dir
  tmp_dir=$(mktemp -d "${TMPDIR:-/tmp}/subminer-verify-$(timestamp)-XXXXXX")
  basename "$tmp_dir"
  rmdir "$tmp_dir"
}

has_item() {
  local needle=$1
  shift || true
  local item
  for item in "$@"; do
    if [[ "$item" == "$needle" ]]; then
      return 0
    fi
  done
  return 1
}

normalize_lane_name() {
  case "$1" in
    real-gui)
      printf '%s' "real-runtime"
      ;;
    *)
      printf '%s' "$1"
      ;;
  esac
}

add_lane() {
  local lane
  lane=$(normalize_lane_name "$1")
  if ! has_item "$lane" "${SELECTED_LANES[@]:-}"; then
    SELECTED_LANES+=("$lane")
  fi
}

add_blocker() {
  BLOCKERS+=("$1")
  BLOCKED=1
}

validate_artifact_dir() {
  local candidate=$1
  if [[ ! "$candidate" =~ ^[A-Za-z0-9._/@:+-]+$ ]]; then
    echo "Invalid characters in --artifact-dir path" >&2
    exit 2
  fi
}

append_step_record() {
  printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
    "$1" "$2" "$3" "$4" "$5" "$6" "$7" "$8" >>"$STEPS_TSV"
}

record_env() {
  {
    printf 'repo_root=%s\n' "$REPO_ROOT"
    printf 'session_id=%s\n' "$SESSION_ID"
    printf 'artifact_dir=%s\n' "$ARTIFACT_DIR"
    printf 'path_selection_mode=%s\n' "$PATH_SELECTION_MODE"
    printf 'dry_run=%s\n' "$DRY_RUN"
    printf 'allow_real_runtime=%s\n' "$ALLOW_REAL_RUNTIME"
    printf 'session_home=%s\n' "$SESSION_HOME"
    printf 'session_xdg_config_home=%s\n' "$SESSION_XDG_CONFIG_HOME"
    printf 'session_mpv_dir=%s\n' "$SESSION_MPV_DIR"
    printf 'session_logs_dir=%s\n' "$SESSION_LOGS_DIR"
    printf 'session_mpv_log=%s\n' "$SESSION_MPV_LOG"
    printf 'pwd=%s\n' "$(pwd)"
    git rev-parse --short HEAD 2>/dev/null | sed 's/^/git_head=/' || true
    git status --short 2>/dev/null || true
    if [[ ${#PATH_ARGS[@]} -gt 0 ]]; then
      printf 'requested_paths=\n'
      printf '  %s\n' "${PATH_ARGS[@]}"
    fi
  } >"$ARTIFACT_DIR/env.txt"
}

run_step() {
  local lane=$1
  local name=$2
  local command=$3
  local note=${4:-}
  local lane_slug=${lane//[^a-zA-Z0-9_-]/-}
  local slug=${name//[^a-zA-Z0-9_-]/-}
  local step_slug="${lane_slug}--${slug}"
  local stdout_rel="steps/${step_slug}.stdout.log"
  local stderr_rel="steps/${step_slug}.stderr.log"
  local stdout_path="$ARTIFACT_DIR/$stdout_rel"
  local stderr_path="$ARTIFACT_DIR/$stderr_rel"
  local status exit_code

  COMMANDS_RUN+=("$command")
  printf '%s\n' "$command" >"$ARTIFACT_DIR/steps/${step_slug}.command.txt"

  if [[ "$DRY_RUN" == "1" ]]; then
    printf '[dry-run] %s\n' "$command" >"$stdout_path"
    : >"$stderr_path"
    status="dry-run"
    exit_code=0
  else
    if HOME="$SESSION_HOME" \
      XDG_CONFIG_HOME="$SESSION_XDG_CONFIG_HOME" \
      SUBMINER_SESSION_LOGS_DIR="$SESSION_LOGS_DIR" \
      SUBMINER_SESSION_MPV_LOG="$SESSION_MPV_LOG" \
      bash -c "cd \"$REPO_ROOT\" && $command" >"$stdout_path" 2>"$stderr_path"; then
      status="passed"
      exit_code=0
      EXECUTED_REAL_STEPS=1
    else
      exit_code=$?
      status="failed"
      FAILED=1
    fi
  fi

  append_step_record "$lane" "$name" "$status" "$exit_code" "$command" "$stdout_rel" "$stderr_rel" "$note"
  printf '%s\t%s\t%s\n' "$lane" "$name" "$status"

  if [[ "$status" == "failed" ]]; then
    FAILURE_STEP="$name"
    FAILURE_COMMAND="$command"
    FAILURE_STDOUT="$stdout_rel"
    FAILURE_STDERR="$stderr_rel"
    return "$exit_code"
  fi
}

record_nonpassing_step() {
  local lane=$1
  local name=$2
  local status=$3
  local note=$4
  local lane_slug=${lane//[^a-zA-Z0-9_-]/-}
  local slug=${name//[^a-zA-Z0-9_-]/-}
  local step_slug="${lane_slug}--${slug}"
  local stdout_rel="steps/${step_slug}.stdout.log"
  local stderr_rel="steps/${step_slug}.stderr.log"
  printf '%s\n' "$note" >"$ARTIFACT_DIR/$stdout_rel"
  : >"$ARTIFACT_DIR/$stderr_rel"
  append_step_record "$lane" "$name" "$status" "0" "" "$stdout_rel" "$stderr_rel" "$note"
  printf '%s\t%s\t%s\n' "$lane" "$name" "$status"
}

record_skipped_step() {
  record_nonpassing_step "$1" "$2" "skipped" "$3"
}

record_blocked_step() {
  add_blocker "$3"
  record_nonpassing_step "$1" "$2" "blocked" "$3"
}

record_failed_step() {
  FAILED=1
  FAILURE_STEP=$2
  FAILURE_COMMAND=${FAILURE_COMMAND:-"(validation)"}
  local lane_slug=${1//[^a-zA-Z0-9_-]/-}
  local step_slug=${2//[^a-zA-Z0-9_-]/-}
  FAILURE_STDOUT="steps/${lane_slug}--${step_slug}.stdout.log"
  FAILURE_STDERR="steps/${lane_slug}--${step_slug}.stderr.log"
  add_blocker "$3"
  record_nonpassing_step "$1" "$2" "failed" "$3"
}

find_real_runtime_helper() {
  local candidate
  for candidate in \
    "$SCRIPT_DIR/run_real_runtime_smoke.sh" \
    "$SCRIPT_DIR/run_real_mpv_smoke.sh"; do
    if [[ -x "$candidate" ]]; then
      printf '%s' "$candidate"
      return 0
    fi
  done
  return 1
}

acquire_real_runtime_lease() {
  local lease_root="$REPO_ROOT/.tmp/skill-verification/locks"
  local lease_dir="$lease_root/exclusive-real-runtime"
  mkdir -p "$lease_root"
  if mkdir "$lease_dir" 2>/dev/null; then
    REAL_RUNTIME_LEASE_DIR="$lease_dir"
    printf '%s\n' "$SESSION_ID" >"$lease_dir/session_id"
    return 0
  fi

  local owner=""
  if [[ -f "$lease_dir/session_id" ]]; then
    owner=$(cat "$lease_dir/session_id")
  fi
  REAL_RUNTIME_LEASE_ERROR="real-runtime lease already held${owner:+ by $owner}"
  return 1
}

release_real_runtime_lease() {
  if [[ -n "$REAL_RUNTIME_LEASE_DIR" && -d "$REAL_RUNTIME_LEASE_DIR" ]]; then
    if [[ -f "$REAL_RUNTIME_LEASE_DIR/session_id" ]]; then
      local owner
      owner=$(cat "$REAL_RUNTIME_LEASE_DIR/session_id")
      if [[ "$owner" != "$SESSION_ID" ]]; then
        return 0
      fi
    fi
    rm -rf "$REAL_RUNTIME_LEASE_DIR"
  fi
}

compute_final_status() {
  if [[ "$FAILED" == "1" ]]; then
    FINAL_STATUS="failed"
  elif [[ "$BLOCKED" == "1" ]]; then
    FINAL_STATUS="blocked"
  elif [[ "$EXECUTED_REAL_STEPS" == "1" ]]; then
    FINAL_STATUS="passed"
  else
    FINAL_STATUS="skipped"
  fi
}

write_summary_files() {
  local lane_lines
  lane_lines=$(printf '%s\n' "${SELECTED_LANES[@]}")
  printf '%s\n' "$lane_lines" >"$ARTIFACT_DIR/lanes.txt"
  printf '%s\n' "${BLOCKERS[@]}" >"$ARTIFACT_DIR/blockers.txt"
  printf '%s\n' "${PATH_ARGS[@]}" >"$ARTIFACT_DIR/requested-paths.txt"

  ARTIFACT_DIR_ENV="$ARTIFACT_DIR" \
  SESSION_ID_ENV="$SESSION_ID" \
  FINAL_STATUS_ENV="$FINAL_STATUS" \
  PATH_SELECTION_MODE_ENV="$PATH_SELECTION_MODE" \
  ALLOW_REAL_RUNTIME_ENV="$ALLOW_REAL_RUNTIME" \
  SESSION_HOME_ENV="$SESSION_HOME" \
  SESSION_XDG_CONFIG_HOME_ENV="$SESSION_XDG_CONFIG_HOME" \
  SESSION_MPV_DIR_ENV="$SESSION_MPV_DIR" \
  SESSION_LOGS_DIR_ENV="$SESSION_LOGS_DIR" \
  SESSION_MPV_LOG_ENV="$SESSION_MPV_LOG" \
  STARTED_AT_ENV="$STARTED_AT" \
  FINISHED_AT_ENV="$FINISHED_AT" \
  FAILED_ENV="$FAILED" \
  FAILURE_COMMAND_ENV="${FAILURE_COMMAND:-}" \
  FAILURE_STDOUT_ENV="${FAILURE_STDOUT:-}" \
  FAILURE_STDERR_ENV="${FAILURE_STDERR:-}" \
  bun -e '
    const fs = require("fs");
    const path = require("path");

    const lines = fs
      .readFileSync(path.join(process.env.ARTIFACT_DIR_ENV, "steps.tsv"), "utf8")
      .trim()
      .split("\n")
      .filter(Boolean)
      .slice(1)
      .map((line) => {
        const [lane, name, status, exitCode, command, stdout, stderr, note] = line.split("\t");
        return { lane, name, status, exitCode: Number(exitCode), command, stdout, stderr, note };
      });

    const payload = {
      sessionId: process.env.SESSION_ID_ENV,
      startedAt: process.env.STARTED_AT_ENV,
      finishedAt: process.env.FINISHED_AT_ENV,
      status: process.env.FINAL_STATUS_ENV,
      pathSelectionMode: process.env.PATH_SELECTION_MODE_ENV,
      allowRealRuntime: process.env.ALLOW_REAL_RUNTIME_ENV === "1",
      sessionHome: process.env.SESSION_HOME_ENV,
      sessionXdgConfigHome: process.env.SESSION_XDG_CONFIG_HOME_ENV,
      sessionMpvDir: process.env.SESSION_MPV_DIR_ENV,
      sessionLogsDir: process.env.SESSION_LOGS_DIR_ENV,
      sessionMpvLog: process.env.SESSION_MPV_LOG_ENV,
      failed: process.env.FAILED_ENV === "1",
      failure: process.env.FAILURE_COMMAND_ENV
        ? {
            command: process.env.FAILURE_COMMAND_ENV,
            stdout: process.env.FAILURE_STDOUT_ENV,
            stderr: process.env.FAILURE_STDERR_ENV,
          }
        : null,
      blockers: fs
        .readFileSync(path.join(process.env.ARTIFACT_DIR_ENV, "blockers.txt"), "utf8")
        .split("\n")
        .filter(Boolean),
      lanes: fs
        .readFileSync(path.join(process.env.ARTIFACT_DIR_ENV, "lanes.txt"), "utf8")
        .split("\n")
        .filter(Boolean),
      requestedPaths: fs
        .readFileSync(path.join(process.env.ARTIFACT_DIR_ENV, "requested-paths.txt"), "utf8")
        .split("\n")
        .filter(Boolean),
      steps: lines,
    };

    fs.writeFileSync(
      path.join(process.env.ARTIFACT_DIR_ENV, "summary.json"),
      JSON.stringify(payload, null, 2) + "\n",
    );

    const summaryLines = [
      `status: ${payload.status}`,
      `session: ${payload.sessionId}`,
      `artifacts: ${process.env.ARTIFACT_DIR_ENV}`,
      `lanes: ${payload.lanes.join(", ") || "(none)"}`,
    ];

    if (payload.requestedPaths.length > 0) {
      summaryLines.push("requested paths:");
      for (const entry of payload.requestedPaths) {
        summaryLines.push(`- ${entry}`);
      }
    }

    if (payload.failure) {
      summaryLines.push(`failure command: ${payload.failure.command}`);
      summaryLines.push(`failure stdout: ${payload.failure.stdout}`);
      summaryLines.push(`failure stderr: ${payload.failure.stderr}`);
    }

    if (payload.blockers.length > 0) {
      summaryLines.push("blockers:");
      for (const blocker of payload.blockers) {
        summaryLines.push(`- ${blocker}`);
      }
    }

    summaryLines.push("steps:");
    for (const step of payload.steps) {
      summaryLines.push(`- ${step.lane}/${step.name}: ${step.status}`);
    }

    fs.writeFileSync(
      path.join(process.env.ARTIFACT_DIR_ENV, "summary.txt"),
      summaryLines.join("\n") + "\n",
    );
  '
}

SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
SKILL_DIR=$(cd "$SCRIPT_DIR/.." && pwd)
REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null || pwd)

declare -a PATH_ARGS=()
declare -a SELECTED_LANES=()
declare -a COMMANDS_RUN=()
declare -a BLOCKERS=()

ALLOW_REAL_RUNTIME=0
DRY_RUN=0
FAILED=0
BLOCKED=0
EXECUTED_REAL_STEPS=0
FAILURE_STEP=""
FAILURE_COMMAND=""
FAILURE_STDOUT=""
FAILURE_STDERR=""
REAL_RUNTIME_LEASE_DIR=""
REAL_RUNTIME_LEASE_ERROR=""
PATH_SELECTION_MODE="auto"

trap 'release_real_runtime_lease' EXIT

while [[ $# -gt 0 ]]; do
  case "$1" in
    --lane)
      shift
      [[ $# -gt 0 ]] || {
        echo "Missing value for --lane" >&2
        exit 2
      }
      add_lane "$1"
      PATH_SELECTION_MODE="explicit-lanes"
      ;;
    --artifact-dir)
      shift
      [[ $# -gt 0 ]] || {
        echo "Missing value for --artifact-dir" >&2
        exit 2
      }
      ARTIFACT_DIR=$1
      ;;
    --allow-real-runtime|--allow-real-gui)
      ALLOW_REAL_RUNTIME=1
      ;;
    --dry-run)
      DRY_RUN=1
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      PATH_ARGS+=("$1")
      ;;
  esac
  shift || true
done

if [[ -z "${ARTIFACT_DIR:-}" ]]; then
  SESSION_ID=$(generate_session_id)
  ARTIFACT_DIR="$REPO_ROOT/.tmp/skill-verification/$SESSION_ID"
else
  validate_artifact_dir "$ARTIFACT_DIR"
  SESSION_ID=$(basename "$ARTIFACT_DIR")
fi

mkdir -p "$ARTIFACT_DIR/steps"
STEPS_TSV="$ARTIFACT_DIR/steps.tsv"
printf 'lane\tstep\tstatus\texit_code\tcommand\tstdout\tstderr\tnote\n' >"$STEPS_TSV"

STARTED_AT=$(timestamp_iso)
SESSION_HOME="$REPO_ROOT/.tmp/skill-verification/runtime/$SESSION_ID/home"
SESSION_XDG_CONFIG_HOME="$REPO_ROOT/.tmp/skill-verification/runtime/$SESSION_ID/xdg-config"
SESSION_MPV_DIR="$SESSION_XDG_CONFIG_HOME/mpv"
SESSION_LOGS_DIR="$REPO_ROOT/.tmp/skill-verification/runtime/$SESSION_ID/logs"
SESSION_MPV_LOG="$SESSION_LOGS_DIR/mpv.log"
mkdir -p "$SESSION_HOME" "$SESSION_MPV_DIR" "$SESSION_LOGS_DIR"

CLASSIFIER_OUTPUT="$ARTIFACT_DIR/classification.txt"
if [[ ${#SELECTED_LANES[@]} -eq 0 ]]; then
  if [[ ${#PATH_ARGS[@]} -gt 0 ]]; then
    PATH_SELECTION_MODE="explicit-paths"
  fi
  if "$SCRIPT_DIR/classify_subminer_diff.sh" "${PATH_ARGS[@]}" >"$CLASSIFIER_OUTPUT"; then
    while IFS= read -r line; do
      case "$line" in
        lane:*)
          add_lane "${line#lane:}"
          ;;
      esac
    done <"$CLASSIFIER_OUTPUT"
  else
    record_failed_step "meta" "classify" "classification failed"
  fi
else
  : >"$CLASSIFIER_OUTPUT"
fi

record_env

if [[ ${#SELECTED_LANES[@]} -eq 0 ]]; then
  add_lane "core"
fi

for lane in "${SELECTED_LANES[@]}"; do
  case "$lane" in
    docs)
      run_step "$lane" "docs-kb" "bun run test:docs:kb" || break
      ;;
    config)
      run_step "$lane" "config" "bun run test:config" || break
      ;;
    core)
      run_step "$lane" "typecheck" "bun run typecheck" || break
      run_step "$lane" "fast-tests" "bun run test:fast" || break
      ;;
    launcher-plugin)
      run_step "$lane" "launcher" "bun run test:launcher" || break
      run_step "$lane" "plugin-src" "bun run test:plugin:src" || break
      ;;
    runtime-compat)
      run_step "$lane" "runtime-compat" "bun run test:runtime:compat" || break
      ;;
    real-runtime)
      if [[ "$ALLOW_REAL_RUNTIME" != "1" ]]; then
        record_blocked_step "$lane" "real-runtime" "real-runtime requested without --allow-real-runtime"
        continue
      fi
      if ! acquire_real_runtime_lease; then
        record_blocked_step "$lane" "real-runtime-lease" "$REAL_RUNTIME_LEASE_ERROR"
        continue
      fi
      helper=$(find_real_runtime_helper || true)
      if [[ -z "${helper:-}" ]]; then
        record_blocked_step "$lane" "real-runtime-helper" "no real-runtime helper script available in $SCRIPT_DIR"
        continue
      fi
      run_step "$lane" "real-runtime" "\"$helper\" \"$SESSION_ID\" \"$ARTIFACT_DIR\"" || break
      ;;
    *)
      record_blocked_step "$lane" "unknown-lane" "unknown lane: $lane"
      ;;
  esac
done

release_real_runtime_lease
FINISHED_AT=$(timestamp_iso)
compute_final_status
write_summary_files

printf 'summary:%s\n' "$ARTIFACT_DIR/summary.txt"
cat "$ARTIFACT_DIR/summary.txt"
