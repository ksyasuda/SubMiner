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
  local slug=${name//[^a-zA-Z0-9_-]/-}
  local stdout_rel="steps/${slug}.stdout.log"
  local stderr_rel="steps/${slug}.stderr.log"
  local stdout_path="$ARTIFACT_DIR/$stdout_rel"
  local stderr_path="$ARTIFACT_DIR/$stderr_rel"
  local status exit_code

  COMMANDS_RUN+=("$command")
  printf '%s\n' "$command" >"$ARTIFACT_DIR/steps/${slug}.command.txt"

  if [[ "$DRY_RUN" == "1" ]]; then
    printf '[dry-run] %s\n' "$command" >"$stdout_path"
    : >"$stderr_path"
    status="dry-run"
    exit_code=0
  else
    if bash -lc "cd \"$REPO_ROOT\" && $command" >"$stdout_path" 2>"$stderr_path"; then
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
  local slug=${name//[^a-zA-Z0-9_-]/-}
  local stdout_rel="steps/${slug}.stdout.log"
  local stderr_rel="steps/${slug}.stderr.log"
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
  FAILURE_STDOUT="steps/${2//[^a-zA-Z0-9_-]/-}.stdout.log"
  FAILURE_STDERR="steps/${2//[^a-zA-Z0-9_-]/-}.stderr.log"
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
  add_blocker "real-runtime lease already held${owner:+ by $owner}"
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

    function readLines(filePath) {
      if (!fs.existsSync(filePath)) return [];
      return fs.readFileSync(filePath, "utf8").split(/\r?\n/).filter(Boolean);
    }

    const artifactDir = process.env.ARTIFACT_DIR_ENV;
    const reportsDir = path.join(artifactDir, "reports");
    const lanes = readLines(path.join(artifactDir, "lanes.txt"));
    const blockers = readLines(path.join(artifactDir, "blockers.txt"));
    const requestedPaths = readLines(path.join(artifactDir, "requested-paths.txt"));
    const steps = readLines(path.join(artifactDir, "steps.tsv")).map((line) => {
      const [lane, name, status, exitCode, command, stdout, stderr, note] = line.split("\t");
      return {
        lane,
        name,
        status,
        exitCode: Number(exitCode || 0),
        command,
        stdout,
        stderr,
        note,
      };
    });
    const summary = {
      sessionId: process.env.SESSION_ID_ENV || "",
      artifactDir,
      reportsDir,
      status: process.env.FINAL_STATUS_ENV || "failed",
      selectedLanes: lanes,
      failed: process.env.FAILED_ENV === "1",
      failure:
        process.env.FAILED_ENV === "1"
          ? {
              command: process.env.FAILURE_COMMAND_ENV || "",
              stdout: process.env.FAILURE_STDOUT_ENV || "",
              stderr: process.env.FAILURE_STDERR_ENV || "",
            }
          : null,
      blockers,
      pathSelectionMode: process.env.PATH_SELECTION_MODE_ENV || "git-inferred",
      requestedPaths,
      allowRealRuntime: process.env.ALLOW_REAL_RUNTIME_ENV === "1",
      startedAt: process.env.STARTED_AT_ENV || "",
      finishedAt: process.env.FINISHED_AT_ENV || "",
      env: {
        home: process.env.SESSION_HOME_ENV || "",
        xdgConfigHome: process.env.SESSION_XDG_CONFIG_HOME_ENV || "",
        mpvDir: process.env.SESSION_MPV_DIR_ENV || "",
        logsDir: process.env.SESSION_LOGS_DIR_ENV || "",
        mpvLog: process.env.SESSION_MPV_LOG_ENV || "",
      },
      steps,
    };

    const summaryJson = JSON.stringify(summary, null, 2) + "\n";
    fs.writeFileSync(path.join(artifactDir, "summary.json"), summaryJson);
    fs.writeFileSync(path.join(reportsDir, "summary.json"), summaryJson);

    const lines = [];
    lines.push(`session_id: ${summary.sessionId}`);
    lines.push(`artifact_dir: ${artifactDir}`);
    lines.push(`selected_lanes: ${lanes.join(", ") || "(none)"}`);
    lines.push(`status: ${summary.status}`);
    lines.push(`path_selection_mode: ${summary.pathSelectionMode}`);
    if (requestedPaths.length > 0) {
      lines.push(`requested_paths: ${requestedPaths.join(", ")}`);
    }
    if (blockers.length > 0) {
      lines.push(`blockers: ${blockers.join(" | ")}`);
    }
    for (const step of steps) {
      lines.push(`${step.lane}/${step.name}: ${step.status}`);
      if (step.command) lines.push(`  command: ${step.command}`);
      lines.push(`  stdout: ${step.stdout}`);
      lines.push(`  stderr: ${step.stderr}`);
      if (step.note) lines.push(`  note: ${step.note}`);
    }
    if (summary.failed) {
      lines.push(`failure_command: ${process.env.FAILURE_COMMAND_ENV || ""}`);
    }
    const summaryText = lines.join("\n") + "\n";
    fs.writeFileSync(path.join(artifactDir, "summary.txt"), summaryText);
    fs.writeFileSync(path.join(reportsDir, "summary.txt"), summaryText);
  '
}

cleanup() {
  release_real_runtime_lease
}

CLASSIFIER_OUTPUT=""
ARTIFACT_DIR=""
ALLOW_REAL_RUNTIME=0
DRY_RUN=0
FAILED=0
BLOCKED=0
EXECUTED_REAL_STEPS=0
FINAL_STATUS=""
FAILURE_STEP=""
FAILURE_COMMAND=""
FAILURE_STDOUT=""
FAILURE_STDERR=""
REAL_RUNTIME_LEASE_DIR=""
STARTED_AT=""
FINISHED_AT=""

declare -a EXPLICIT_LANES=()
declare -a SELECTED_LANES=()
declare -a PATH_ARGS=()
declare -a COMMANDS_RUN=()
declare -a BLOCKERS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --lane)
      EXPLICIT_LANES+=("$(normalize_lane_name "$2")")
      shift 2
      ;;
    --artifact-dir)
      ARTIFACT_DIR=$2
      shift 2
      ;;
    --allow-real-runtime|--allow-real-gui)
      ALLOW_REAL_RUNTIME=1
      shift
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    --)
      shift
      while [[ $# -gt 0 ]]; do
        PATH_ARGS+=("$1")
        shift
      done
      ;;
    *)
      PATH_ARGS+=("$1")
      shift
      ;;
  esac
done

SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null || pwd)
SESSION_ID=$(generate_session_id)
PATH_SELECTION_MODE="git-inferred"
if [[ ${#PATH_ARGS[@]} -gt 0 ]]; then
  PATH_SELECTION_MODE="explicit"
fi

if [[ -z "$ARTIFACT_DIR" ]]; then
  mkdir -p "$REPO_ROOT/.tmp/skill-verification"
  ARTIFACT_DIR="$REPO_ROOT/.tmp/skill-verification/$SESSION_ID"
fi

SESSION_HOME="$ARTIFACT_DIR/home"
SESSION_XDG_CONFIG_HOME="$ARTIFACT_DIR/xdg"
SESSION_MPV_DIR="$ARTIFACT_DIR/mpv"
SESSION_LOGS_DIR="$ARTIFACT_DIR/logs"
SESSION_MPV_LOG="$SESSION_LOGS_DIR/mpv.log"

mkdir -p "$ARTIFACT_DIR/steps" "$ARTIFACT_DIR/reports" "$SESSION_HOME" "$SESSION_XDG_CONFIG_HOME" "$SESSION_MPV_DIR" "$SESSION_LOGS_DIR"
STEPS_TSV="$ARTIFACT_DIR/steps.tsv"
: >"$STEPS_TSV"

trap cleanup EXIT
STARTED_AT=$(timestamp_iso)

if [[ ${#EXPLICIT_LANES[@]} -gt 0 ]]; then
  local_lane=""
  for local_lane in "${EXPLICIT_LANES[@]}"; do
    add_lane "$local_lane"
  done
  printf 'reason:explicit lanes supplied\n' >"$ARTIFACT_DIR/classification.txt"
else
  if [[ ${#PATH_ARGS[@]} -gt 0 ]]; then
    CLASSIFIER_OUTPUT=$(bash "$SCRIPT_DIR/classify_subminer_diff.sh" "${PATH_ARGS[@]}")
  else
    CLASSIFIER_OUTPUT=$(bash "$SCRIPT_DIR/classify_subminer_diff.sh")
  fi
  printf '%s\n' "$CLASSIFIER_OUTPUT" >"$ARTIFACT_DIR/classification.txt"
  while IFS= read -r line; do
    case "$line" in
      lane:*)
        add_lane "${line#lane:}"
        ;;
    esac
  done <<<"$CLASSIFIER_OUTPUT"
fi

record_env

printf 'artifact_dir=%s\n' "$ARTIFACT_DIR"
printf 'selected_lanes=%s\n' "$(IFS=,; echo "${SELECTED_LANES[*]}")"

for lane in "${SELECTED_LANES[@]}"; do
  case "$lane" in
    docs)
      run_step "$lane" "docs-test" "bun run docs:test" || break
      [[ "$FAILED" == "1" ]] && break
      run_step "$lane" "docs-build" "bun run docs:build" || break
      ;;
    config)
      run_step "$lane" "test-config" "bun run test:config" || break
      ;;
    core)
      run_step "$lane" "typecheck" "bun run typecheck" || break
      [[ "$FAILED" == "1" ]] && break
      run_step "$lane" "test-fast" "bun run test:fast" || break
      ;;
    launcher-plugin)
      run_step "$lane" "launcher-smoke-src" "bun run test:launcher:smoke:src" || break
      [[ "$FAILED" == "1" ]] && break
      run_step "$lane" "plugin-src" "bun run test:plugin:src" || break
      ;;
    runtime-compat)
      run_step "$lane" "build" "bun run build" || break
      [[ "$FAILED" == "1" ]] && break
      run_step "$lane" "test-runtime-compat" "bun run test:runtime:compat" || break
      [[ "$FAILED" == "1" ]] && break
      run_step "$lane" "test-smoke-dist" "bun run test:smoke:dist" || break
      ;;
    real-runtime)
      if [[ "$PATH_SELECTION_MODE" != "explicit" ]]; then
        record_blocked_step \
          "$lane" \
          "real-runtime-guard" \
          "real-runtime lane requires explicit paths; inferred local git changes are non-authoritative"
        break
      fi

      if [[ "$ALLOW_REAL_RUNTIME" != "1" ]]; then
        record_blocked_step \
          "$lane" \
          "real-runtime-guard" \
          "real-runtime lane requested but --allow-real-runtime was not supplied"
        break
      fi

      if ! acquire_real_runtime_lease; then
        record_blocked_step \
          "$lane" \
          "real-runtime-lease" \
          "real-runtime lease already held; rerun after the active runtime verification finishes"
        break
      fi

      if ! REAL_RUNTIME_HELPER=$(find_real_runtime_helper); then
        record_blocked_step \
          "$lane" \
          "real-runtime-helper" \
          "real-runtime helper not implemented yet"
        break
      fi

      printf -v REAL_RUNTIME_COMMAND \
        'SESSION_ID=%q HOME=%q XDG_CONFIG_HOME=%q SUBMINER_MPV_LOG=%q bash %q' \
        "$SESSION_ID" \
        "$SESSION_HOME" \
        "$SESSION_XDG_CONFIG_HOME" \
        "$SESSION_MPV_LOG" \
        "$REAL_RUNTIME_HELPER"

      run_step "$lane" "real-runtime-smoke" "$REAL_RUNTIME_COMMAND" || break
      ;;
    *)
      record_failed_step "$lane" "lane-validation" "unknown lane: $lane"
      break
      ;;
  esac

  if [[ "$FAILED" == "1" || "$BLOCKED" == "1" ]]; then
    break
  fi
done

FINISHED_AT=$(timestamp_iso)
compute_final_status
write_summary_files

printf 'status=%s\n' "$FINAL_STATUS"
printf 'artifact_dir=%s\n' "$ARTIFACT_DIR"

case "$FINAL_STATUS" in
  failed)
    printf 'result=failed\n'
    printf 'failure_command=%s\n' "$FAILURE_COMMAND"
    exit 1
    ;;
  blocked)
    printf 'result=blocked\n'
    exit 2
    ;;
  *)
    printf 'result=ok\n'
    exit 0
    ;;
esac
