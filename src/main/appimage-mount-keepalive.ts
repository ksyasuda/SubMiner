// Background AppImage launches must not execute directly from a FUSE mount whose
// lifetime is owned by a short-lived helper's AppImage runtime: when the app later
// quits, the runtime unmounts the squashfs while Chromium utility children (network
// service et al.) are still mid-shutdown, and they die with SIGBUS on their mmapped
// executable — surfacing as "Service Crash" desktop notifications on every quit.
//
// Fix: detach a tiny POSIX-sh supervisor instead of the raw process. It mounts the
// AppImage via `--appimage-mount` (holder process keeps the mount alive), runs
// AppRun from that mount, and after the app exits waits until no process is still
// executing from the mount before releasing the holder.

export interface AppImageMountKeepaliveInvocation {
  command: string;
  args: string[];
}

const DISABLE_ENV = 'SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE';

// $0 is set to this label so the supervisor is identifiable in `ps` output.
export const APPIMAGE_MOUNT_KEEPALIVE_LABEL = 'subminer-appimage-keepalive';

// POSIX sh only; every failure path falls back to executing the AppImage directly,
// which is exactly the pre-wrapper behavior.
export const APPIMAGE_MOUNT_KEEPALIVE_SCRIPT = `
set -u
appimage=$1
shift
run_direct() {
  exec "$appimage" "$@"
}
fifo=$(mktemp -u "\${TMPDIR:-/tmp}/subminer-appimage-mount-XXXXXX") || run_direct "$@"
mkfifo "$fifo" || run_direct "$@"
"$appimage" --appimage-mount >"$fifo" 2>/dev/null &
holder=$!
mount_point=""
IFS= read -r mount_point <"$fifo" || true
rm -f "$fifo"
app_run=""
if [ -n "$mount_point" ] && [ -x "$mount_point/AppRun" ]; then
  app_run="$mount_point/AppRun"
fi
if [ -z "$app_run" ]; then
  kill "$holder" 2>/dev/null
  run_direct "$@"
fi
APPDIR="$mount_point" "$app_run" "$@"
rc=$?
# Do not release the mount while any process still executes from it; releasing
# early SIGBUSes Chromium children that are mid-shutdown.
tries=0
while [ "$tries" -lt 100 ]; do
  if readlink /proc/[0-9]*/exe 2>/dev/null | grep -qF "$mount_point/"; then
    sleep 0.1
    tries=$((tries + 1))
  else
    break
  fi
done
kill "$holder" 2>/dev/null
exit "$rc"
`;

export function resolveAppImageMountKeepaliveInvocation(
  env: NodeJS.ProcessEnv,
  platform: NodeJS.Platform = process.platform,
): AppImageMountKeepaliveInvocation | null {
  if (platform !== 'linux') return null;
  if (env[DISABLE_ENV] === '1') return null;
  const appImagePath = env.APPIMAGE?.trim();
  if (!appImagePath) return null;
  return {
    command: '/bin/sh',
    args: ['-c', APPIMAGE_MOUNT_KEEPALIVE_SCRIPT, APPIMAGE_MOUNT_KEEPALIVE_LABEL, appImagePath],
  };
}
