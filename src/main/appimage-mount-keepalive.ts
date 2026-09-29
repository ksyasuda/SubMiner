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
//
// Only a runtime-owned FUSE mount has that lifetime problem. Sandboxes such as
// `firejail --appimage` (used by the AppImage catalog test) loop-mount the image
// for the sandbox's lifetime and set NoNewPrivs, so FUSE cannot mount there and
// the supervisor could never start the detached app.

import fs from 'node:fs';

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

function unescapeMountInfoPath(value: string): string {
  return value.replace(/\\([0-7]{3})/g, (_, octal: string) =>
    String.fromCharCode(parseInt(octal, 8)),
  );
}

// Filesystem type of the mount containing `filePath`, from /proc/<pid>/mountinfo
// text. Null when no mount matches.
export function resolveMountFsType(filePath: string, mountInfo: string): string | null {
  let best: { mountPoint: string; fsType: string } | null = null;
  for (const line of mountInfo.split('\n')) {
    const [mountFields, fsFields] = line.split(' - ');
    const mountPointField = mountFields?.split(' ')[4];
    const fsType = fsFields?.split(' ')[0];
    if (!mountPointField || !fsType) continue;
    const mountPoint = unescapeMountInfoPath(mountPointField);
    const contains =
      mountPoint === '/' || filePath === mountPoint || filePath.startsWith(`${mountPoint}/`);
    if (contains && (!best || mountPoint.length >= best.mountPoint.length)) {
      best = { mountPoint, fsType };
    }
  }
  return best?.fsType ?? null;
}

function readSelfMountInfo(): string | null {
  try {
    return fs.readFileSync('/proc/self/mountinfo', 'utf8');
  } catch {
    return null;
  }
}

export function resolveAppImageMountKeepaliveInvocation(
  env: NodeJS.ProcessEnv,
  platform: NodeJS.Platform = process.platform,
  execPath: string = process.execPath,
  readMountInfo: () => string | null = readSelfMountInfo,
): AppImageMountKeepaliveInvocation | null {
  if (platform !== 'linux') return null;
  if (env[DISABLE_ENV] === '1') return null;
  const appImagePath = env.APPIMAGE?.trim();
  if (!appImagePath) return null;
  // Keep the supervisor when the mount type is unknown; skip it only when we
  // positively run from a non-FUSE mount (sandbox loop mount, extracted image).
  const mountInfo = readMountInfo();
  const fsType = mountInfo === null ? null : resolveMountFsType(execPath, mountInfo);
  if (fsType !== null && !fsType.startsWith('fuse')) return null;
  return {
    command: '/bin/sh',
    args: ['-c', APPIMAGE_MOUNT_KEEPALIVE_SCRIPT, APPIMAGE_MOUNT_KEEPALIVE_LABEL, appImagePath],
  };
}
