import { execFile } from 'child_process';
import electron from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { createLogger } from '../../logger';

const { Notification, nativeImage } = electron;
const logger = createLogger('core:notification');

export function resolveDefaultNotificationIconPath(deps: {
  platform: string;
  resourcesPath: string;
  appPath: string;
  dirname: string;
  cwd: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
}): string | null {
  const iconNames =
    deps.platform === 'win32'
      ? ['SubMiner.ico', 'SubMiner-square.png', 'SubMiner.png']
      : ['SubMiner-square.png', 'SubMiner.png'];

  const baseDirs = [
    deps.joinPath(deps.resourcesPath, 'assets'),
    deps.joinPath(deps.appPath, 'assets'),
    deps.joinPath(deps.dirname, '..', 'assets'),
    deps.joinPath(deps.dirname, '..', '..', 'assets'),
    deps.joinPath(deps.cwd, 'assets'),
  ];

  for (const baseDir of baseDirs) {
    for (const iconName of iconNames) {
      const candidate = deps.joinPath(baseDir, iconName);
      if (deps.fileExists(candidate)) {
        return candidate;
      }
    }
  }

  return null;
}

function resolveRuntimeDefaultNotificationIconPath(): string | null {
  return resolveDefaultNotificationIconPath({
    platform: process.platform,
    resourcesPath: process.resourcesPath,
    appPath: electron.app?.getAppPath?.() ?? process.cwd(),
    dirname: __dirname,
    cwd: process.cwd(),
    joinPath: (...parts) => path.join(...parts),
    fileExists: (candidate) => fs.existsSync(candidate),
  });
}

/**
 * Live Electron notifications keyed by `replaceId`, for platforms without true in-place
 * replacement (and as the Linux fallback when notify-send is unusable). Electron exposes no native
 * "replace this notification" flag, so a repeated status closes its predecessor instead of
 * stacking a fresh toast per update.
 */
const notificationsByReplaceId = new Map<string, Electron.Notification>();

/** Untracks first, so the notification's own `close` handler cannot race a replacement into it. */
function closeTrackedElectronNotification(replaceId: string): void {
  const tracked = notificationsByReplaceId.get(replaceId);
  if (!tracked) return;
  notificationsByReplaceId.delete(replaceId);
  tracked.close();
}

/** The freedesktop body is markup; unescaped `&`/`<` in an anime title would corrupt or drop it. */
function escapeFreedesktopNotificationBody(body: string): string {
  return body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

type NotifySendExec = (
  args: string[],
  callback: (error: Error | null, stdout: string) => void,
) => void;

type NotifySendNotification = { title: string; body?: string; iconPath?: string };

type NotifySendPending = { notification: NotifySendNotification; fallback: () => void };

type NotifySendEntry = {
  dbusId: string | null;
  sending: boolean;
  pending: NotifySendPending | null;
};

/**
 * A sick daemon should not make every later update wait out the spawn timeout first, so a run of
 * failures still gives up on notify-send even when none of them is a missing binary.
 */
const NOTIFY_SEND_MAX_CONSECUTIVE_FAILURES = 3;

/**
 * A missing or non-executable binary never fixes itself, so notify-send is abandoned on the spot.
 * Every other failure (a bad exit code, a kill signal, or a transient spawn error like `EMFILE`
 * under fd pressure) goes through the consecutive-failure threshold instead.
 */
const NOTIFY_SEND_PERMANENT_ERROR_CODES = new Set(['ENOENT', 'EACCES', 'EPERM']);

function isNotifySendUnusable(error: unknown): boolean {
  const code = (error as NodeJS.ErrnoException | null)?.code;
  return typeof code === 'string' && NOTIFY_SEND_PERMANENT_ERROR_CODES.has(code);
}

/**
 * In-place notification replacement for Linux. Electron cannot reuse a freedesktop notification id,
 * so its close+show fallback makes every progress update flicker off-screen and back. notify-send
 * `--replace-id` updates the existing popup statically instead. The daemon-assigned id comes from
 * `--print-id`, so only one send per replaceId is in flight at a time and a fast progress stream
 * cannot race the id capture. Updates arriving mid-send collapse to the latest one, since a stale
 * progress message is never worth showing. Any failed update falls back to the Electron path, and
 * notify-send is abandoned for good once it looks unusable rather than merely unlucky.
 */
export function createNotifySendReplacer(
  execNotifySend: NotifySendExec,
): (replaceId: string, notification: NotifySendNotification, fallback: () => void) => void {
  const stateByReplaceId = new Map<string, NotifySendEntry>();
  let unavailable = false;
  let consecutiveFailures = 0;

  const flush = (entry: NotifySendEntry): void => {
    if (entry.sending) return;
    const next = entry.pending;
    if (!next) return;
    entry.pending = null;
    if (unavailable) {
      next.fallback();
      return;
    }

    const args = ['--app-name=SubMiner', '--print-id'];
    if (entry.dbusId) {
      args.push(`--replace-id=${entry.dbusId}`);
    }
    if (next.notification.iconPath) {
      args.push(`--icon=${next.notification.iconPath}`);
    }
    args.push('--', next.notification.title);
    if (next.notification.body) {
      args.push(escapeFreedesktopNotificationBody(next.notification.body));
    }

    const finish = (): void => {
      entry.sending = false;
      flush(entry);
    };

    const handleFailure = (error: unknown, unusable: boolean): void => {
      consecutiveFailures += 1;
      if (unusable || consecutiveFailures >= NOTIFY_SEND_MAX_CONSECUTIVE_FAILURES) {
        unavailable = true;
        logger.warn('notify-send unusable; falling back to Electron notifications', error);
      } else {
        logger.warn('notify-send update failed; showing it as an Electron notification', error);
      }
      // A queued update has already superseded this one, so flushing it would flash a stale message
      // before the newer one renders. The pending update falls back on its own turn if needed.
      if (!entry.pending) {
        next.fallback();
      }
      finish();
    };

    entry.sending = true;
    // A synchronous throw would otherwise leave `sending` stuck true and wedge the entry, silently
    // dropping every later update for this replaceId, so spawn failures are caught here too. It
    // also means the call itself is malformed, which will not fix itself on the next update.
    try {
      execNotifySend(args, (error, stdout) => {
        if (error) {
          handleFailure(error, isNotifySendUnusable(error));
          return;
        }
        consecutiveFailures = 0;
        const id = stdout.trim();
        if (/^[1-9]\d*$/.test(id)) {
          entry.dbusId = id;
        }
        finish();
      });
    } catch (error) {
      handleFailure(error, true);
    }
  };

  return (replaceId, notification, fallback) => {
    if (unavailable) {
      fallback();
      return;
    }
    let entry = stateByReplaceId.get(replaceId);
    if (!entry) {
      entry = { dbusId: null, sending: false, pending: null };
      stateByReplaceId.set(replaceId, entry);
    }
    entry.pending = { notification, fallback };
    flush(entry);
  };
}

const showLinuxReplaceableNotification = createNotifySendReplacer((args, callback) =>
  execFile('notify-send', args, { timeout: 5_000 }, (error, stdout) =>
    callback(error, stdout ?? ''),
  ),
);

export function showDesktopNotification(
  title: string,
  options: { body?: string; icon?: string; replaceId?: string },
): void {
  const notificationOptions: {
    title: string;
    body?: string;
    icon?: Electron.NativeImage | string;
  } = { title };

  if (options.body) {
    notificationOptions.body = options.body;
  }

  const icon = options.icon ?? resolveRuntimeDefaultNotificationIconPath() ?? undefined;

  if (icon) {
    const isFilePath =
      typeof icon === 'string' && (icon.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(icon));

    if (isFilePath) {
      if (fs.existsSync(icon)) {
        notificationOptions.icon = icon;
      } else {
        logger.warn('Notification icon file not found', icon);
      }
    } else if (typeof icon === 'string' && icon.startsWith('data:image/')) {
      const base64Data = icon.replace(/^data:image\/\w+;base64,/, '');
      try {
        const image = nativeImage.createFromBuffer(Buffer.from(base64Data, 'base64'));
        if (image.isEmpty()) {
          logger.warn(
            'Notification icon created from base64 is empty - image format may not be supported by Electron',
          );
        } else {
          notificationOptions.icon = image;
        }
      } catch (err) {
        logger.error('Failed to create notification icon from base64', err);
      }
    } else {
      notificationOptions.icon = icon;
    }
  }

  const replaceId = options.replaceId?.trim();
  const showElectronNotification = (): void => {
    const notification = new Notification(notificationOptions);
    if (replaceId) {
      closeTrackedElectronNotification(replaceId);
      notificationsByReplaceId.set(replaceId, notification);
      notification.once('close', () => {
        if (notificationsByReplaceId.get(replaceId) === notification) {
          notificationsByReplaceId.delete(replaceId);
        }
      });
    }
    notification.show();
  };

  // notify-send takes a path or a theme name, so a base64 icon only exists as an in-memory
  // NativeImage. Those keep the Electron path and their icon rather than losing it to in-place
  // replacement.
  const iconPath =
    typeof notificationOptions.icon === 'string' ? notificationOptions.icon : undefined;
  const iconNeedsElectron = notificationOptions.icon !== undefined && iconPath === undefined;

  if (replaceId && process.platform === 'linux' && !iconNeedsElectron) {
    // An earlier update for this id may have rendered through Electron (a transient notify-send
    // failure, or a NativeImage icon). That toast is not the one notify-send replaces, so it would
    // sit on screen next to the updated popup.
    closeTrackedElectronNotification(replaceId);
    showLinuxReplaceableNotification(
      replaceId,
      { title, body: options.body, iconPath },
      showElectronNotification,
    );
    return;
  }

  showElectronNotification();
}
