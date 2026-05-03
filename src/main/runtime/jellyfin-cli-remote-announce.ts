import type { CliArgs } from '../../cli/args';

type JellyfinRemoteSession = {
  advertiseNow: () => Promise<boolean>;
};

export function createHandleJellyfinRemoteAnnounceCommand(deps: {
  startJellyfinRemoteSession: (options?: { explicit?: boolean }) => Promise<void>;
  getRemoteSession: () => JellyfinRemoteSession | null;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
}) {
  return async (args: CliArgs): Promise<boolean> => {
    if (!args.jellyfinRemoteAnnounce) {
      return false;
    }

    await deps.startJellyfinRemoteSession({ explicit: true });
    const remoteSession = deps.getRemoteSession();
    if (!remoteSession) {
      deps.logWarn('Jellyfin remote session is not available.');
      return true;
    }

    const visible = await remoteSession.advertiseNow();
    if (visible) {
      deps.logInfo('Jellyfin cast target is visible in server sessions.');
    } else {
      deps.logWarn(
        'Jellyfin remote announce sent, but cast target is not visible in server sessions yet.',
      );
    }
    return true;
  };
}
