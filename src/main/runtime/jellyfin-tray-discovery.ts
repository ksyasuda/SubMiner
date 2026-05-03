type JellyfinTrayConfig = {
  enabled?: boolean;
  serverUrl?: string | null;
  accessToken?: string | null;
  userId?: string | null;
};

type JellyfinTrayRemoteSession = {
  advertiseNow: () => Promise<boolean>;
};

type JellyfinTrayLogger = {
  info: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, error?: unknown) => void;
};

type JellyfinTrayDiscoveryDeps<TSession extends JellyfinTrayRemoteSession> = {
  getResolvedJellyfinConfig: () => JellyfinTrayConfig;
  getRemoteSession: () => TSession | null;
  clearStoredSession: () => void;
  stopRemoteSession: () => void;
  startRemoteSession: (options: { explicit: true }) => Promise<void>;
  refreshTrayMenu: () => void;
  logger: JellyfinTrayLogger;
  showMpvOsd: (message: string) => void;
};

export function isJellyfinConfiguredForTray(
  deps: Pick<JellyfinTrayDiscoveryDeps<JellyfinTrayRemoteSession>, 'getResolvedJellyfinConfig'>,
): boolean {
  const jellyfin = deps.getResolvedJellyfinConfig();
  return Boolean(jellyfin.enabled && jellyfin.serverUrl && jellyfin.accessToken && jellyfin.userId);
}

export function clearJellyfinAuthSessionAndRefreshTray<TSession extends JellyfinTrayRemoteSession>(
  deps: Pick<
    JellyfinTrayDiscoveryDeps<TSession>,
    'clearStoredSession' | 'getRemoteSession' | 'stopRemoteSession' | 'refreshTrayMenu' | 'logger'
  >,
): void {
  try {
    deps.clearStoredSession();
  } catch (error) {
    deps.logger.error('Failed to clear Jellyfin auth session.', error);
  }

  try {
    if (deps.getRemoteSession()) {
      deps.stopRemoteSession();
    }
  } catch (error) {
    deps.logger.error('Failed to stop Jellyfin discovery while clearing auth session.', error);
  } finally {
    deps.refreshTrayMenu();
  }
}

export async function toggleJellyfinDiscoveryFromTray<TSession extends JellyfinTrayRemoteSession>(
  deps: Pick<
    JellyfinTrayDiscoveryDeps<TSession>,
    | 'getRemoteSession'
    | 'stopRemoteSession'
    | 'startRemoteSession'
    | 'refreshTrayMenu'
    | 'logger'
    | 'showMpvOsd'
  >,
): Promise<void> {
  try {
    const activeSession = deps.getRemoteSession();
    if (activeSession) {
      deps.stopRemoteSession();
      deps.logger.info('Jellyfin discovery stopped.');
      deps.showMpvOsd('Jellyfin discovery stopped');
      return;
    }

    await deps.startRemoteSession({ explicit: true });
    const remoteSession = deps.getRemoteSession();
    if (!remoteSession) {
      deps.logger.warn('Jellyfin discovery could not start. Configure Jellyfin first.');
      deps.showMpvOsd('Jellyfin discovery unavailable');
      return;
    }

    const visible = await remoteSession.advertiseNow();
    if (visible) {
      deps.logger.info('Jellyfin discovery started; cast target is visible in server sessions.');
      deps.showMpvOsd('Jellyfin discovery started');
    } else {
      deps.logger.warn('Jellyfin discovery started, but cast target is not visible yet.');
      deps.showMpvOsd('Jellyfin discovery started; waiting for visibility');
    }
  } catch (error) {
    deps.logger.error('Failed to toggle Jellyfin discovery.', error);
    deps.showMpvOsd('Jellyfin discovery failed');
  } finally {
    deps.refreshTrayMenu();
  }
}
