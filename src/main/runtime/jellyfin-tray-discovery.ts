import { i18n } from '../../i18n/index.js';

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
  return Boolean(jellyfin.enabled !== false && jellyfin.serverUrl);
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
  options: { desiredActive?: boolean } = {},
): Promise<void> {
  try {
    const activeSession = deps.getRemoteSession();
    if (options.desiredActive === false) {
      if (activeSession) {
        deps.stopRemoteSession();
        deps.logger.info('Jellyfin discovery stopped.');
        deps.showMpvOsd(i18n.t('osd.jellyfinStopped'));
      }
      return;
    }

    if (activeSession) {
      let visible = false;
      try {
        visible = await activeSession.advertiseNow();
      } catch {
        deps.logger.warn('Jellyfin discovery visibility check failed; restarting.');
      }

      if (visible) {
        if (options.desiredActive === true) {
          deps.logger.info('Jellyfin discovery already active.');
        } else {
          deps.stopRemoteSession();
          deps.logger.info('Jellyfin discovery stopped.');
          deps.showMpvOsd(i18n.t('osd.jellyfinStopped'));
        }
        return;
      }

      deps.logger.warn('Jellyfin discovery was active but not visible; restarting.');
      deps.stopRemoteSession();
    }

    await deps.startRemoteSession({ explicit: true });
    const remoteSession = deps.getRemoteSession();
    if (!remoteSession) {
      deps.logger.warn('Jellyfin discovery could not start. Configure Jellyfin first.');
      deps.showMpvOsd(i18n.t('osd.jellyfinUnavailable'));
      return;
    }

    const visible = await remoteSession.advertiseNow();
    if (visible) {
      deps.logger.info('Jellyfin discovery started; cast target is visible in server sessions.');
      deps.showMpvOsd(i18n.t('osd.jellyfinStarted'));
    } else {
      deps.logger.warn('Jellyfin discovery started, but cast target is not visible yet.');
      deps.showMpvOsd(i18n.t('osd.jellyfinWaiting'));
    }
  } catch (error) {
    deps.logger.error('Failed to toggle Jellyfin discovery.', error);
    deps.showMpvOsd(i18n.t('osd.jellyfinFailed'));
  } finally {
    deps.refreshTrayMenu();
  }
}
