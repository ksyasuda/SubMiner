import { syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore } from '../../core/services';
import type { ResolvedConfig } from '../../types';
import {
  getPreferredYomitanAnkiServerUrl as getPreferredYomitanAnkiServerUrlRuntime,
  shouldForceOverrideYomitanAnkiServer,
} from './yomitan-anki-server';

export interface YomitanAnkiServerSyncRuntimeDeps {
  isExternalReadOnlyMode: () => boolean;
  getResolvedConfig: () => ResolvedConfig;
  getYomitanParserRuntimeDeps: () => Parameters<typeof syncYomitanDefaultAnkiServerCore>[1];
  logError: (message: string, ...args: unknown[]) => void;
  logInfo: (message: string, ...args: unknown[]) => void;
}

export function createYomitanAnkiServerSyncRuntime(deps: YomitanAnkiServerSyncRuntimeDeps): {
  syncYomitanDefaultProfileAnkiServer: () => Promise<void>;
} {
  let lastSyncedYomitanAnkiSettingsKey: string | null = null;

  function getPreferredYomitanAnkiServerUrl(): string {
    return getPreferredYomitanAnkiServerUrlRuntime(deps.getResolvedConfig().ankiConnect);
  }

  async function syncYomitanDefaultProfileAnkiServer(): Promise<void> {
    if (deps.isExternalReadOnlyMode()) {
      return;
    }

    const targetUrl = getPreferredYomitanAnkiServerUrl().trim();
    const ankiConnectConfig = deps.getResolvedConfig().ankiConnect;
    const targetDeck = ankiConnectConfig?.deck?.trim() ?? '';
    const targetSettingsKey = `${targetUrl}\n${targetDeck}`;
    if (!targetUrl || targetSettingsKey === lastSyncedYomitanAnkiSettingsKey) {
      return;
    }

    const synced = await syncYomitanDefaultAnkiServerCore(
      targetUrl,
      deps.getYomitanParserRuntimeDeps(),
      {
        error: (message, ...args) => {
          deps.logError(message, ...args);
        },
        info: (message, ...args) => {
          deps.logInfo(message, ...args);
        },
      },
      {
        forceOverride: ankiConnectConfig
          ? shouldForceOverrideYomitanAnkiServer(ankiConnectConfig)
          : false,
        deck: targetDeck,
      },
    );

    if (synced) {
      lastSyncedYomitanAnkiSettingsKey = targetSettingsKey;
    }
  }

  return { syncYomitanDefaultProfileAnkiServer };
}
