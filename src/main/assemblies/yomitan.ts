import type { BrowserWindow, Extension, IpcMain, Session } from 'electron';
import {
  clearYomitanParserCachesForWindow,
  destroyYomitanSettingsWindow,
  getYomitanCurrentAnkiDeckName as getYomitanCurrentAnkiDeckNameCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  openYomitanSettingsWindow,
} from '../../core/services';
import { generateSentenceFurigana } from '../../core/services/tokenizer/sentence-furigana';
import type { Logger } from '../../logger';
import {
  DICTIONARY_EXTERNAL_LINK_CHANNEL,
  parseDictionaryExternalUrl,
} from '../../shared/dictionary-external-link';
import type { ResolvedConfig } from '../../types';
import type { DictionaryBackend } from '../../types/config';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createYomitanAnkiServerSyncRuntime } from '../runtime/yomitan-anki-server-sync';
import { reloadOverlayWindowsForYomitanContentScripts } from '../runtime/yomitan-extension-overlay-reload';
import { createYomitanExtensionRuntime } from '../runtime/yomitan-extension-runtime';
import { createYomitanProfilePolicy } from '../runtime/yomitan-profile-policy';
import { createYomitanSettingsRuntime } from '../runtime/yomitan-settings-runtime';
import type { AppState } from '../state';

/** The bundled Hachidori extension. Injected because it reaches Electron's session API. */
export interface HachidoriExtensionLike {
  ensureLoaded: () => Promise<Extension>;
  getSession: () => Session;
}

/**
 * What the dictionary-extension domain needs from the rest of the app. Siblings built after
 * this assembly in `main.ts` are reached through callbacks.
 */
export interface YomitanAssemblyDeps {
  appState: Pick<
    AppState,
    | 'yomitanExt'
    | 'yomitanSession'
    | 'yomitanParserWindow'
    | 'yomitanParserReadyPromise'
    | 'yomitanParserInitPromise'
    | 'yomitanSettingsWindow'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'info' | 'warn' | 'error' | 'debug'>;
  userDataPath: string;
  /** Fixed for the process lifetime; backend changes apply on restart. */
  dictionaryBackend: DictionaryBackend;
  hachidori: HachidoriExtensionLike;
  getOverlayWindows: () => BrowserWindow[];
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification'
  >;
  ipcMain: Pick<IpcMain, 'handle'>;
  openExternal: (url: string) => Promise<void>;
}

export type YomitanAssembly = ReturnType<typeof createYomitanAssembly>;

/**
 * Owns the Yomitan/Hachidori extension: loading, the parser runtime deps every tokenizer
 * caller shares, the external-profile policy, and the dictionary settings windows. The
 * inactive backend's settings window stays openable so users can manage both.
 */
export function createYomitanAssembly(deps: YomitanAssemblyDeps) {
  const { appState, logger, dictionaryBackend, hachidori } = deps;
  let yomitanLoadInFlight: Promise<Extension | null> | null = null;
  let hachidoriSettingsWindow: BrowserWindow | null = null;
  let inactiveYomitanExtension: Extension | null = null;
  let inactiveYomitanSettingsWindow: BrowserWindow | null = null;
  let inactiveYomitanLoad: Promise<Extension | null> | null = null;

  const profilePolicy = createYomitanProfilePolicy({
    externalProfilePath:
      dictionaryBackend === 'yomitan' ? deps.getConfig().yomitan.externalProfilePath : '',
    logInfo: (message) => logger.info(message),
  });

  function getParserRuntimeDeps() {
    return {
      getYomitanExt: () => appState.yomitanExt,
      getYomitanSession: () => appState.yomitanSession,
      getYomitanParserWindow: () => appState.yomitanParserWindow,
      setYomitanParserWindow: (window: BrowserWindow | null) => {
        appState.yomitanParserWindow = window;
      },
      getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (promise: Promise<void> | null) => {
        appState.yomitanParserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
      setYomitanParserInitPromise: (promise: Promise<boolean> | null) => {
        appState.yomitanParserInitPromise = promise;
      },
    };
  }

  const extensionRuntime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async (loadDeps) => {
      if (dictionaryBackend === 'yomitan') return loadYomitanExtensionCore(loadDeps);
      try {
        const extension = await hachidori.ensureLoaded();
        loadDeps.setYomitanExtension(extension);
        loadDeps.setYomitanSession(hachidori.getSession());
        return extension;
      } catch (error) {
        logger.error('Failed to load Hachidori extension:', error);
        loadDeps.setYomitanExtension(null);
        loadDeps.setYomitanSession(null);
        return null;
      }
    },
    userDataPath: deps.userDataPath,
    externalProfilePath: profilePolicy.externalProfilePath,
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    setYomitanParserWindow: (window) => {
      appState.yomitanParserWindow = window as BrowserWindow | null;
    },
    setYomitanParserReadyPromise: (promise) => {
      appState.yomitanParserReadyPromise = promise;
    },
    setYomitanParserInitPromise: (promise) => {
      appState.yomitanParserInitPromise = promise;
    },
    setYomitanExtension: (extension) => {
      appState.yomitanExt = extension;
    },
    setYomitanSession: (nextSession) => {
      appState.yomitanSession = nextSession;
    },
    getYomitanExtension: () => appState.yomitanExt,
    getLoadInFlight: () => yomitanLoadInFlight,
    setLoadInFlight: (promise) => {
      yomitanLoadInFlight = promise;
    },
    onYomitanExtensionLoaded: () => {
      const reloaded = reloadOverlayWindowsForYomitanContentScripts(
        deps.getOverlayWindows(),
        (message, error) => logger.warn(message, error),
      );
      if (reloaded > 0) {
        logger.debug(`Reloaded ${reloaded} overlay window(s) after Yomitan extension load.`);
      }
    },
  });

  const { syncYomitanDefaultProfileAnkiServer } = createYomitanAnkiServerSyncRuntime({
    isExternalReadOnlyMode: () => profilePolicy.isExternalReadOnlyMode(),
    getResolvedConfig: () => deps.getConfig(),
    getYomitanParserRuntimeDeps: () => getParserRuntimeDeps(),
    logError: (message, ...args) => {
      logger.error(message, ...args);
    },
    logInfo: (message, ...args) => {
      logger.info(message, ...args);
    },
  });

  async function loadExtension(): Promise<Extension | null> {
    const extension = await extensionRuntime.loadYomitanExtension();
    if (extension && !profilePolicy.isExternalReadOnlyMode()) {
      await syncYomitanDefaultProfileAnkiServer();
    }
    return extension;
  }

  async function ensureExtensionLoaded(): Promise<Extension | null> {
    const extension = await extensionRuntime.ensureYomitanExtensionLoaded();
    if (extension && !profilePolicy.isExternalReadOnlyMode()) {
      await syncYomitanDefaultProfileAnkiServer();
    }
    return extension;
  }

  const { openYomitanSettings: openYomitanSettingsHandler } = createYomitanSettingsRuntime({
    ensureYomitanExtensionLoaded: () => ensureExtensionLoaded(),
    getYomitanExtension: () => appState.yomitanExt,
    getYomitanExtensionLoadInFlight: () => yomitanLoadInFlight,
    getYomitanSession: () => appState.yomitanSession,
    openYomitanSettingsWindow: ({ yomitanExt, getExistingWindow, setWindow, yomitanSession }) => {
      openYomitanSettingsWindow({
        yomitanExt: yomitanExt as Extension,
        getExistingWindow: () => getExistingWindow() as BrowserWindow | null,
        setWindow: (window) => setWindow(window as BrowserWindow | null),
        yomitanSession: (yomitanSession as Session | null | undefined) ?? appState.yomitanSession,
        onWindowClosed: () => {
          if (appState.yomitanParserWindow) {
            clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
          }
        },
      });
    },
    getExistingWindow: () => appState.yomitanSettingsWindow,
    setWindow: (window) => {
      appState.yomitanSettingsWindow = window as BrowserWindow | null;
    },
    logWarn: (message) => logger.warn(message),
    logError: (message, error) => logger.error(message, error),
  });

  function openYomitanSettings(): boolean {
    if (dictionaryBackend === 'hachidori') {
      if (deps.getConfig().yomitan.externalProfilePath.trim()) {
        logger.warn('Yomitan settings unavailable while using read-only external-profile mode.');
        return false;
      }
      inactiveYomitanLoad ??= inactiveYomitanExtension
        ? Promise.resolve(inactiveYomitanExtension)
        : loadYomitanExtensionCore({
            userDataPath: deps.userDataPath,
            getYomitanParserWindow: () => null,
            setYomitanParserWindow: () => {},
            setYomitanParserReadyPromise: () => {},
            setYomitanParserInitPromise: () => {},
            setYomitanExtension: (extension) => {
              inactiveYomitanExtension = extension;
            },
            setYomitanSession: () => {},
          }).finally(() => {
            inactiveYomitanLoad = null;
          });
      void (
        inactiveYomitanExtension ? Promise.resolve(inactiveYomitanExtension) : inactiveYomitanLoad
      )
        .then((extension) =>
          openYomitanSettingsWindow({
            yomitanExt: extension,
            getExistingWindow: () => inactiveYomitanSettingsWindow,
            setWindow: (window) => {
              inactiveYomitanSettingsWindow = window;
            },
          }),
        )
        .catch((error: unknown) => logger.error('Failed to open Yomitan settings', error));
      return true;
    }
    if (profilePolicy.isExternalReadOnlyMode()) {
      const message = 'Yomitan settings unavailable while using read-only external-profile mode.';
      logger.warn(
        'Yomitan settings window disabled while yomitan.externalProfilePath is configured because external profile mode is read-only.',
      );
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification(message, { variant: 'warning' });
      return false;
    }
    openYomitanSettingsHandler();
    return true;
  }

  function openHachidoriSettings(): void {
    void (dictionaryBackend === 'hachidori' ? ensureExtensionLoaded() : hachidori.ensureLoaded())
      .then((extension) =>
        openYomitanSettingsWindow({
          backend: 'hachidori',
          yomitanExt: extension,
          yomitanSession: hachidori.getSession(),
          getExistingWindow: () => hachidoriSettingsWindow,
          setWindow: (window) => {
            hachidoriSettingsWindow = window;
          },
          onWindowClosed: () => {
            if (dictionaryBackend === 'hachidori' && appState.yomitanParserWindow) {
              clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
            }
          },
        }),
      )
      .catch((error: unknown) => logger.error('Failed to open Hachidori settings', error));
  }

  function openDictionarySettings(): boolean {
    if (dictionaryBackend === 'hachidori') {
      openHachidoriSettings();
      return true;
    }
    return openYomitanSettings();
  }

  async function getCurrentAnkiDeckName(): Promise<string> {
    await extensionRuntime.ensureYomitanExtensionLoaded();
    return getYomitanCurrentAnkiDeckNameCore(getParserRuntimeDeps(), {
      error: (message, ...args) => {
        logger.error(message, ...args);
      },
      info: (message, ...args) => {
        logger.info(message, ...args);
      },
    });
  }

  return {
    profilePolicy,
    getParserRuntimeDeps,
    loadExtension,
    ensureExtensionLoaded,
    syncYomitanDefaultProfileAnkiServer,
    getCurrentAnkiDeckName,
    openYomitanSettings,
    openHachidoriSettings,
    openDictionarySettings,
    generateSentenceFurigana: (text: string, highlightedText?: string): Promise<string | null> =>
      generateSentenceFurigana(text, highlightedText, getParserRuntimeDeps(), logger),
    /** The session overlay windows load the active backend's content scripts from. */
    getOverlaySession: (): Session | null =>
      dictionaryBackend === 'hachidori' ? hachidori.getSession() : appState.yomitanSession,
    /** Dictionary settings windows the overlay must not cover. */
    getSettingsWindows: (): Array<BrowserWindow | null> => [
      appState.yomitanSettingsWindow,
      hachidoriSettingsWindow,
      inactiveYomitanSettingsWindow,
    ],
    destroySettingsWindows: (): void => {
      destroyYomitanSettingsWindow(hachidoriSettingsWindow);
      destroyYomitanSettingsWindow(inactiveYomitanSettingsWindow);
      destroyYomitanSettingsWindow(appState.yomitanSettingsWindow);
      appState.yomitanSettingsWindow = null;
    },
    /** Hachidori dictionary entries link out; only the active overlay's main frame may ask. */
    registerDictionaryExternalLinkHandler: (): void => {
      deps.ipcMain.handle(DICTIONARY_EXTERNAL_LINK_CHANNEL, async (event, value: unknown) => {
        if (
          dictionaryBackend !== 'hachidori' ||
          event.senderFrame !== event.sender.mainFrame ||
          !deps.getOverlayWindows().some((window) => window.webContents === event.sender)
        ) {
          throw new Error('Dictionary links are only available from the active overlay');
        }
        await deps.openExternal(parseDictionaryExternalUrl(value));
      });
    },
  };
}
