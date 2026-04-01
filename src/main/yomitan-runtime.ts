import type { BrowserWindow, Extension, Session } from 'electron';

import { clearYomitanParserCachesForWindow, syncYomitanDefaultAnkiServer } from '../core/services';
import type { YomitanExtensionLoaderDeps } from '../core/services/yomitan-extension-loader';
import type { ResolvedConfig } from '../types';
import { createYomitanExtensionRuntime } from './runtime/yomitan-extension-runtime';
import { createYomitanProfilePolicy } from './runtime/yomitan-profile-policy';
import { createYomitanSettingsRuntime } from './runtime/yomitan-settings-runtime';
import {
  getPreferredYomitanAnkiServerUrl,
  shouldForceOverrideYomitanAnkiServer,
} from './runtime/yomitan-anki-server';
export interface YomitanParserRuntimeDeps {
  getYomitanExt: () => Extension | null;
  getYomitanSession: () => Session | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
}

export interface YomitanRuntimeInput {
  userDataPath: string;
  externalProfilePath: string;
  loadYomitanExtensionCore: (options: YomitanExtensionLoaderDeps) => Promise<Extension | null>;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  getYomitanExtension: () => Extension | null;
  setYomitanExtension: (extension: Extension | null) => void;
  getYomitanSession: () => Session | null;
  setYomitanSession: (session: Session | null) => void;
  getLoadInFlight: () => Promise<Extension | null> | null;
  setLoadInFlight: (promise: Promise<Extension | null> | null) => void;
  getResolvedConfig: () => ResolvedConfig;
  openYomitanSettingsWindow: (params: {
    yomitanExt: Extension | null;
    getExistingWindow: () => BrowserWindow | null;
    setWindow: (window: BrowserWindow | null) => void;
    yomitanSession?: Session | null;
    onWindowClosed?: () => void;
  }) => void;
  getExistingSettingsWindow: () => BrowserWindow | null;
  setSettingsWindow: (window: BrowserWindow | null) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
}

export interface YomitanRuntime {
  loadYomitanExtension: () => Promise<Extension | null>;
  ensureYomitanExtensionLoaded: () => Promise<Extension | null>;
  openYomitanSettings: () => boolean;
  syncDefaultProfileAnkiServer: () => Promise<void>;
  getParserRuntimeDeps: () => YomitanParserRuntimeDeps;
  getPreferredAnkiServerUrl: () => string;
  isExternalReadOnlyMode: () => boolean;
  isCharacterDictionaryEnabled: () => boolean;
  getCharacterDictionaryDisabledReason: () => string | null;
  clearParserCachesForWindow: (window: BrowserWindow) => void;
}

export function createYomitanRuntime(input: YomitanRuntimeInput): YomitanRuntime {
  const profilePolicy = createYomitanProfilePolicy({
    externalProfilePath: input.externalProfilePath,
    logInfo: (message) => input.logInfo(message),
  });

  const extensionRuntime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: input.loadYomitanExtensionCore,
    userDataPath: input.userDataPath,
    externalProfilePath: profilePolicy.externalProfilePath,
    getYomitanParserWindow: () => input.getYomitanParserWindow(),
    setYomitanParserWindow: (window) => input.setYomitanParserWindow(window),
    setYomitanParserReadyPromise: (promise) => input.setYomitanParserReadyPromise(promise),
    setYomitanParserInitPromise: (promise) => input.setYomitanParserInitPromise(promise),
    setYomitanExtension: (extension) => input.setYomitanExtension(extension),
    setYomitanSession: (session) => input.setYomitanSession(session),
    getYomitanExtension: () => input.getYomitanExtension(),
    getLoadInFlight: () => input.getLoadInFlight(),
    setLoadInFlight: (promise) => input.setLoadInFlight(promise),
  });

  const settingsRuntime = createYomitanSettingsRuntime({
    ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
    openYomitanSettingsWindow: (params) =>
      input.openYomitanSettingsWindow({
        yomitanExt: params.yomitanExt as Extension | null,
        getExistingWindow: () => params.getExistingWindow() as BrowserWindow | null,
        setWindow: (window) => params.setWindow(window),
        yomitanSession: (params.yomitanSession as Session | null | undefined) ?? null,
        onWindowClosed: () => {
          input.setSettingsWindow(null);
          params.onWindowClosed?.();
        },
      }),
    getExistingWindow: () => input.getExistingSettingsWindow(),
    setWindow: (window) => input.setSettingsWindow(window as BrowserWindow | null),
    getYomitanSession: () => input.getYomitanSession(),
    logWarn: (message) => input.logWarn(message),
    logError: (message, error) => input.logError(message, error),
  });

  let lastSyncedYomitanAnkiServer: string | null = null;

  const getParserRuntimeDeps = (): YomitanParserRuntimeDeps => ({
    getYomitanExt: () => input.getYomitanExtension(),
    getYomitanSession: () => input.getYomitanSession(),
    getYomitanParserWindow: () => input.getYomitanParserWindow(),
    setYomitanParserWindow: (window) => input.setYomitanParserWindow(window),
    getYomitanParserReadyPromise: () => input.getYomitanParserReadyPromise(),
    setYomitanParserReadyPromise: (promise) => input.setYomitanParserReadyPromise(promise),
    getYomitanParserInitPromise: () => input.getYomitanParserInitPromise(),
    setYomitanParserInitPromise: (promise) => input.setYomitanParserInitPromise(promise),
  });

  const syncDefaultProfileAnkiServer = async (): Promise<void> => {
    if (profilePolicy.isExternalReadOnlyMode()) {
      return;
    }

    const targetUrl = getPreferredYomitanAnkiServerUrl(
      input.getResolvedConfig().ankiConnect,
    ).trim();
    if (!targetUrl || targetUrl === lastSyncedYomitanAnkiServer) {
      return;
    }

    const synced = await syncYomitanDefaultAnkiServer(
      targetUrl,
      getParserRuntimeDeps(),
      {
        error: (message, ...args) => {
          input.logError(message, args[0]);
        },
        info: (message, ...args) => {
          input.logInfo([message, ...args].join(' '));
        },
      },
      {
        forceOverride: shouldForceOverrideYomitanAnkiServer(input.getResolvedConfig().ankiConnect),
      },
    );

    if (synced) {
      lastSyncedYomitanAnkiServer = targetUrl;
    }
  };

  const loadYomitanExtension = async (): Promise<Extension | null> => {
    const extension = await extensionRuntime.loadYomitanExtension();
    if (extension && !profilePolicy.isExternalReadOnlyMode()) {
      await syncDefaultProfileAnkiServer();
    }
    return extension;
  };

  const ensureYomitanExtensionLoaded = async (): Promise<Extension | null> => {
    const extension = await extensionRuntime.ensureYomitanExtensionLoaded();
    if (extension && !profilePolicy.isExternalReadOnlyMode()) {
      await syncDefaultProfileAnkiServer();
    }
    return extension;
  };

  const openYomitanSettings = (): boolean => {
    if (profilePolicy.isExternalReadOnlyMode()) {
      const message = 'Yomitan settings unavailable while using read-only external-profile mode.';
      input.logWarn(
        'Yomitan settings window disabled while yomitan.externalProfilePath is configured because external profile mode is read-only.',
      );
      input.showDesktopNotification('SubMiner', { body: message });
      input.showMpvOsd(message);
      return false;
    }

    settingsRuntime.openYomitanSettings();
    return true;
  };

  return {
    loadYomitanExtension,
    ensureYomitanExtensionLoaded,
    openYomitanSettings,
    syncDefaultProfileAnkiServer,
    getParserRuntimeDeps,
    getPreferredAnkiServerUrl: () =>
      getPreferredYomitanAnkiServerUrl(input.getResolvedConfig().ankiConnect),
    isExternalReadOnlyMode: () => profilePolicy.isExternalReadOnlyMode(),
    isCharacterDictionaryEnabled: () => profilePolicy.isCharacterDictionaryEnabled(),
    getCharacterDictionaryDisabledReason: () =>
      profilePolicy.getCharacterDictionaryDisabledReason(),
    clearParserCachesForWindow: (window) => clearYomitanParserCachesForWindow(window),
  };
}
