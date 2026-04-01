import type { BrowserWindow, Extension, Session } from 'electron';

import type { YomitanExtensionLoaderDeps } from '../core/services/yomitan-extension-loader';
import type { ResolvedConfig } from '../types';
import { createYomitanProfilePolicy } from './runtime/yomitan-profile-policy';
import { createYomitanRuntime, type YomitanRuntime } from './yomitan-runtime';

export interface YomitanRuntimeBootstrapInput {
  userDataPath: string;
  getResolvedConfig: () => ResolvedConfig;
  appState: {
    yomitanParserWindow: BrowserWindow | null;
    yomitanParserReadyPromise: Promise<void> | null;
    yomitanParserInitPromise: Promise<boolean> | null;
    yomitanExt: Extension | null;
    yomitanSession: Session | null;
    yomitanSettingsWindow: BrowserWindow | null;
  };
  loadYomitanExtensionCore: (options: YomitanExtensionLoaderDeps) => Promise<Extension | null>;
  getLoadInFlight: () => Promise<Extension | null> | null;
  setLoadInFlight: (promise: Promise<Extension | null> | null) => void;
  openYomitanSettingsWindow: (params: {
    yomitanExt: Extension | null;
    getExistingWindow: () => BrowserWindow | null;
    setWindow: (window: BrowserWindow | null) => void;
    yomitanSession?: Session | null;
    onWindowClosed?: () => void;
  }) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
}

export interface YomitanRuntimeBootstrap {
  yomitan: YomitanRuntime;
  yomitanProfilePolicy: ReturnType<typeof createYomitanProfilePolicy>;
}

export function createYomitanRuntimeBootstrap(
  input: YomitanRuntimeBootstrapInput,
): YomitanRuntimeBootstrap {
  const yomitanProfilePolicy = createYomitanProfilePolicy({
    externalProfilePath: input.getResolvedConfig().yomitan.externalProfilePath,
    logInfo: (message) => input.logInfo(message),
  });

  const yomitan = createYomitanRuntime({
    userDataPath: input.userDataPath,
    externalProfilePath: yomitanProfilePolicy.externalProfilePath,
    loadYomitanExtensionCore: input.loadYomitanExtensionCore,
    getYomitanParserWindow: () => input.appState.yomitanParserWindow,
    setYomitanParserWindow: (window) => {
      input.appState.yomitanParserWindow = window;
    },
    getYomitanParserReadyPromise: () => input.appState.yomitanParserReadyPromise,
    setYomitanParserReadyPromise: (promise) => {
      input.appState.yomitanParserReadyPromise = promise;
    },
    getYomitanParserInitPromise: () => input.appState.yomitanParserInitPromise,
    setYomitanParserInitPromise: (promise) => {
      input.appState.yomitanParserInitPromise = promise;
    },
    getYomitanExtension: () => input.appState.yomitanExt,
    setYomitanExtension: (extension) => {
      input.appState.yomitanExt = extension;
    },
    getYomitanSession: () => input.appState.yomitanSession,
    setYomitanSession: (session) => {
      input.appState.yomitanSession = session;
    },
    getLoadInFlight: () => input.getLoadInFlight(),
    setLoadInFlight: (promise) => {
      input.setLoadInFlight(promise);
    },
    getResolvedConfig: () => input.getResolvedConfig(),
    openYomitanSettingsWindow: (params) =>
      input.openYomitanSettingsWindow({
        yomitanExt: params.yomitanExt,
        getExistingWindow: params.getExistingWindow,
        setWindow: params.setWindow,
        yomitanSession: params.yomitanSession,
        onWindowClosed: params.onWindowClosed,
      }),
    getExistingSettingsWindow: () => input.appState.yomitanSettingsWindow,
    setSettingsWindow: (window) => {
      input.appState.yomitanSettingsWindow = window;
    },
    logInfo: (message) => input.logInfo(message),
    logWarn: (message) => input.logWarn(message),
    logError: (message, error) => input.logError(message, error),
    showMpvOsd: (message) => input.showMpvOsd(message),
    showDesktopNotification: (title, options) => input.showDesktopNotification(title, options),
  });

  return {
    yomitan,
    yomitanProfilePolicy,
  };
}
