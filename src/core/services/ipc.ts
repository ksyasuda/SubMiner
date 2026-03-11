import electron from 'electron';
import type { IpcMainEvent } from 'electron';
import type {
  ControllerPreferenceUpdate,
  ResolvedControllerConfig,
  RuntimeOptionId,
  RuntimeOptionValue,
  SubtitlePosition,
  SubsyncManualRunRequest,
  SubsyncResult,
} from '../../types';
import { IPC_CHANNELS, type OverlayHostedModal } from '../../shared/ipc/contracts';
import {
  parseMpvCommand,
  parseControllerPreferenceUpdate,
  parseOptionalForwardingOptions,
  parseOverlayHostedModal,
  parseRuntimeOptionDirection,
  parseRuntimeOptionId,
  parseRuntimeOptionValue,
  parseSubtitlePosition,
  parseSubsyncManualRunRequest,
} from '../../shared/ipc/validators';

const { BrowserWindow, ipcMain } = electron;

export interface IpcServiceDeps {
  onOverlayModalClosed: (modal: OverlayHostedModal) => void;
  onOverlayModalOpened?: (modal: OverlayHostedModal) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleDevTools: () => void;
  getVisibleOverlayVisibility: () => boolean;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getPlaybackPaused: () => boolean | null;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => {
    available: boolean;
    enabled: boolean;
    path: string | null;
  };
  setMecabEnabled: (enabled: boolean) => void;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getConfiguredShortcuts: () => unknown;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void;
  getSecondarySubMode: () => unknown;
  getCurrentSecondarySub: () => string;
  focusMainWindow: () => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
  appendClipboardVideoToQueue: () => { ok: boolean; message: string };
}

interface WindowLike {
  isDestroyed: () => boolean;
  focus: () => void;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  webContents: {
    toggleDevTools: () => void;
  };
}

interface MecabTokenizerLike {
  getStatus: () => {
    available: boolean;
    enabled: boolean;
    path: string | null;
  };
  setEnabled: (enabled: boolean) => void;
}

interface MpvClientLike {
  currentSecondarySubText?: string;
}

interface IpcMainRegistrar {
  on: (channel: string, listener: (event: unknown, ...args: unknown[]) => void) => void;
  handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => void;
}

export interface IpcDepsRuntimeOptions {
  getMainWindow: () => WindowLike | null;
  getVisibleOverlayVisibility: () => boolean;
  onOverlayModalClosed: (modal: OverlayHostedModal) => void;
  onOverlayModalOpened?: (modal: OverlayHostedModal) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getPlaybackPaused: () => boolean | null;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabTokenizer: () => MecabTokenizerLike | null;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getConfiguredShortcuts: () => unknown;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void;
  getSecondarySubMode: () => unknown;
  getMpvClient: () => MpvClientLike | null;
  focusMainWindow: () => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
  appendClipboardVideoToQueue: () => { ok: boolean; message: string };
}

export function createIpcDepsRuntime(options: IpcDepsRuntimeOptions): IpcServiceDeps {
  return {
    onOverlayModalClosed: options.onOverlayModalClosed,
    onOverlayModalOpened: options.onOverlayModalOpened,
    openYomitanSettings: options.openYomitanSettings,
    quitApp: options.quitApp,
    toggleDevTools: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.webContents.toggleDevTools();
    },
    getVisibleOverlayVisibility: options.getVisibleOverlayVisibility,
    toggleVisibleOverlay: options.toggleVisibleOverlay,
    tokenizeCurrentSubtitle: options.tokenizeCurrentSubtitle,
    getCurrentSubtitleRaw: options.getCurrentSubtitleRaw,
    getCurrentSubtitleAss: options.getCurrentSubtitleAss,
    getPlaybackPaused: options.getPlaybackPaused,
    getSubtitlePosition: options.getSubtitlePosition,
    getSubtitleStyle: options.getSubtitleStyle,
    saveSubtitlePosition: options.saveSubtitlePosition,
    getMecabStatus: () => {
      const mecabTokenizer = options.getMecabTokenizer();
      return mecabTokenizer
        ? mecabTokenizer.getStatus()
        : { available: false, enabled: false, path: null };
    },
    setMecabEnabled: (enabled) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) return;
      mecabTokenizer.setEnabled(enabled);
    },
    handleMpvCommand: options.handleMpvCommand,
    getKeybindings: options.getKeybindings,
    getConfiguredShortcuts: options.getConfiguredShortcuts,
    getControllerConfig: options.getControllerConfig,
    saveControllerPreference: options.saveControllerPreference,
    getSecondarySubMode: options.getSecondarySubMode,
    getCurrentSecondarySub: () => options.getMpvClient()?.currentSecondarySubText || '',
    focusMainWindow: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.focus();
    },
    runSubsyncManual: options.runSubsyncManual,
    getAnkiConnectStatus: options.getAnkiConnectStatus,
    getRuntimeOptions: options.getRuntimeOptions,
    setRuntimeOption: options.setRuntimeOption,
    cycleRuntimeOption: options.cycleRuntimeOption,
    reportOverlayContentBounds: options.reportOverlayContentBounds,
    getAnilistStatus: options.getAnilistStatus,
    clearAnilistToken: options.clearAnilistToken,
    openAnilistSetup: options.openAnilistSetup,
    getAnilistQueueStatus: options.getAnilistQueueStatus,
    retryAnilistQueueNow: options.retryAnilistQueueNow,
    appendClipboardVideoToQueue: options.appendClipboardVideoToQueue,
  };
}

export function registerIpcHandlers(deps: IpcServiceDeps, ipc: IpcMainRegistrar = ipcMain): void {
  ipc.on(
    IPC_CHANNELS.command.setIgnoreMouseEvents,
    (event: unknown, ignore: unknown, options: unknown = {}) => {
      if (typeof ignore !== 'boolean') return;
      const parsedOptions = parseOptionalForwardingOptions(options);
      const senderWindow = BrowserWindow.fromWebContents((event as IpcMainEvent).sender);
      if (senderWindow && !senderWindow.isDestroyed()) {
        senderWindow.setIgnoreMouseEvents(ignore, parsedOptions);
      }
    },
  );

  ipc.on(IPC_CHANNELS.command.overlayModalClosed, (_event: unknown, modal: unknown) => {
    const parsedModal = parseOverlayHostedModal(modal);
    if (!parsedModal) return;
    deps.onOverlayModalClosed(parsedModal);
  });
  ipc.on(IPC_CHANNELS.command.overlayModalOpened, (_event: unknown, modal: unknown) => {
    const parsedModal = parseOverlayHostedModal(modal);
    if (!parsedModal) return;
    if (!deps.onOverlayModalOpened) return;
    deps.onOverlayModalOpened(parsedModal);
  });

  ipc.on(IPC_CHANNELS.command.openYomitanSettings, () => {
    deps.openYomitanSettings();
  });

  ipc.on(IPC_CHANNELS.command.quitApp, () => {
    deps.quitApp();
  });

  ipc.on(IPC_CHANNELS.command.toggleDevTools, () => {
    deps.toggleDevTools();
  });

  ipc.on(IPC_CHANNELS.command.toggleOverlay, () => {
    deps.toggleVisibleOverlay();
  });

  ipc.handle(IPC_CHANNELS.request.getVisibleOverlayVisibility, () => {
    return deps.getVisibleOverlayVisibility();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitle, async () => {
    return await deps.tokenizeCurrentSubtitle();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitleRaw, () => {
    return deps.getCurrentSubtitleRaw();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitleAss, () => {
    return deps.getCurrentSubtitleAss();
  });

  ipc.handle(IPC_CHANNELS.request.getPlaybackPaused, () => {
    return deps.getPlaybackPaused();
  });

  ipc.handle(IPC_CHANNELS.request.getSubtitlePosition, () => {
    return deps.getSubtitlePosition();
  });

  ipc.handle(IPC_CHANNELS.request.getSubtitleStyle, () => {
    return deps.getSubtitleStyle();
  });

  ipc.on(IPC_CHANNELS.command.saveSubtitlePosition, (_event: unknown, position: unknown) => {
    const parsedPosition = parseSubtitlePosition(position);
    if (!parsedPosition) return;
    deps.saveSubtitlePosition(parsedPosition);
  });

  ipc.on(IPC_CHANNELS.command.saveControllerPreference, (_event: unknown, update: unknown) => {
    const parsedUpdate = parseControllerPreferenceUpdate(update);
    if (!parsedUpdate) return;
    deps.saveControllerPreference(parsedUpdate);
  });

  ipc.handle(IPC_CHANNELS.request.getMecabStatus, () => {
    return deps.getMecabStatus();
  });

  ipc.on(IPC_CHANNELS.command.setMecabEnabled, (_event: unknown, enabled: unknown) => {
    if (typeof enabled !== 'boolean') return;
    deps.setMecabEnabled(enabled);
  });

  ipc.on(IPC_CHANNELS.command.mpvCommand, (_event: unknown, command: unknown) => {
    const parsedCommand = parseMpvCommand(command);
    if (!parsedCommand) return;
    deps.handleMpvCommand(parsedCommand);
  });

  ipc.handle(IPC_CHANNELS.request.getKeybindings, () => {
    return deps.getKeybindings();
  });

  ipc.handle(IPC_CHANNELS.request.getConfigShortcuts, () => {
    return deps.getConfiguredShortcuts();
  });

  ipc.handle(IPC_CHANNELS.request.getControllerConfig, () => {
    return deps.getControllerConfig();
  });

  ipc.handle(IPC_CHANNELS.request.getSecondarySubMode, () => {
    return deps.getSecondarySubMode();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSecondarySub, () => {
    return deps.getCurrentSecondarySub();
  });

  ipc.handle(IPC_CHANNELS.request.focusMainWindow, () => {
    deps.focusMainWindow();
  });

  ipc.handle(IPC_CHANNELS.request.runSubsyncManual, async (_event, request: unknown) => {
    const parsedRequest = parseSubsyncManualRunRequest(request);
    if (!parsedRequest) {
      return { ok: false, message: 'Invalid subsync manual request payload' };
    }
    return await deps.runSubsyncManual(parsedRequest);
  });

  ipc.handle(IPC_CHANNELS.request.getAnkiConnectStatus, () => {
    return deps.getAnkiConnectStatus();
  });

  ipc.handle(IPC_CHANNELS.request.getRuntimeOptions, () => {
    return deps.getRuntimeOptions();
  });

  ipc.handle(IPC_CHANNELS.request.setRuntimeOption, (_event, id: unknown, value: unknown) => {
    const parsedId = parseRuntimeOptionId(id);
    if (!parsedId) {
      return { ok: false, error: 'Invalid runtime option id' };
    }
    const parsedValue = parseRuntimeOptionValue(value);
    if (parsedValue === null) {
      return { ok: false, error: 'Invalid runtime option value payload' };
    }
    return deps.setRuntimeOption(parsedId, parsedValue);
  });

  ipc.handle(IPC_CHANNELS.request.cycleRuntimeOption, (_event, id: unknown, direction: unknown) => {
    const parsedId = parseRuntimeOptionId(id);
    if (!parsedId) {
      return { ok: false, error: 'Invalid runtime option id' };
    }
    const parsedDirection = parseRuntimeOptionDirection(direction);
    if (!parsedDirection) {
      return { ok: false, error: 'Invalid runtime option cycle direction' };
    }
    return deps.cycleRuntimeOption(parsedId, parsedDirection);
  });

  ipc.on(IPC_CHANNELS.command.reportOverlayContentBounds, (_event: unknown, payload: unknown) => {
    deps.reportOverlayContentBounds(payload);
  });

  ipc.handle(IPC_CHANNELS.request.getAnilistStatus, () => {
    return deps.getAnilistStatus();
  });

  ipc.handle(IPC_CHANNELS.request.clearAnilistToken, () => {
    deps.clearAnilistToken();
    return { ok: true };
  });

  ipc.handle(IPC_CHANNELS.request.openAnilistSetup, () => {
    deps.openAnilistSetup();
    return { ok: true };
  });

  ipc.handle(IPC_CHANNELS.request.getAnilistQueueStatus, () => {
    return deps.getAnilistQueueStatus();
  });

  ipc.handle(IPC_CHANNELS.request.retryAnilistNow, async () => {
    return await deps.retryAnilistQueueNow();
  });

  ipc.handle(IPC_CHANNELS.request.appendClipboardVideoToQueue, () => {
    return deps.appendClipboardVideoToQueue();
  });
}
