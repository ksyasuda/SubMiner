import * as path from 'node:path';
import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import { releaseDockIcon, retainDockIcon } from '../../core/services/dock-icon-visibility';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { AnimeBrowserPlaybackState } from '../../types/anime-browser';
import {
  registerAnimeBrowserIpcHandlers,
  type AnimeBrowserIpcSender,
} from './anime-browser-ipc-handlers';
import { createAnimeBrowserRuntime, type AnimeBrowserRuntimeDeps } from './anime-browser-runtime';
import { createAnimeBrowserSessionRegistry } from './anime-browser-sessions';
import { createOpenConfigSettingsWindowHandler } from './config-settings-window';
import { createCreateAnimeBrowserWindowHandler } from './setup-window-factory';
import {
  toAnimeBrowserPlaybackState,
  type StreamPlaybackMetadataStore,
} from './stream-playback-metadata';

type AnimeBrowserWindow = Pick<
  BrowserWindow,
  'destroy' | 'focus' | 'isDestroyed' | 'loadFile' | 'on' | 'show' | 'webContents'
>;

interface DockLike {
  show: () => Promise<void> | void;
  hide: () => void;
}

export interface AnimeBrowserApplicationRuntimeDeps {
  userDataPath: string;
  mainModuleDir: string;
  runtime: Omit<
    AnimeBrowserRuntimeDeps,
    | 'extensionsDir'
    | 'onBridgeState'
    | 'onPlaybackMetadata'
    | 'onPreparedPlaybackMetadata'
    | 'onQueueState'
    | 'onSearchUpdate'
  >;
  configuredExtensionsDir: () => string | undefined;
  ipcMain: {
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
  };
  playbackMetadata: StreamPlaybackMetadataStore;
  getInitialPlaybackMetadata: () => AnimeStreamMetadata | null;
  handlePlaybackMetadata: (metadata: AnimeStreamMetadata, prepared: boolean) => void;
  getAnimeBrowserWindow: () => AnimeBrowserWindow | null;
  setAnimeBrowserWindow: (window: AnimeBrowserWindow | null) => void;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => AnimeBrowserWindow;
  promoteWindowAboveOverlay: (window: AnimeBrowserWindow) => void;
  activateApp: () => void;
  dock: DockLike | null | undefined;
  shouldRehideDockIcon: () => boolean;
  isStandaloneAnimeBrowserLaunch: () => boolean;
  isMpvConnected: () => boolean;
  ensureTray: () => void;
  requestAppQuit: () => void;
  logError: (message: string) => void;
}

export interface AnimeBrowserApplicationRuntime {
  publishPlaybackState: (mediaPath: string | null) => void;
  openWindow: () => boolean;
}

/** Compose the anime browser runtime, IPC sessions, event fanout, and standalone window. */
export function createAnimeBrowserApplicationRuntime(
  deps: AnimeBrowserApplicationRuntimeDeps,
): AnimeBrowserApplicationRuntime {
  let playbackState: AnimeBrowserPlaybackState | null = toAnimeBrowserPlaybackState(
    deps.getInitialPlaybackMetadata(),
  );

  const broadcast = (channel: string, payload: unknown): void => {
    const targets = new Set<AnimeBrowserIpcSender>();
    const standalone = deps.getAnimeBrowserWindow();
    if (standalone && !standalone.isDestroyed()) targets.add(standalone.webContents);
    for (const sender of sessions.values()) {
      if (!sender.isDestroyed()) targets.add(sender);
    }
    for (const sender of targets) sender.send(channel, payload);
  };

  const publishPlaybackState = (mediaPath: string | null): void => {
    playbackState = toAnimeBrowserPlaybackState(deps.playbackMetadata.match(mediaPath));
    broadcast(IPC_CHANNELS.event.animeBrowserPlaybackState, playbackState);
  };

  const runtime = createAnimeBrowserRuntime({
    ...deps.runtime,
    extensionsDir: () => {
      const configured = deps.configuredExtensionsDir()?.trim();
      return configured && configured.length > 0
        ? configured
        : path.join(deps.userDataPath, 'anime-extensions');
    },
    onPlaybackMetadata: (metadata) => {
      deps.playbackMetadata.set(metadata);
      publishPlaybackState(metadata.mediaPath);
      deps.handlePlaybackMetadata(metadata, false);
    },
    onPreparedPlaybackMetadata: (metadata) => {
      deps.playbackMetadata.set(metadata);
      deps.handlePlaybackMetadata(metadata, true);
    },
    onBridgeState: (state) => broadcast(IPC_CHANNELS.event.animeBrowserBridgeState, state),
    onSearchUpdate: (update, sessionId) => {
      const sender = sessions.get(sessionId);
      if (sender && !sender.isDestroyed()) {
        sender.send(IPC_CHANNELS.event.animeBrowserSearchUpdate, update);
      }
    },
    onQueueState: (state) => broadcast(IPC_CHANNELS.event.animeBrowserQueueState, state),
  });
  const sessions = createAnimeBrowserSessionRegistry((sessionId) =>
    runtime.releaseSession(sessionId),
  );

  registerAnimeBrowserIpcHandlers({
    ipcMain: deps.ipcMain,
    runtime,
    getPlaybackState: () => playbackState,
    registerSession: sessions.register,
  });

  let dockIconRetained = false;
  const releaseBrowserDockIcon = (): void => {
    if (!dockIconRetained) return;
    dockIconRetained = false;
    releaseDockIcon({ dock: deps.dock, shouldRehide: deps.shouldRehideDockIcon });
  };

  const openWindowBase = createOpenConfigSettingsWindowHandler({
    getSettingsWindow: deps.getAnimeBrowserWindow,
    setSettingsWindow: deps.setAnimeBrowserWindow,
    createSettingsWindow: createCreateAnimeBrowserWindowHandler({
      createBrowserWindow: deps.createBrowserWindow,
      preloadPath: path.join(deps.mainModuleDir, 'preload-animeui.js'),
    }),
    settingsHtmlPath: path.join(deps.mainModuleDir, 'animeui', 'index.html'),
    promoteSettingsWindowAboveOverlay: deps.promoteWindowAboveOverlay,
    activateApp: deps.activateApp,
    onClosed: () => {
      releaseBrowserDockIcon();
      if (deps.isStandaloneAnimeBrowserLaunch() && !deps.isMpvConnected()) {
        void runtime.dispose().finally(deps.requestAppQuit);
      }
    },
    log: deps.logError,
  });

  const openWindow = (): boolean => {
    if (!dockIconRetained) {
      dockIconRetained = true;
      retainDockIcon({ dock: deps.dock });
    }
    const opened = openWindowBase();
    if (!opened) {
      releaseBrowserDockIcon();
      return false;
    }
    deps.ensureTray();
    return true;
  };

  return { publishPlaybackState, openWindow };
}
