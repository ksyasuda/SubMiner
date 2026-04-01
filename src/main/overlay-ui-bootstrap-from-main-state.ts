import fs from 'node:fs';
import path from 'node:path';

import {
  Menu,
  MenuItem,
  nativeImage,
  Tray,
  type BrowserWindow,
  type MenuItemConstructorOptions,
} from 'electron';

import type { AnilistRuntime } from './anilist-runtime';
import type { DictionarySupportRuntime } from './dictionary-support-runtime';
import type { FirstRunRuntime } from './first-run-runtime';
import type { JellyfinRuntime } from './jellyfin-runtime';
import type { MpvRuntime } from './mpv-runtime';
import type { ResolvedConfig } from '../types';
import {
  broadcastRuntimeOptionsChangedRuntime,
  createOverlayWindow as createOverlayWindowCore,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  setOverlayDebugVisualizationEnabledRuntime,
  syncOverlayWindowLayer,
} from '../core/services';
import {
  buildTrayMenuTemplateRuntime,
  resolveTrayIconPathRuntime,
} from './runtime/domains/overlay';
import {
  createOverlayUiBootstrapRuntime,
  type OverlayUiBootstrapInput,
  type OverlayUiBootstrapRuntime,
} from './overlay-ui-bootstrap-runtime';
import type { OverlayModalRuntime } from './overlay-runtime';
import type { ShortcutsRuntimeBootstrap } from './shortcuts-runtime';
import { createWindowTracker as createWindowTrackerCore } from '../window-trackers';

export interface OverlayUiBootstrapFromMainStateInput<
  TWindow extends BrowserWindow,
  TMenuItem = MenuItemConstructorOptions | MenuItem,
> {
  appState: OverlayUiBootstrapInput<TWindow>['appState'];
  overlayManager: OverlayUiBootstrapInput<TWindow>['overlayManager'];
  overlayModalInputState: OverlayUiBootstrapInput<TWindow>['overlayModalInputState'];
  overlayModalRuntime: OverlayModalRuntime;
  overlayShortcutsRuntime: ShortcutsRuntimeBootstrap['overlayShortcutsRuntime'];
  runtimes: {
    dictionarySupport: Pick<DictionarySupportRuntime, 'createFieldGroupingCallback'>;
    firstRun: Pick<FirstRunRuntime, 'isSetupCompleted' | 'openFirstRunSetupWindow'>;
    yomitan: {
      openYomitanSettings: () => boolean;
    };
    jellyfin: Pick<JellyfinRuntime, 'openJellyfinSetupWindow'>;
    anilist: Pick<AnilistRuntime, 'openAnilistSetupWindow'>;
    shortcuts: Pick<ShortcutsRuntimeBootstrap['shortcuts'], 'registerGlobalShortcuts'>;
    mpvRuntime: Pick<MpvRuntime, 'startBackgroundWarmups'>;
  };
  electron: OverlayUiBootstrapInput<TWindow>['electron'] & {
    buildMenuFromTemplate: (template: TMenuItem[]) => unknown;
    createTray: (
      icon: ReturnType<OverlayUiBootstrapInput<TWindow>['electron']['createEmptyImage']>,
    ) => Tray;
  };
  windowing: OverlayUiBootstrapInput<TWindow>['windowing'];
  actions: Omit<
    OverlayUiBootstrapInput<TWindow>['actions'],
    'registerGlobalShortcuts' | 'startBackgroundWarmups'
  >;
  trayState: OverlayUiBootstrapInput<TWindow>['trayState'];
  startup: OverlayUiBootstrapInput<TWindow>['startup'];
}

export function createOverlayUiBootstrapFromMainState<TWindow extends BrowserWindow>(
  input: OverlayUiBootstrapFromMainStateInput<TWindow>,
): OverlayUiBootstrapRuntime<TWindow> {
  return createOverlayUiBootstrapRuntime<TWindow>({
    appState: input.appState,
    overlayManager: input.overlayManager,
    overlayModalInputState: input.overlayModalInputState,
    overlayModalRuntime: input.overlayModalRuntime,
    overlayShortcutsRuntime: input.overlayShortcutsRuntime,
    dictionarySupport: {
      createFieldGroupingCallback: () =>
        input.runtimes.dictionarySupport.createFieldGroupingCallback(),
    },
    firstRun: {
      isSetupCompleted: () => input.runtimes.firstRun.isSetupCompleted(),
      openFirstRunSetupWindow: () => input.runtimes.firstRun.openFirstRunSetupWindow(),
    },
    yomitan: {
      openYomitanSettings: () => {
        input.runtimes.yomitan.openYomitanSettings();
      },
    },
    jellyfin: {
      openJellyfinSetupWindow: () => input.runtimes.jellyfin.openJellyfinSetupWindow(),
    },
    anilist: {
      openAnilistSetupWindow: () => input.runtimes.anilist.openAnilistSetupWindow(),
    },
    electron: input.electron,
    windowing: input.windowing,
    actions: {
      ...input.actions,
      registerGlobalShortcuts: () => input.runtimes.shortcuts.registerGlobalShortcuts(),
      startBackgroundWarmups: () => input.runtimes.mpvRuntime.startBackgroundWarmups(),
    },
    trayState: input.trayState,
    startup: input.startup,
  });
}

export interface OverlayUiBootstrapCoordinatorInput<TWindow extends BrowserWindow> {
  appState: OverlayUiBootstrapFromMainStateInput<TWindow>['appState'];
  overlayManager: OverlayUiBootstrapFromMainStateInput<TWindow>['overlayManager'];
  overlayModalInputState: OverlayUiBootstrapFromMainStateInput<TWindow>['overlayModalInputState'];
  overlayModalRuntime: OverlayUiBootstrapFromMainStateInput<TWindow>['overlayModalRuntime'];
  overlayShortcutsRuntime: OverlayUiBootstrapFromMainStateInput<TWindow>['overlayShortcutsRuntime'];
  runtimes: OverlayUiBootstrapFromMainStateInput<TWindow>['runtimes'];
  env: {
    screen: OverlayUiBootstrapFromMainStateInput<TWindow>['electron']['screen'];
    appPath: string;
    resourcesPath: string;
    dirname: string;
    platform: NodeJS.Platform;
  };
  windowing: OverlayUiBootstrapFromMainStateInput<TWindow>['windowing'];
  actions: Omit<
    OverlayUiBootstrapFromMainStateInput<TWindow>['actions'],
    | 'resolveTrayIconPathRuntime'
    | 'buildTrayMenuTemplateRuntime'
    | 'broadcastRuntimeOptionsChangedRuntime'
    | 'setOverlayDebugVisualizationEnabledRuntime'
    | 'initializeOverlayRuntimeCore'
  > &
    Pick<
      OverlayUiBootstrapFromMainStateInput<TWindow>['actions'],
      | 'resolveTrayIconPathRuntime'
      | 'buildTrayMenuTemplateRuntime'
      | 'broadcastRuntimeOptionsChangedRuntime'
      | 'setOverlayDebugVisualizationEnabledRuntime'
      | 'initializeOverlayRuntimeCore'
    >;
  trayState: OverlayUiBootstrapFromMainStateInput<TWindow>['trayState'];
  startup: OverlayUiBootstrapFromMainStateInput<TWindow>['startup'];
}

export function createOverlayUiBootstrapCoordinator<TWindow extends BrowserWindow>(
  input: OverlayUiBootstrapCoordinatorInput<TWindow>,
): OverlayUiBootstrapRuntime<TWindow> {
  return createOverlayUiBootstrapFromMainState<TWindow>({
    appState: input.appState,
    overlayManager: input.overlayManager,
    overlayModalInputState: input.overlayModalInputState,
    overlayModalRuntime: input.overlayModalRuntime,
    overlayShortcutsRuntime: input.overlayShortcutsRuntime,
    runtimes: input.runtimes,
    electron: {
      screen: input.env.screen,
      appPath: input.env.appPath,
      resourcesPath: input.env.resourcesPath,
      dirname: input.env.dirname,
      platform: input.env.platform,
      joinPath: (...parts) => path.join(...parts),
      fileExists: (candidate) => fs.existsSync(candidate),
      createImageFromPath: (iconPath) => nativeImage.createFromPath(iconPath),
      createEmptyImage: () => nativeImage.createEmpty(),
      createTray: (icon) => new Tray(icon as ConstructorParameters<typeof Tray>[0]),
      buildMenuFromTemplate: (template) =>
        Menu.buildFromTemplate(template as (MenuItemConstructorOptions | MenuItem)[]),
    },
    windowing: input.windowing,
    actions: input.actions,
    trayState: input.trayState,
    startup: input.startup,
  });
}

export interface OverlayUiBootstrapFromProcessStateInput<TWindow extends BrowserWindow> {
  appState: OverlayUiBootstrapCoordinatorInput<TWindow>['appState'];
  overlayManager: OverlayUiBootstrapCoordinatorInput<TWindow>['overlayManager'];
  overlayModalInputState: OverlayUiBootstrapCoordinatorInput<TWindow>['overlayModalInputState'];
  overlayModalRuntime: OverlayUiBootstrapCoordinatorInput<TWindow>['overlayModalRuntime'];
  overlayShortcutsRuntime: OverlayUiBootstrapCoordinatorInput<TWindow>['overlayShortcutsRuntime'];
  runtimes: OverlayUiBootstrapCoordinatorInput<TWindow>['runtimes'];
  env: OverlayUiBootstrapCoordinatorInput<TWindow>['env'] & {
    isDev: boolean;
  };
  actions: {
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    sendMpvCommand: (command: (string | number)[]) => void;
    ensureOverlayMpvSubtitlesHidden: () => Promise<void>;
    syncOverlayMpvSubtitleSuppression: () => void;
    getResolvedConfig: () => ResolvedConfig;
    requestAppQuit: () => void;
  };
  trayState: OverlayUiBootstrapCoordinatorInput<TWindow>['trayState'];
  startup: OverlayUiBootstrapCoordinatorInput<TWindow>['startup'];
}

export function createOverlayUiBootstrapFromProcessState<TWindow extends BrowserWindow>(
  input: OverlayUiBootstrapFromProcessStateInput<TWindow>,
): OverlayUiBootstrapRuntime<TWindow> {
  return createOverlayUiBootstrapCoordinator({
    appState: input.appState,
    overlayManager: input.overlayManager,
    overlayModalInputState: input.overlayModalInputState,
    overlayModalRuntime: input.overlayModalRuntime,
    overlayShortcutsRuntime: input.overlayShortcutsRuntime,
    runtimes: input.runtimes,
    env: input.env,
    windowing: {
      isDev: input.env.isDev,
      createOverlayWindowCore: (kind, options) =>
        createOverlayWindowCore(kind, options as never) as TWindow,
      ensureOverlayWindowLevelCore: (window) =>
        ensureOverlayWindowLevelCore(window as BrowserWindow),
      syncOverlayWindowLayer: (window, layer) =>
        syncOverlayWindowLayer(window as BrowserWindow, layer),
      enforceOverlayLayerOrderCore: (params) =>
        enforceOverlayLayerOrderCore({
          ...params,
          mainWindow: params.mainWindow as BrowserWindow | null,
          ensureOverlayWindowLevel: (window) => params.ensureOverlayWindowLevel(window as TWindow),
        }),
      createWindowTrackerCore: (override, targetMpvSocketPath) =>
        createWindowTrackerCore(override, targetMpvSocketPath),
    },
    actions: {
      showMpvOsd: (message) => input.actions.showMpvOsd(message),
      showDesktopNotification: (title, options) =>
        input.actions.showDesktopNotification(title, options),
      sendMpvCommand: (command) => input.actions.sendMpvCommand(command),
      broadcastRuntimeOptionsChangedRuntime,
      setOverlayDebugVisualizationEnabledRuntime,
      resolveTrayIconPathRuntime,
      buildTrayMenuTemplateRuntime,
      initializeOverlayRuntimeCore: (options) => initializeOverlayRuntimeCore(options as never),
      ensureOverlayMpvSubtitlesHidden: () => input.actions.ensureOverlayMpvSubtitlesHidden(),
      syncOverlayMpvSubtitleSuppression: () => input.actions.syncOverlayMpvSubtitleSuppression(),
      getResolvedConfig: () => input.actions.getResolvedConfig(),
      requestAppQuit: input.actions.requestAppQuit,
    },
    trayState: input.trayState,
    startup: input.startup,
  });
}
