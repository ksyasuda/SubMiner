import path from 'node:path';
import fs from 'node:fs';
import type { Menu, MenuItemConstructorOptions, NativeImage, Tray } from 'electron';
import type { Logger } from '../../logger';
import {
  buildTrayMenuTemplateRuntime,
  resolveTrayIconPathRuntime,
} from '../runtime/domains/overlay';
import type { createBuildTrayMenuTemplateHandler } from '../runtime/tray-main-actions';
import { createTrayRuntimeHandlers } from '../runtime/tray-runtime-handlers';
import type { AppState } from '../state';

type TrayMenuTemplateDeps = Parameters<
  typeof createBuildTrayMenuTemplateHandler<MenuItemConstructorOptions>
>[0];

/** The tray menu's entries; each one calls into the domain that owns it. */
export type TrayMenuActions = Omit<
  TrayMenuTemplateDeps,
  | 'buildTrayMenuTemplateRuntime'
  | 'platform'
  | 'initializeOverlayRuntime'
  | 'isOverlayRuntimeInitialized'
  | 'showWindowsMpvLauncherSetup'
>;

export interface TrayAssemblyDeps {
  appState: Pick<AppState, 'overlayRuntimeInitialized'>;
  logger: Pick<Logger, 'warn'>;
  /** Where the packaged and dev builds keep the tray icon. */
  iconRoots: { resourcesPath: string; appPath: string; mainDirname: string };
  /** Electron tray primitives, injected so the assembly loads under `bun test`. */
  electron: {
    createTray: (icon: NativeImage) => Tray;
    createImageFromPath: (iconPath: string) => NativeImage;
    createEmptyImage: () => NativeImage;
    buildMenuFromTemplate: (template: MenuItemConstructorOptions[]) => Menu;
  };
  menu: TrayMenuActions;
  initializeOverlayRuntime: () => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
}

export type TrayAssembly = ReturnType<typeof createTrayAssembly>;

/**
 * Owns the system tray icon and its menu. The tray exists in background mode and for managed
 * playback; clicking it shows the overlay.
 */
export function createTrayAssembly(deps: TrayAssemblyDeps) {
  const { appState } = deps;
  let tray: Tray | null = null;

  const { ensureTray, destroyTray } = createTrayRuntimeHandlers({
    resolveTrayIconPathDeps: {
      resolveTrayIconPathRuntime,
      platform: process.platform,
      resourcesPath: deps.iconRoots.resourcesPath,
      appPath: deps.iconRoots.appPath,
      dirname: deps.iconRoots.mainDirname,
      joinPath: (...parts) => path.join(...parts),
      fileExists: (candidate) => fs.existsSync(candidate),
    },
    buildTrayMenuTemplateDeps: {
      ...deps.menu,
      buildTrayMenuTemplateRuntime,
      platform: process.platform,
      initializeOverlayRuntime: () => deps.initializeOverlayRuntime(),
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      showWindowsMpvLauncherSetup: () => process.platform === 'win32',
    },
    ensureTrayDeps: {
      getTray: () => tray,
      setTray: (nextTray) => {
        tray = nextTray as Tray | null;
      },
      createImageFromPath: (iconPath) => deps.electron.createImageFromPath(iconPath),
      createEmptyImage: () => deps.electron.createEmptyImage(),
      createTray: (icon) => deps.electron.createTray(icon as NativeImage),
      trayTooltip: 'SubMiner',
      platform: process.platform,
      logWarn: (message) => deps.logger.warn(message),
      initializeOverlayRuntime: () => deps.initializeOverlayRuntime(),
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      setVisibleOverlayVisible: (visible) => deps.setVisibleOverlayVisible(visible),
    },
    destroyTrayDeps: {
      getTray: () => tray,
      setTray: (nextTray) => {
        tray = nextTray as Tray | null;
      },
    },
    buildMenuFromTemplate: (template) => deps.electron.buildMenuFromTemplate(template),
  });

  return {
    ensureTray,
    destroyTray,
    hasTray: (): boolean => Boolean(tray),
    /** Rebuilds the menu so state-dependent entries (setup, Jellyfin) stay current. */
    refreshMenu: (): void => {
      if (tray) {
        ensureTray();
      }
    },
  };
}
