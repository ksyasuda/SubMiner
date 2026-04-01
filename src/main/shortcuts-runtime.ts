import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { ResolvedConfig } from '../types';
import type { ConfiguredShortcuts } from '../core/utils/shortcut-config';
import { DEFAULT_CONFIG } from '../config';
import { resolveConfiguredShortcuts } from '../core/utils';
import type { AppState } from './state';
import type { MiningRuntime } from './mining-runtime';
import type { OverlayModalRuntime } from './overlay-runtime';
import { createBuildOverlayShortcutsRuntimeMainDepsHandler } from './runtime/domains/shortcuts';
import {
  composeShortcutRuntimes,
  type ShortcutsRuntimeComposerOptions,
} from './runtime/composers/shortcuts-runtime-composer';
import { createOverlayShortcutsRuntimeService } from './overlay-shortcuts-runtime';

type GlobalShortcutsInput = ShortcutsRuntimeComposerOptions['globalShortcuts'];
type NumericShortcutRuntimeMainDepsInput =
  ShortcutsRuntimeComposerOptions['numericShortcutRuntimeMainDeps'];
type NumericSessionsInput = ShortcutsRuntimeComposerOptions['numericSessions'];
type OverlayShortcutsRuntimeMainDepsInput =
  ShortcutsRuntimeComposerOptions['overlayShortcutsRuntimeMainDeps'];

export interface ShortcutsRuntimeInput {
  globalShortcuts: GlobalShortcutsInput;
  numericShortcutRuntimeMainDeps: NumericShortcutRuntimeMainDepsInput;
  numericSessions: NumericSessionsInput;
  overlayShortcutsRuntimeMainDeps: OverlayShortcutsRuntimeMainDepsInput;
}

export interface ShortcutsRuntime {
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  registerGlobalShortcuts: () => void;
  refreshGlobalAndOverlayShortcuts: () => void;
  cancelPendingMultiCopy: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  cancelPendingMineSentenceMultiple: () => void;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  registerOverlayShortcuts: () => void;
  unregisterOverlayShortcuts: () => void;
  syncOverlayShortcuts: () => void;
  refreshOverlayShortcuts: () => void;
  syncOverlayShortcutsForModal: (isActive: boolean) => void;
}

export interface ShortcutsRuntimeBootstrapInput {
  globalShortcuts: ShortcutsRuntimeInput['globalShortcuts'];
  numericShortcutRuntimeMainDeps: ShortcutsRuntimeInput['numericShortcutRuntimeMainDeps'];
  numericSessions: ShortcutsRuntimeInput['numericSessions'];
  overlayShortcuts: {
    getConfiguredShortcuts: () => ConfiguredShortcuts;
    getShortcutsRegistered: () => boolean;
    setShortcutsRegistered: (registered: boolean) => void;
    isOverlayRuntimeInitialized: () => boolean;
    isOverlayShortcutContextActive: () => boolean;
    showMpvOsd: (text: string) => void;
    openRuntimeOptionsPalette: () => void;
    openJimaku: () => void;
    markAudioCard: () => void | Promise<void>;
    copySubtitle: () => void | Promise<void>;
    toggleSecondarySubMode: () => void;
    updateLastCardFromClipboard: () => void | Promise<void>;
    triggerFieldGrouping: () => void | Promise<void>;
    triggerSubsyncFromConfig: () => void | Promise<void>;
    mineSentenceCard: () => void | Promise<void>;
  };
}

export function createShortcutsRuntime(input: ShortcutsRuntimeInput): ShortcutsRuntime {
  const shortcutsRuntime = composeShortcutRuntimes({
    globalShortcuts: input.globalShortcuts,
    numericShortcutRuntimeMainDeps: input.numericShortcutRuntimeMainDeps,
    numericSessions: input.numericSessions,
    overlayShortcutsRuntimeMainDeps: input.overlayShortcutsRuntimeMainDeps,
  });

  return {
    ...shortcutsRuntime,
    syncOverlayShortcutsForModal: (isActive: boolean) => {
      if (isActive) {
        shortcutsRuntime.unregisterOverlayShortcuts();
        return;
      }
      shortcutsRuntime.syncOverlayShortcuts();
    },
  };
}

export interface ShortcutsRuntimeBootstrap {
  shortcuts: ShortcutsRuntime;
  overlayShortcutsRuntime: ReturnType<typeof createOverlayShortcutsRuntimeService>;
  syncOverlayShortcutsForModal: (isActive: boolean) => void;
}

export interface ShortcutsRuntimeFromMainStateInput {
  appState: Pick<AppState, 'overlayRuntimeInitialized' | 'shortcutsRegistered' | 'windowTracker'>;
  getResolvedConfig: () => ResolvedConfig;
  globalShortcut: NumericShortcutRuntimeMainDepsInput['globalShortcut'] & {
    unregisterAll: () => void;
  };
  registerGlobalShortcutsCore: typeof import('../core/services').registerGlobalShortcuts;
  isDev: boolean;
  overlay: {
    getOverlayUi: () =>
      | {
          toggleVisibleOverlay: () => void;
          openRuntimeOptionsPalette: () => void;
        }
      | null
      | undefined;
    overlayManager: {
      getMainWindow: () => Electron.BrowserWindow | null;
      getVisibleOverlayVisible: () => boolean;
    };
    overlayModalRuntime: Pick<OverlayModalRuntime, 'sendToActiveOverlayWindow'>;
  };
  actions: {
    showMpvOsd: (text: string) => void;
    openYomitanSettings: () => boolean;
    triggerSubsyncFromConfig: () => Promise<void>;
    handleCycleSecondarySubMode: () => void;
    handleMultiCopyDigit: (count: number) => void;
  };
  mining: {
    copyCurrentSubtitle: () => void;
    handleMineSentenceDigit: (count: number) => void;
    markLastCardAsAudioCard: () => Promise<void>;
    mineSentenceCard: () => Promise<void>;
    triggerFieldGrouping: () => Promise<void>;
    updateLastCardFromClipboard: () => Promise<void>;
  };
}

export function createShortcutsRuntimeBootstrap(
  input: ShortcutsRuntimeBootstrapInput,
): ShortcutsRuntimeBootstrap {
  let shortcuts: ShortcutsRuntime;

  const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService(
    createBuildOverlayShortcutsRuntimeMainDepsHandler({
      getConfiguredShortcuts: () => input.overlayShortcuts.getConfiguredShortcuts(),
      getShortcutsRegistered: () => input.overlayShortcuts.getShortcutsRegistered(),
      setShortcutsRegistered: (registered: boolean) => {
        input.overlayShortcuts.setShortcutsRegistered(registered);
      },
      isOverlayRuntimeInitialized: () => input.overlayShortcuts.isOverlayRuntimeInitialized(),
      isOverlayShortcutContextActive: () => input.overlayShortcuts.isOverlayShortcutContextActive(),
      showMpvOsd: (text: string) => input.overlayShortcuts.showMpvOsd(text),
      openRuntimeOptionsPalette: () => {
        input.overlayShortcuts.openRuntimeOptionsPalette();
      },
      openJimaku: () => {
        input.overlayShortcuts.openJimaku();
      },
      markAudioCard: () => Promise.resolve(input.overlayShortcuts.markAudioCard()),
      copySubtitleMultiple: (timeoutMs: number) => {
        shortcuts.startPendingMultiCopy(timeoutMs);
      },
      copySubtitle: () => Promise.resolve(input.overlayShortcuts.copySubtitle()),
      toggleSecondarySubMode: () => input.overlayShortcuts.toggleSecondarySubMode(),
      updateLastCardFromClipboard: () =>
        Promise.resolve(input.overlayShortcuts.updateLastCardFromClipboard()),
      triggerFieldGrouping: () => Promise.resolve(input.overlayShortcuts.triggerFieldGrouping()),
      triggerSubsyncFromConfig: () =>
        Promise.resolve(input.overlayShortcuts.triggerSubsyncFromConfig()),
      mineSentenceCard: () => Promise.resolve(input.overlayShortcuts.mineSentenceCard()),
      mineSentenceMultiple: (timeoutMs: number) => {
        shortcuts.startPendingMineSentenceMultiple(timeoutMs);
      },
      cancelPendingMultiCopy: () => {
        shortcuts.cancelPendingMultiCopy();
      },
      cancelPendingMineSentenceMultiple: () => {
        shortcuts.cancelPendingMineSentenceMultiple();
      },
    })(),
  );

  shortcuts = createShortcutsRuntime({
    globalShortcuts: input.globalShortcuts,
    numericShortcutRuntimeMainDeps: input.numericShortcutRuntimeMainDeps,
    numericSessions: input.numericSessions,
    overlayShortcutsRuntimeMainDeps: {
      overlayShortcutsRuntime,
    },
  });

  return {
    shortcuts,
    overlayShortcutsRuntime,
    syncOverlayShortcutsForModal: (isActive: boolean) => {
      shortcuts.syncOverlayShortcutsForModal(isActive);
    },
  };
}

export function createShortcutsRuntimeFromMainState(
  input: ShortcutsRuntimeFromMainStateInput,
): ShortcutsRuntimeBootstrap {
  let shortcuts: ShortcutsRuntime;

  const bootstrap = createShortcutsRuntimeBootstrap({
    globalShortcuts: {
      getConfiguredShortcutsMainDeps: {
        getResolvedConfig: () => input.getResolvedConfig(),
        defaultConfig: DEFAULT_CONFIG,
        resolveConfiguredShortcuts,
      },
      buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
        getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
        registerGlobalShortcutsCore: input.registerGlobalShortcutsCore,
        toggleVisibleOverlay: () => input.overlay.getOverlayUi()?.toggleVisibleOverlay(),
        openYomitanSettings: () => {
          input.actions.openYomitanSettings();
        },
        isDev: input.isDev,
        getMainWindow: () => input.overlay.overlayManager.getMainWindow(),
      }),
      buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcutsHandler) => ({
        unregisterAllGlobalShortcuts: () => input.globalShortcut.unregisterAll(),
        registerGlobalShortcuts: () => registerGlobalShortcutsHandler(),
        syncOverlayShortcuts: () => shortcuts.syncOverlayShortcuts(),
      }),
    },
    numericShortcutRuntimeMainDeps: {
      globalShortcut: input.globalShortcut,
      showMpvOsd: (text) => input.actions.showMpvOsd(text),
      setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
      clearTimer: (timer) => clearTimeout(timer),
    },
    numericSessions: {
      onMultiCopyDigit: (count) => input.actions.handleMultiCopyDigit(count),
      onMineSentenceDigit: (count) => input.mining.handleMineSentenceDigit(count),
    },
    overlayShortcuts: {
      getConfiguredShortcuts: () => shortcuts.getConfiguredShortcuts(),
      getShortcutsRegistered: () => input.appState.shortcutsRegistered,
      setShortcutsRegistered: (registered: boolean) => {
        input.appState.shortcutsRegistered = registered;
      },
      isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
      isOverlayShortcutContextActive: () => {
        if (process.platform !== 'win32') {
          return true;
        }

        if (!input.overlay.overlayManager.getVisibleOverlayVisible()) {
          return false;
        }

        const windowTracker = input.appState.windowTracker;
        if (!windowTracker || !windowTracker.isTracking()) {
          return false;
        }

        return windowTracker.isTargetWindowFocused();
      },
      showMpvOsd: (text: string) => input.actions.showMpvOsd(text),
      openRuntimeOptionsPalette: () => {
        input.overlay.getOverlayUi()?.openRuntimeOptionsPalette();
      },
      openJimaku: () => {
        input.overlay.overlayModalRuntime.sendToActiveOverlayWindow('jimaku:open', undefined, {
          restoreOnModalClose: 'jimaku' as OverlayHostedModal,
        });
      },
      markAudioCard: () => input.mining.markLastCardAsAudioCard(),
      copySubtitle: () => input.mining.copyCurrentSubtitle(),
      toggleSecondarySubMode: () => input.actions.handleCycleSecondarySubMode(),
      updateLastCardFromClipboard: () => input.mining.updateLastCardFromClipboard(),
      triggerFieldGrouping: () => input.mining.triggerFieldGrouping(),
      triggerSubsyncFromConfig: () => input.actions.triggerSubsyncFromConfig(),
      mineSentenceCard: () => input.mining.mineSentenceCard(),
    },
  });

  shortcuts = bootstrap.shortcuts;
  return bootstrap;
}
