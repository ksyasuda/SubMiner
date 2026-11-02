import type { GlobalShortcut } from 'electron';
import { DEFAULT_CONFIG } from '../../config/definitions';
import {
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  type createOverlayManager,
} from '../../core/services';
import { resolveConfiguredShortcuts } from '../../core/utils';
import type { Logger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { ResolvedConfig } from '../../types';
import { createOverlayShortcutsRuntimeService } from '../overlay-shortcuts-runtime';
import { composeShortcutRuntimes } from '../runtime/composers';
import { tryBeginVisibleOverlayNumericSelection } from '../runtime/overlay-numeric-selection';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { readMpvInputBindings } from '../runtime/mpv-input-bindings';
import { createSessionBindingsRuntime } from '../runtime/session-bindings-runtime';
import type { AppState } from '../state';
import type { MiningAssembly } from './mining';
import type { OverlayModalsAssembly } from './overlay-modals';
import type { CliSlice, IpcSlice } from './slices';

type OverlayManager = ReturnType<typeof createOverlayManager>;

/**
 * What shortcut handling needs from the rest of the app. Siblings built after this assembly
 * in `main.ts` are reached through callbacks.
 */
export interface ShortcutsAssemblyDeps {
  appState: Pick<
    AppState,
    | 'shortcutsRegistered'
    | 'overlayRuntimeInitialized'
    | 'windowTracker'
    | 'keybindings'
    | 'mpvClient'
    | 'sessionBindings'
    | 'sessionBindingsInitialized'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'warn'>;
  configDir: string;
  isDev: boolean;
  overlayManager: Pick<
    OverlayManager,
    'getMainWindow' | 'getVisibleOverlayVisible' | 'broadcastToOverlayWindows'
  >;
  /** Electron's globalShortcut, injected so the assembly loads under `bun test`. */
  globalShortcut: GlobalShortcut;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification' | 'showOverlayNotification'
  >;
  getOverlayModals: () => Pick<
    OverlayModalsAssembly,
    | 'openRuntimeOptionsPalette'
    | 'openCharacterDictionaryManagerOverlay'
    | 'openJimakuOverlay'
    | 'openTsukihimeOverlay'
  >;
  getMining: () => Pick<
    MiningAssembly,
    | 'markLastCardAsAudioCard'
    | 'copyCurrentSubtitle'
    | 'copySubtitleCount'
    | 'updateLastCardFromClipboard'
    | 'triggerFieldGrouping'
    | 'mineSentenceCard'
    | 'mineSentenceCount'
  >;
  triggerSubsyncFromConfig: () => Promise<void>;
  cycleSecondarySubMode: () => void;
  toggleVisibleOverlay: () => void;
  openDictionarySettings: () => void;
}

export type ShortcutsAssembly = ReturnType<typeof createShortcutsAssembly>;

/**
 * Owns keyboard input outside the renderer: the overlay's local shortcuts, global shortcuts,
 * the numeric multi-copy/multi-mine sessions, and the session bindings shared with mpv.
 */
export function createShortcutsAssembly(deps: ShortcutsAssemblyDeps) {
  const { appState, overlayManager } = deps;
  const showStatus = (text: string) =>
    deps.getOverlayNotifications().showConfiguredStatusNotification(text);

  const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService({
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getShortcutsRegistered: () => appState.shortcutsRegistered,
    setShortcutsRegistered: (registered: boolean) => {
      appState.shortcutsRegistered = registered;
    },
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    isOverlayShortcutContextActive: () => {
      if (process.platform !== 'win32') {
        return true;
      }

      if (!overlayManager.getVisibleOverlayVisible()) {
        return false;
      }

      const windowTracker = appState.windowTracker;
      if (!windowTracker || !windowTracker.isTracking()) {
        return false;
      }

      return windowTracker.isTargetWindowFocused();
    },
    showMpvOsd: showStatus,
    openRuntimeOptionsPalette: () => {
      deps.getOverlayModals().openRuntimeOptionsPalette();
    },
    openCharacterDictionaryManager: () => {
      deps.getOverlayModals().openCharacterDictionaryManagerOverlay();
    },
    openJimaku: () => {
      deps.getOverlayModals().openJimakuOverlay();
    },
    openTsukihime: () => {
      deps.getOverlayModals().openTsukihimeOverlay();
    },
    markAudioCard: () => deps.getMining().markLastCardAsAudioCard(),
    copySubtitleMultiple: (timeoutMs: number) => {
      startPendingMultiCopy(timeoutMs);
    },
    copySubtitle: () => {
      deps.getMining().copyCurrentSubtitle();
    },
    toggleSecondarySubMode: () => deps.cycleSecondarySubMode(),
    updateLastCardFromClipboard: () => deps.getMining().updateLastCardFromClipboard(),
    triggerFieldGrouping: () => deps.getMining().triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    mineSentenceCard: () => deps.getMining().mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs: number) => {
      startPendingMineSentenceMultiple(timeoutMs);
    },
    cancelPendingMultiCopy: () => {
      cancelPendingMultiCopy();
    },
    cancelPendingMineSentenceMultiple: () => {
      cancelPendingMineSentenceMultiple();
    },
  });

  const {
    getConfiguredShortcuts,
    registerGlobalShortcuts,
    refreshGlobalAndOverlayShortcuts,
    cancelPendingMultiCopy,
    startPendingMultiCopy,
    cancelPendingMineSentenceMultiple,
    startPendingMineSentenceMultiple,
    syncOverlayShortcuts,
    refreshOverlayShortcuts,
  } = composeShortcutRuntimes({
    globalShortcuts: {
      getConfiguredShortcutsMainDeps: {
        getResolvedConfig: () => deps.getConfig(),
        defaultConfig: DEFAULT_CONFIG,
        resolveConfiguredShortcuts,
      },
      buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
        getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
        registerGlobalShortcutsCore,
        onToggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
        onOpenYomitanSettings: () => deps.openDictionarySettings(),
        isDev: deps.isDev,
        getMainWindow: () => overlayManager.getMainWindow(),
      }),
      buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcutsHandler) => ({
        unregisterAllGlobalShortcuts: () => deps.globalShortcut.unregisterAll(),
        registerGlobalShortcuts: () => registerGlobalShortcutsHandler(),
        syncOverlayShortcuts: () => syncOverlayShortcuts(),
      }),
    },
    numericShortcutRuntimeMainDeps: {
      globalShortcut: deps.globalShortcut,
      showMpvOsd: showStatus,
      setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
      clearTimer: (timer) => clearTimeout(timer),
    },
    numericSessions: {
      onMultiCopyDigit: (count) => deps.getMining().copySubtitleCount(count),
      onMineSentenceDigit: (count) => deps.getMining().mineSentenceCount(count),
      tryBeginMultiCopyOverlaySelection: (timeoutMs) =>
        tryBeginVisibleOverlayNumericSelection({
          actionId: 'copySubtitleMultiple',
          timeoutMs,
          getMainWindow: () => overlayManager.getMainWindow(),
          getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        }),
      tryBeginMineSentenceOverlaySelection: (timeoutMs) =>
        tryBeginVisibleOverlayNumericSelection({
          actionId: 'mineSentenceMultiple',
          timeoutMs,
          getMainWindow: () => overlayManager.getMainWindow(),
          getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        }),
    },
    overlayShortcutsRuntimeMainDeps: {
      overlayShortcutsRuntime,
    },
  });

  const { persistSessionBindings, refreshCurrentSessionBindings, refreshMpvSessionBindings } =
    createSessionBindingsRuntime({
      configDir: deps.configDir,
      getKeybindings: () => appState.keybindings,
      getConfiguredShortcuts: () => getConfiguredShortcuts(),
      getResolvedConfig: () => deps.getConfig(),
      getMpvClient: () => appState.mpvClient,
      setSessionBindings: (bindings) => {
        appState.sessionBindings = bindings;
      },
      setSessionBindingsInitialized: (initialized) => {
        appState.sessionBindingsInitialized = initialized;
      },
      logWarn: (message) => deps.logger.warn(message),
      onBindingsChanged: (bindings) =>
        overlayManager.broadcastToOverlayWindows(
          IPC_CHANNELS.event.sessionBindingsChanged,
          bindings,
        ),
      onMpvInputBindingsChanged: () =>
        overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.mpvInputBindingsChanged),
      onWarning: (warning) => {
        if (warning.kind !== 'conflict') return;
        deps.getOverlayNotifications().showOverlayNotification({
          id: `session-binding-conflict:${warning.path}`,
          title: 'Shortcut conflict',
          body: warning.message,
          variant: 'warning',
        });
      },
    });

  const ipc: IpcSlice<'getMpvInputBindings' | 'getSessionBindings' | 'getConfiguredShortcuts'> = {
    getMpvInputBindings: async () => {
      await refreshMpvSessionBindings();
      return readMpvInputBindings({
        getMpvClient: () => appState.mpvClient,
        getConfiguredKeybindings: () => deps.getConfig().keybindings ?? [],
        platform:
          process.platform === 'darwin'
            ? 'darwin'
            : process.platform === 'win32'
              ? 'win32'
              : 'linux',
      });
    },
    getSessionBindings: async () => {
      await refreshMpvSessionBindings();
      return appState.sessionBindings;
    },
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
  };

  const cli: CliSlice<
    'startPendingMultiCopy' | 'startPendingMineSentenceMultiple' | 'getMultiCopyTimeoutMs'
  > = {
    startPendingMultiCopy: (timeoutMs) => startPendingMultiCopy(timeoutMs),
    startPendingMineSentenceMultiple: (timeoutMs) => startPendingMineSentenceMultiple(timeoutMs),
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
  };

  return {
    ipc,
    cli,
    overlayShortcutsRuntime,
    getConfiguredShortcuts,
    registerGlobalShortcuts,
    refreshGlobalAndOverlayShortcuts,
    startPendingMultiCopy,
    startPendingMineSentenceMultiple,
    refreshOverlayShortcuts,
    persistSessionBindings,
    refreshCurrentSessionBindings,
    refreshMpvSessionBindings,
    /** Overlay shortcuts are released while a modal owns keyboard input. */
    syncOverlayShortcutsForModal: (isActive: boolean): void => {
      if (isActive) {
        overlayShortcutsRuntime.unregisterOverlayShortcuts();
      } else {
        overlayShortcutsRuntime.syncOverlayShortcuts();
      }
    },
  };
}
