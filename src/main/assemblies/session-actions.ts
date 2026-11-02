import {
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  playNextSubtitleRuntime,
  replayCurrentSubtitleRuntime,
  sendMpvCommandRuntime,
  type createOverlayManager,
} from '../../core/services';
import { applyRuntimeOptionResultRuntime } from '../../core/services/runtime-options-ipc';
import { dispatchSessionAction as dispatchSessionActionCore } from '../../core/services/session-actions';
import type { Logger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { ResolvedConfig, SecondarySubMode } from '../../types';
import type { RuntimeOptionApplyResult, RuntimeOptionId } from '../../types/runtime-options';
import type { SessionActionDispatchRequest } from '../../types/runtime';
import type { MpvCommandFromIpcRuntimeDeps } from '../ipc-mpv-command';
import { appendClipboardVideoToQueueRuntime } from '../runtime/clipboard-queue';
import { createAppendClipboardVideoToQueueHandler } from '../runtime/overlay-main-actions';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createCycleSecondarySubModeRuntimeHandler } from '../runtime/secondary-sub-mode-runtime-handler';
import { openSubtitleSelectionModal } from '../runtime/subtitle-selection';
import type { AppState } from '../state';
import type { MiningAssembly } from './mining';
import type { OverlayModalsAssembly } from './overlay-modals';
import type { PluginsAssembly } from './plugins';
import type { StatsAssembly } from './stats';
import type { YoutubeAssembly } from './youtube';

/** What session actions call into; every sibling here is built before the assembly. */
export interface SessionActionsAssemblyDeps {
  appState: AppState;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'warn'>;
  overlayManager: Pick<ReturnType<typeof createOverlayManager>, 'broadcastToOverlayWindows'>;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    | 'showConfiguredPlaybackFeedback'
    | 'showConfiguredStatusNotification'
    | 'toggleNotificationHistoryPanel'
  >;
  stats: Pick<StatsAssembly, 'toggleStatsOverlay' | 'markActiveVideoWatched'>;
  mining: Pick<
    MiningAssembly,
    | 'copyCurrentSubtitle'
    | 'copySubtitleCount'
    | 'updateLastCardFromClipboard'
    | 'triggerFieldGrouping'
    | 'mineSentenceCard'
    | 'mineSentenceCount'
    | 'markLastCardAsAudioCard'
  >;
  overlayModals: Pick<
    OverlayModalsAssembly,
    | 'openRuntimeOptionsPalette'
    | 'openJimakuOverlay'
    | 'openTsukihimeOverlay'
    | 'openSessionHelpOverlay'
    | 'openOverlayHostedModalWithOsd'
    | 'openSubtitleGenerationOverlay'
    | 'openCharacterDictionaryManagerOverlay'
    | 'openControllerSelectOverlay'
    | 'openControllerDebugOverlay'
    | 'openPlaylistBrowser'
  >;
  youtube: Pick<YoutubeAssembly, 'openTrackPickerFromPlayback'>;
  plugins: { host: Pick<PluginsAssembly['host'], 'runAction'> };
  toggleVisibleOverlay: () => void;
  triggerSubsyncFromConfig: () => Promise<void>;
  readClipboardText: () => string;
  /** Raw mpv OSD, bypassing the configured notification type. */
  showMpvOsd: (text: string) => void;
  /** Renderer-issued mpv commands; playback sanitizes subtitle seeks. */
  sendRendererMpvCommand: (command: (string | number)[]) => void;
}

export type SessionActionsAssembly = ReturnType<typeof createSessionActionsAssembly>;

/**
 * Owns what a shortcut, controller button, CLI command or plugin can trigger in the running
 * session: the session action dispatcher, the secondary-subtitle mode cycle, queueing the
 * clipboard's video, and the mpv commands the overlay renderer sends.
 */
export function createSessionActionsAssembly(deps: SessionActionsAssemblyDeps) {
  const { appState, overlayManager, overlayModals, mining } = deps;

  function setSecondarySubMode(mode: SecondarySubMode): void {
    appState.secondarySubMode = mode;
  }

  const cycleSecondarySubMode = createCycleSecondarySubModeRuntimeHandler({
    cycleSecondarySubModeMainDeps: {
      getSecondarySubMode: () => appState.secondarySubMode,
      setSecondarySubMode: (mode: SecondarySubMode) => {
        setSecondarySubMode(mode);
      },
      getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
      setLastSecondarySubToggleAtMs: (timestampMs: number) => {
        appState.lastSecondarySubToggleAtMs = timestampMs;
      },
      broadcastSecondarySubMode: (mode) => {
        overlayManager.broadcastToOverlayWindows('secondary-subtitle:mode', mode);
      },
      showMpvOsd: (text: string) =>
        deps.getOverlayNotifications().showConfiguredPlaybackFeedback(text),
    },
    cycleSecondarySubMode: (cycleDeps) => cycleSecondarySubModeCore(cycleDeps),
  });

  const appendClipboardVideoToQueue = createAppendClipboardVideoToQueueHandler({
    appendClipboardVideoToQueueRuntime,
    getMpvClient: () => appState.mpvClient,
    readClipboardText: () => deps.readClipboardText(),
    showMpvOsd: (text) => deps.getOverlayNotifications().showConfiguredStatusNotification(text),
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
  });

  /** Steps a runtime option and reports the new value as playback feedback. */
  function cycleRuntimeOption(id: RuntimeOptionId, direction: 1 | -1): RuntimeOptionApplyResult {
    if (!appState.runtimeOptionsManager) {
      return { ok: false, error: 'Runtime options manager unavailable' };
    }
    return applyRuntimeOptionResultRuntime(
      appState.runtimeOptionsManager.cycleOption(id, direction),
      (text) => deps.getOverlayNotifications().showConfiguredPlaybackFeedback(text),
    );
  }

  async function dispatchSessionAction(request: SessionActionDispatchRequest): Promise<void> {
    await dispatchSessionActionCore(request, {
      toggleStatsOverlay: async () => await deps.stats.toggleStatsOverlay(),
      toggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
      copyCurrentSubtitle: () => mining.copyCurrentSubtitle(),
      copySubtitleCount: (count) => mining.copySubtitleCount(count),
      updateLastCardFromClipboard: () => mining.updateLastCardFromClipboard(),
      triggerFieldGrouping: () => mining.triggerFieldGrouping(),
      triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
      mineSentenceCard: () => mining.mineSentenceCard(),
      mineSentenceCount: (count) => mining.mineSentenceCount(count),
      toggleSecondarySub: () => cycleSecondarySubMode(),
      toggleSubtitleSidebar: () =>
        overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.subtitleSidebarToggle),
      toggleNotificationHistory: () =>
        deps.getOverlayNotifications().toggleNotificationHistoryPanel(),
      appendClipboardVideoToQueue: () => {
        appendClipboardVideoToQueue();
      },
      markLastCardAsAudioCard: () => mining.markLastCardAsAudioCard(),
      markActiveVideoWatched: () => deps.stats.markActiveVideoWatched(),
      openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
      openJimaku: () => overlayModals.openJimakuOverlay(),
      openTsukihime: () => overlayModals.openTsukihimeOverlay(),
      openSessionHelp: () => overlayModals.openSessionHelpOverlay(),
      openSubtitleSelection: () => {
        if (!deps.getConfig().subtitleSelection.enabled) return;
        overlayModals.openOverlayHostedModalWithOsd(
          openSubtitleSelectionModal,
          'Subtitle selection overlay unavailable.',
          'Failed to open subtitle selection overlay.',
        );
      },
      openSubtitleGeneration: () => overlayModals.openSubtitleGenerationOverlay(),
      openCharacterDictionaryManager: () => overlayModals.openCharacterDictionaryManagerOverlay(),
      openControllerSelect: () => overlayModals.openControllerSelectOverlay(),
      openControllerDebug: () => overlayModals.openControllerDebugOverlay(),
      openYoutubeTrackPicker: () => deps.youtube.openTrackPickerFromPlayback(),
      openPlaylistBrowser: () => overlayModals.openPlaylistBrowser(),
      replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
      playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
      cycleRuntimeOption,
      playNextPlaylistItem: () =>
        sendMpvCommandRuntime(appState.mpvClient, ['playlist-next', 'force']),
      runPluginAction: async (pluginActionId) => {
        const ran = await deps.plugins.host.runAction(pluginActionId);
        if (!ran) deps.logger.warn(`Plugin action is not registered: ${pluginActionId}`);
        return ran;
      },
      showMpvOsd: (text) => deps.getOverlayNotifications().showConfiguredPlaybackFeedback(text),
    });
  }

  /** The mpv commands the overlay renderer sends over IPC (`mpv:command`). */
  const mpvCommandDeps: MpvCommandFromIpcRuntimeDeps = {
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
    openJimaku: () => overlayModals.openJimakuOverlay(),
    openTsukihime: () => overlayModals.openTsukihimeOverlay(),
    openYoutubeTrackPicker: () => deps.youtube.openTrackPickerFromPlayback(),
    openPlaylistBrowser: () => overlayModals.openPlaylistBrowser(),
    cycleRuntimeOption,
    showMpvOsd: (text: string) =>
      deps.getOverlayNotifications().showConfiguredStatusNotification(text),
    showRawMpvOsd: (text: string) => deps.showMpvOsd(text),
    showPlaybackFeedback: (text: string) =>
      deps.getOverlayNotifications().showConfiguredPlaybackFeedback(text),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    sendMpvCommand: (rawCommand: (string | number)[]) => deps.sendRendererMpvCommand(rawCommand),
    getMpvClient: () => appState.mpvClient,
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  };

  return { dispatchSessionAction, cycleSecondarySubMode, setSecondarySubMode, mpvCommandDeps };
}
