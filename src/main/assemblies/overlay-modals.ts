import type { createLogger } from '../../logger';
import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import type { ResolvedConfig } from '../../types';
import type { OverlayModalRuntime } from '../overlay-runtime';
import { openChangelogModal as openChangelogModalRuntime } from '../runtime/changelog-open';
import { openCharacterDictionaryManagerWithConfigGate } from '../runtime/character-dictionary-manager-gate';
import { openCharacterDictionaryManagerModal as openCharacterDictionaryManagerModalRuntime } from '../runtime/character-dictionary-open';
import { openControllerDebugModal as openControllerDebugModalRuntime } from '../runtime/controller-debug-open';
import { openControllerSelectModal as openControllerSelectModalRuntime } from '../runtime/controller-select-open';
import { openJimakuModal as openJimakuModalRuntime } from '../runtime/jimaku-open';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { openPlaylistBrowser as openPlaylistBrowserRuntime } from '../runtime/playlist-browser-open';
import { openRuntimeOptionsModal as openRuntimeOptionsModalRuntime } from '../runtime/runtime-options-open';
import { openSessionHelpModal as openSessionHelpModalRuntime } from '../runtime/session-help-open';
import { openSubtitleGenerationModal } from '../runtime/subtitle-generation-open';
import { openTsukihimeModal as openTsukihimeModalRuntime } from '../runtime/tsukihime-open';
import type { AppState } from '../state';

type OverlayNotifications = Pick<
  ReturnType<typeof createOverlayNotificationsRuntime>,
  | 'getConfiguredStatusNotificationType'
  | 'showConfiguredStatusNotification'
  | 'showOverlayNotification'
>;

/** What the overlay-hosted modal openers need from the rest of the app. */
export interface OverlayModalsAssemblyDeps {
  appState: Pick<AppState, 'mpvClient' | 'currentMediaPath'>;
  getConfig: () => ResolvedConfig;
  logger: Pick<ReturnType<typeof createLogger>, 'warn' | 'error'>;
  overlayModalRuntime: Pick<OverlayModalRuntime, 'waitForModalOpen'>;
  getOverlayNotifications: () => OverlayNotifications;
  ensureOverlayStartupPrereqs: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
  showMpvOsd: (text: string) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
}

export type OverlayModalsAssembly = ReturnType<typeof createOverlayModalsAssembly>;

/**
 * Opens the overlay-hosted modals (runtime options, Jimaku, session help, playlist browser, ...)
 * and reports through the configured status notification when one cannot open.
 */
export function createOverlayModalsAssembly(deps: OverlayModalsAssemblyDeps) {
  const { appState, logger } = deps;

  function createOverlayHostedModalOpenDeps(): {
    ensureOverlayStartupPrereqs: () => void;
    ensureOverlayWindowsReadyForVisibilityActions: () => void;
    sendToActiveOverlayWindow: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: {
        restoreOnModalClose?: OverlayHostedModal;
        preferModalWindow?: boolean;
      },
    ) => boolean;
    waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
    logWarn: (message: string) => void;
  } {
    return {
      ensureOverlayStartupPrereqs: () => deps.ensureOverlayStartupPrereqs(),
      ensureOverlayWindowsReadyForVisibilityActions: () =>
        deps.ensureOverlayWindowsReadyForVisibilityActions(),
      sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
        deps.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
      waitForModalOpen: (modal, timeoutMs) =>
        deps.overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
      logWarn: (message) => logger.warn(message),
    };
  }

  function openOverlayHostedModalWithOsd(
    openModal: (deps: ReturnType<typeof createOverlayHostedModalOpenDeps>) => Promise<boolean>,
    unavailableMessage: string,
    failureLogMessage: string,
  ): void {
    void openModal(createOverlayHostedModalOpenDeps())
      .then((opened) => {
        if (!opened) {
          deps.getOverlayNotifications().showConfiguredStatusNotification(unavailableMessage, {
            variant: 'warning',
          });
        }
      })
      .catch((error) => {
        logger.error(failureLogMessage, error);
        deps.getOverlayNotifications().showConfiguredStatusNotification(unavailableMessage, {
          variant: 'error',
        });
      });
  }

  function openRuntimeOptionsPalette(): void {
    openOverlayHostedModalWithOsd(
      openRuntimeOptionsModalRuntime,
      'Runtime options overlay unavailable.',
      'Failed to open runtime options overlay.',
    );
  }

  function openJimakuOverlay(): void {
    openOverlayHostedModalWithOsd(
      openJimakuModalRuntime,
      'Jimaku overlay unavailable.',
      'Failed to open Jimaku overlay.',
    );
  }

  function openTsukihimeOverlay(): void {
    openOverlayHostedModalWithOsd(
      openTsukihimeModalRuntime,
      'TsukiHime overlay unavailable.',
      'Failed to open TsukiHime overlay.',
    );
  }

  function openSubtitleGenerationOverlay(): void {
    openOverlayHostedModalWithOsd(
      openSubtitleGenerationModal,
      'Subtitle generation overlay unavailable.',
      'Failed to open subtitle generation overlay.',
    );
  }

  // currentMediaPath follows every mpv path change and is cleared on disconnect; the
  // client's currentVideoPath outlives the connection, so a reconnect could reuse it.
  function isMediaPlaybackActive(): boolean {
    return appState.mpvClient?.connected === true && Boolean(appState.currentMediaPath?.trim());
  }

  function openSessionHelpOverlay(): void {
    openOverlayHostedModalWithOsd(
      (modalDeps) => openSessionHelpModalRuntime({ ...modalDeps, isMediaPlaybackActive }),
      'Session help overlay unavailable.',
      'Failed to open session help overlay.',
    );
  }

  function openChangelogOverlay(): void {
    openOverlayHostedModalWithOsd(
      openChangelogModalRuntime,
      'Changelog overlay unavailable.',
      'Failed to open changelog overlay.',
    );
  }

  function openCharacterDictionaryManagerOverlay(): void {
    openCharacterDictionaryManagerWithConfigGate({
      isCharacterDictionaryEnabled: () => deps.getConfig().subtitleStyle.nameMatchEnabled,
      getNotificationType: () =>
        deps.getOverlayNotifications().getConfiguredStatusNotificationType(),
      openManager: () => {
        openOverlayHostedModalWithOsd(
          openCharacterDictionaryManagerModalRuntime,
          'Character dictionary manager unavailable.',
          'Failed to open character dictionary manager.',
        );
      },
      showOsd: (message) => deps.showMpvOsd(message),
      showOverlayNotification: (payload) =>
        deps.getOverlayNotifications().showOverlayNotification(payload),
      showDesktopNotification: (title, options) => deps.showDesktopNotification(title, options),
      logWarn: (message, error) => logger.warn(message, error),
    });
  }

  function openControllerSelectOverlay(): void {
    openOverlayHostedModalWithOsd(
      openControllerSelectModalRuntime,
      'Controller select overlay unavailable.',
      'Failed to open controller select overlay.',
    );
  }

  function openControllerDebugOverlay(): void {
    openOverlayHostedModalWithOsd(
      openControllerDebugModalRuntime,
      'Controller debug overlay unavailable.',
      'Failed to open controller debug overlay.',
    );
  }

  function openPlaylistBrowser(): void {
    if (!appState.mpvClient?.connected) {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification('Playlist browser requires active playback.', {
          title: 'Playlist browser',
          variant: 'warning',
        });
      return;
    }
    openOverlayHostedModalWithOsd(
      openPlaylistBrowserRuntime,
      'Playlist browser overlay unavailable.',
      'Failed to open playlist browser overlay.',
    );
  }

  return {
    createOverlayHostedModalOpenDeps,
    openOverlayHostedModalWithOsd,
    openRuntimeOptionsPalette,
    openJimakuOverlay,
    openTsukihimeOverlay,
    openSubtitleGenerationOverlay,
    openSessionHelpOverlay,
    openChangelogOverlay,
    openCharacterDictionaryManagerOverlay,
    openControllerSelectOverlay,
    openControllerDebugOverlay,
    openPlaylistBrowser,
  };
}
