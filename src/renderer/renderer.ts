/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import type {
  KikuDuplicateCardInfo,
  RuntimeOptionState,
  SecondarySubMode,
  SubtitleData,
  SubtitlePosition,
  SubsyncManualPayload,
  ConfigHotReloadPayload,
  MiningImagePayload,
} from '../types';
import { createKeyboardHandlers } from './handlers/keyboard.js';
import { createGamepadController } from './handlers/gamepad-controller.js';
import { createMouseHandlers } from './handlers/mouse.js';
import { createControllerStatusIndicator } from './controller-status-indicator.js';
import { createControllerDebugModal } from './modals/controller-debug.js';
import { createControllerSelectModal } from './modals/controller-select.js';
import { createJimakuModal } from './modals/jimaku.js';
import { createKikuModal } from './modals/kiku.js';
import { prepareForKikuFieldGroupingOpen } from './kiku-open.js';
import { createPlaylistBrowserModal } from './modals/playlist-browser.js';
import { createSessionHelpModal } from './modals/session-help.js';
import { createSubtitleSidebarModal } from './modals/subtitle-sidebar.js';
import { isControllerInteractionBlocked } from './controller-interaction-blocking.js';
import { createCharacterDictionaryModal } from './modals/character-dictionary.js';
import { createRuntimeOptionsModal } from './modals/runtime-options.js';
import { createSubsyncModal } from './modals/subsync.js';
import { createYoutubeTrackPickerModal } from './modals/youtube-track-picker.js';
import { createPositioningController } from './positioning.js';
import { createOverlayContentMeasurementReporter } from './overlay-content-measurement.js';
import { syncOverlayMouseIgnoreState } from './overlay-mouse-ignore.js';
import { createRendererState } from './state.js';
import { createSubtitleRenderer } from './subtitle-render.js';
import { isYomitanPopupVisible, registerYomitanLookupListener } from './yomitan-popup.js';
import {
  createRendererRecoveryController,
  registerRendererGlobalErrorHandlers,
} from './error-recovery.js';
import { resolveRendererDom } from './utils/dom.js';
import { resolvePlatformInfo } from './utils/platform.js';
import {
  buildMpvLoadfileCommands,
  buildMpvSubtitleAddCommands,
  collectDroppedSubtitlePaths,
  collectDroppedVideoPaths,
} from '../core/services/overlay-drop.js';

const ctx = {
  dom: resolveRendererDom(),
  platform: resolvePlatformInfo(),
  state: createRendererState(),
};

function isAnySettingsModalOpen(): boolean {
  return (
    ctx.state.controllerSelectModalOpen ||
    ctx.state.controllerDebugModalOpen ||
    ctx.state.runtimeOptionsModalOpen ||
    ctx.state.characterDictionaryModalOpen ||
    ctx.state.subsyncModalOpen ||
    ctx.state.kikuModalOpen ||
    ctx.state.jimakuModalOpen ||
    ctx.state.youtubePickerModalOpen ||
    ctx.state.sessionHelpModalOpen ||
    ctx.state.playlistBrowserModalOpen
  );
}

function isAnyModalOpen(): boolean {
  return (
    ctx.state.controllerSelectModalOpen ||
    ctx.state.controllerDebugModalOpen ||
    ctx.state.jimakuModalOpen ||
    ctx.state.kikuModalOpen ||
    ctx.state.runtimeOptionsModalOpen ||
    ctx.state.characterDictionaryModalOpen ||
    ctx.state.subsyncModalOpen ||
    ctx.state.youtubePickerModalOpen ||
    ctx.state.sessionHelpModalOpen ||
    ctx.state.playlistBrowserModalOpen ||
    ctx.state.subtitleSidebarModalOpen
  );
}

function isControllerInputBlocked(): boolean {
  return isControllerInteractionBlocked(ctx.state);
}

function syncSettingsModalSubtitleSuppression(): void {
  const suppressSubtitles = isAnySettingsModalOpen();
  document.body.classList.toggle('settings-modal-open', suppressSubtitles);
  if (suppressSubtitles) {
    ctx.state.isOverSubtitle = false;
  }
}

const subtitleRenderer = createSubtitleRenderer(ctx);
const measurementReporter = createOverlayContentMeasurementReporter(ctx);
const positioning = createPositioningController(ctx);
const runtimeOptionsModal = createRuntimeOptionsModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const characterDictionaryModal = createCharacterDictionaryModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const subsyncModal = createSubsyncModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const controllerSelectModal = createControllerSelectModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
  notifyControllerDisabled: showControllerDisabledNotice,
});
const controllerDebugModal = createControllerDebugModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
  notifyControllerDisabled: showControllerDisabledNotice,
});
const controllerStatusIndicator = createControllerStatusIndicator(ctx.dom);
const sessionHelpModal = createSessionHelpModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const subtitleSidebarModal = createSubtitleSidebarModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  shouldRestoreOpenOnStartup: async () =>
    ctx.platform.overlayLayer === 'visible' && (await window.electronAPI.getSubtitleSidebarOpen()),
  onVisibilityChanged: () => {
    measurementReporter.emitNow();
  },
});
const kikuModal = createKikuModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const jimakuModal = createJimakuModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const mouseHandlers = createMouseHandlers(ctx, {
  modalStateReader: { isAnySettingsModalOpen, isAnyModalOpen },
  applyYPercent: positioning.applyYPercent,
  getCurrentYPercent: positioning.getCurrentYPercent,
  persistSubtitlePositionPatch: positioning.persistSubtitlePositionPatch,
  getSubtitleHoverAutoPauseEnabled: () => ctx.state.autoPauseVideoOnSubtitleHover,
  getYomitanPopupAutoPauseEnabled: () => ctx.state.autoPauseVideoOnYomitanPopup,
  getPlaybackPaused: () => window.electronAPI.getPlaybackPaused(),
  sendMpvCommand: (command) => {
    window.electronAPI.sendMpvCommand(command);
  },
});
const youtubePickerModal = createYoutubeTrackPickerModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  restorePointerInteractionState: mouseHandlers.restorePointerInteractionState,
  syncSettingsModalSubtitleSuppression,
});
const playlistBrowserModal = createPlaylistBrowserModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const keyboardHandlers = createKeyboardHandlers(ctx, {
  handleRuntimeOptionsKeydown: runtimeOptionsModal.handleRuntimeOptionsKeydown,
  handleCharacterDictionaryKeydown: characterDictionaryModal.handleCharacterDictionaryKeydown,
  handleSubsyncKeydown: subsyncModal.handleSubsyncKeydown,
  handleKikuKeydown: kikuModal.handleKikuKeydown,
  handleJimakuKeydown: jimakuModal.handleJimakuKeydown,
  handleYoutubePickerKeydown: youtubePickerModal.handleYoutubePickerKeydown,
  handlePlaylistBrowserKeydown: playlistBrowserModal.handlePlaylistBrowserKeydown,
  handleControllerSelectKeydown: controllerSelectModal.handleControllerSelectKeydown,
  handleControllerDebugKeydown: controllerDebugModal.handleControllerDebugKeydown,
  handleSessionHelpKeydown: sessionHelpModal.handleSessionHelpKeydown,
  openSessionHelpModal: sessionHelpModal.openSessionHelpModal,
  openControllerSelectModal: () => {
    if (controllerSelectModal.openControllerSelectModal()) {
      window.electronAPI.notifyOverlayModalOpened('controller-select');
    }
  },
  openControllerDebugModal: () => {
    if (controllerDebugModal.openControllerDebugModal()) {
      window.electronAPI.notifyOverlayModalOpened('controller-debug');
    }
  },
  appendClipboardVideoToQueue: () => {
    void window.electronAPI.appendClipboardVideoToQueue();
  },
  getPlaybackPaused: () => window.electronAPI.getPlaybackPaused(),
  toggleSubtitleSidebarModal: () => {
    void subtitleSidebarModal.toggleSubtitleSidebarModal();
  },
});

let lastSubtitlePreview = '';
let lastSecondarySubtitlePreview = '';
let overlayErrorToastTimeout: ReturnType<typeof setTimeout> | null = null;
let miningImageToastTimeout: ReturnType<typeof setTimeout> | null = null;
let controllerAnimationFrameId: number | null = null;

function truncateForErrorLog(text: string): string {
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length <= 180) {
    return normalized;
  }
  return `${normalized.slice(0, 177)}...`;
}

function getSubtitleTextForPreview(data: SubtitleData | string): string {
  if (typeof data === 'string') {
    return data;
  }
  if (data && typeof data.text === 'string') {
    return data.text;
  }
  return '';
}

function getActiveModal(): string | null {
  if (ctx.state.controllerSelectModalOpen) return 'controller-select';
  if (ctx.state.controllerDebugModalOpen) return 'controller-debug';
  if (ctx.state.subtitleSidebarModalOpen) return 'subtitle-sidebar';
  if (ctx.state.jimakuModalOpen) return 'jimaku';
  if (ctx.state.youtubePickerModalOpen) return 'youtube-track-picker';
  if (ctx.state.playlistBrowserModalOpen) return 'playlist-browser';
  if (ctx.state.kikuModalOpen) return 'kiku';
  if (ctx.state.runtimeOptionsModalOpen) return 'runtime-options';
  if (ctx.state.characterDictionaryModalOpen) return 'character-dictionary';
  if (ctx.state.subsyncModalOpen) return 'subsync';
  if (ctx.state.sessionHelpModalOpen) return 'session-help';
  return null;
}

function dismissActiveUiAfterError(): void {
  if (ctx.state.controllerSelectModalOpen) {
    controllerSelectModal.closeControllerSelectModal();
  }
  if (ctx.state.controllerDebugModalOpen) {
    controllerDebugModal.closeControllerDebugModal();
  }
  if (ctx.state.subtitleSidebarModalOpen) {
    subtitleSidebarModal.closeSubtitleSidebarModal();
  }
  if (ctx.state.jimakuModalOpen) {
    jimakuModal.closeJimakuModal();
  }
  if (ctx.state.youtubePickerModalOpen) {
    youtubePickerModal.closeYoutubePickerModal();
  }
  if (ctx.state.playlistBrowserModalOpen) {
    playlistBrowserModal.closePlaylistBrowserModal();
  }
  if (ctx.state.runtimeOptionsModalOpen) {
    runtimeOptionsModal.closeRuntimeOptionsModal();
  }
  if (ctx.state.characterDictionaryModalOpen) {
    characterDictionaryModal.closeCharacterDictionaryModal();
  }
  if (ctx.state.subsyncModalOpen) {
    subsyncModal.closeSubsyncModal();
  }
  if (ctx.state.kikuModalOpen) {
    kikuModal.cancelKikuFieldGrouping();
  }
  if (ctx.state.sessionHelpModalOpen) {
    sessionHelpModal.closeSessionHelpModal();
  }

  syncSettingsModalSubtitleSuppression();
}

function applyControllerSnapshot(snapshot: {
  connectedGamepads: Array<{ id: string; index: number; mapping: string; connected: boolean }>;
  activeGamepadId: string | null;
  rawAxes: number[];
  rawButtons: Array<{ value: number; pressed: boolean; touched?: boolean }>;
}): void {
  controllerStatusIndicator.update({
    connectedGamepads: snapshot.connectedGamepads,
    activeGamepadId: snapshot.activeGamepadId,
  });
  ctx.state.connectedGamepads = snapshot.connectedGamepads;
  ctx.state.activeGamepadId = snapshot.activeGamepadId;
  ctx.state.controllerRawAxes = snapshot.rawAxes;
  ctx.state.controllerRawButtons = snapshot.rawButtons;
  controllerSelectModal.updateDevices();
  controllerDebugModal.updateSnapshot();
}

function showControllerDisabledNotice(): void {
  controllerStatusIndicator.show(
    'Controller support disabled. Set controller.enabled to true in config to use controller tools.',
  );
}

function emitControllerPopupScroll(deltaPixels: number): void {
  if (deltaPixels === 0) return;
  keyboardHandlers.scrollPopupByController(0, deltaPixels);
}

function emitControllerPopupJump(deltaPixels: number): void {
  if (deltaPixels === 0) return;
  keyboardHandlers.scrollPopupByController(0, deltaPixels * 4);
}

function startControllerPolling(): void {
  if (controllerAnimationFrameId !== null) {
    cancelAnimationFrame(controllerAnimationFrameId);
    controllerAnimationFrameId = null;
  }

  const gamepadController = createGamepadController({
    getGamepads: () => Array.from(navigator.getGamepads?.() ?? []),
    getConfig: () =>
      ctx.state.controllerConfig ?? {
        enabled: false,
        preferredGamepadId: '',
        preferredGamepadLabel: '',
        smoothScroll: true,
        scrollPixelsPerSecond: 900,
        horizontalJumpPixels: 160,
        stickDeadzone: 0.2,
        triggerInputMode: 'auto',
        triggerDeadzone: 0.5,
        repeatDelayMs: 320,
        repeatIntervalMs: 120,
        buttonIndices: {
          select: 6,
          buttonSouth: 0,
          buttonEast: 1,
          buttonWest: 2,
          buttonNorth: 3,
          leftShoulder: 4,
          rightShoulder: 5,
          leftStickPress: 9,
          rightStickPress: 10,
          leftTrigger: 6,
          rightTrigger: 7,
        },
        bindings: {
          toggleLookup: { kind: 'button', buttonIndex: 0 },
          closeLookup: { kind: 'button', buttonIndex: 1 },
          toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
          mineCard: { kind: 'button', buttonIndex: 2 },
          quitMpv: { kind: 'button', buttonIndex: 6 },
          previousAudio: { kind: 'none' },
          nextAudio: { kind: 'button', buttonIndex: 5 },
          playCurrentAudio: { kind: 'button', buttonIndex: 4 },
          toggleMpvPause: { kind: 'button', buttonIndex: 9 },
          leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
          leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
          rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
          rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
        },
        profiles: {},
      },
    getKeyboardModeEnabled: () => ctx.state.keyboardDrivenModeEnabled,
    getLookupWindowOpen: () => ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document),
    getInteractionBlocked: () => isControllerInputBlocked(),
    toggleKeyboardMode: () => keyboardHandlers.handleKeyboardModeToggleRequested(),
    toggleLookup: () => keyboardHandlers.handleLookupWindowToggleRequested(),
    closeLookup: () => {
      keyboardHandlers.closeLookupWindow();
    },
    moveSelection: (delta) => {
      keyboardHandlers.moveSelectionForController(delta);
    },
    mineCard: () => {
      keyboardHandlers.mineSelectedFromController();
    },
    quitMpv: () => {
      window.electronAPI.sendMpvCommand(['quit']);
    },
    previousAudio: () => {
      keyboardHandlers.cyclePopupAudioSourceForController(-1);
    },
    nextAudio: () => {
      keyboardHandlers.cyclePopupAudioSourceForController(1);
    },
    playCurrentAudio: () => {
      keyboardHandlers.playCurrentAudioForController();
    },
    toggleMpvPause: () => {
      window.electronAPI.sendMpvCommand(['cycle', 'pause']);
    },
    scrollPopup: (deltaPixels) => {
      emitControllerPopupScroll(deltaPixels);
    },
    jumpPopup: (deltaPixels) => {
      emitControllerPopupJump(deltaPixels);
    },
    onState: (snapshot) => {
      applyControllerSnapshot(snapshot);
    },
  });

  const poll = (now: number): void => {
    gamepadController.poll(now);
    controllerAnimationFrameId = requestAnimationFrame(poll);
  };

  controllerAnimationFrameId = requestAnimationFrame(poll);
}

function restoreOverlayInteractionAfterError(): void {
  ctx.state.isOverSubtitle = false;
  ctx.dom.overlay.classList.remove('interactive');
  if (ctx.platform.shouldToggleMouseIgnore) {
    window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
  }
}

function showOverlayErrorToast(message: string): void {
  if (overlayErrorToastTimeout) {
    clearTimeout(overlayErrorToastTimeout);
    overlayErrorToastTimeout = null;
  }
  ctx.dom.overlayErrorToast.textContent = message;
  ctx.dom.overlayErrorToast.classList.remove('hidden');
  overlayErrorToastTimeout = setTimeout(() => {
    ctx.dom.overlayErrorToast.classList.add('hidden');
    ctx.dom.overlayErrorToast.textContent = '';
    overlayErrorToastTimeout = null;
  }, 3200);
}

function showMiningImageToast(image: string): void {
  if (!image) {
    return;
  }
  if (miningImageToastTimeout) {
    clearTimeout(miningImageToastTimeout);
    miningImageToastTimeout = null;
  }
  ctx.dom.miningImageToastImage.src = image;
  ctx.dom.miningImageToast.classList.remove('hidden');
  miningImageToastTimeout = setTimeout(() => {
    ctx.dom.miningImageToast.classList.add('hidden');
    ctx.dom.miningImageToastImage.removeAttribute('src');
    miningImageToastTimeout = null;
  }, 3000);
}

const recovery = createRendererRecoveryController({
  dismissActiveUi: dismissActiveUiAfterError,
  restoreOverlayInteraction: restoreOverlayInteractionAfterError,
  showToast: showOverlayErrorToast,
  getSnapshot: () => ({
    activeModal: getActiveModal(),
    subtitlePreview: lastSubtitlePreview,
    secondarySubtitlePreview: lastSecondarySubtitlePreview,
    isOverlayInteractive: ctx.dom.overlay.classList.contains('interactive'),
    isOverSubtitle: ctx.state.isOverSubtitle,
    overlayLayer: ctx.platform.overlayLayer,
  }),
  logError: (payload) => {
    console.error('renderer overlay recovery', payload);
  },
});

registerRendererGlobalErrorHandlers(window, recovery);

function registerModalOpenHandlers(): void {
  window.electronAPI.onOpenRuntimeOptions(() => {
    runGuarded('runtime-options:open', () => {
      runtimeOptionsModal.openRuntimeOptionsModal();
      window.electronAPI.notifyOverlayModalOpened('runtime-options');
    });
  });
  window.electronAPI.onOpenCharacterDictionaryManager(() => {
    runGuardedAsync('character-dictionary-manager:open', async () => {
      await characterDictionaryModal.openCharacterDictionaryManagerModal();
    });
  });
  window.electronAPI.onOpenSessionHelp(() => {
    runGuarded('session-help:open', () => {
      sessionHelpModal.openSessionHelpModal(keyboardHandlers.getSessionHelpOpeningInfo());
      window.electronAPI.notifyOverlayModalOpened('session-help');
    });
  });
  window.electronAPI.onOpenControllerSelect(() => {
    runGuarded('controller-select:open', () => {
      if (controllerSelectModal.openControllerSelectModal()) {
        window.electronAPI.notifyOverlayModalOpened('controller-select');
      }
    });
  });
  window.electronAPI.onOpenControllerDebug(() => {
    runGuarded('controller-debug:open', () => {
      if (controllerDebugModal.openControllerDebugModal()) {
        window.electronAPI.notifyOverlayModalOpened('controller-debug');
      }
    });
  });
  window.electronAPI.onOpenJimaku(() => {
    runGuarded('jimaku:open', () => {
      jimakuModal.openJimakuModal();
      window.electronAPI.notifyOverlayModalOpened('jimaku');
    });
  });
  window.electronAPI.onOpenYoutubeTrackPicker((payload) => {
    runGuarded('youtube:picker-open', () => {
      youtubePickerModal.openYoutubePickerModal(payload);
    });
  });
  window.electronAPI.onOpenPlaylistBrowser(() => {
    runGuardedAsync('playlist-browser:open', async () => {
      await playlistBrowserModal.openPlaylistBrowserModal();
    });
  });
  window.electronAPI.onCancelYoutubeTrackPicker(() => {
    runGuarded('youtube:picker-cancel', () => {
      youtubePickerModal.closeYoutubePickerModal();
    });
  });
  window.electronAPI.onSubsyncManualOpen((payload: SubsyncManualPayload) => {
    runGuarded('subsync:manual-open', () => {
      subsyncModal.openSubsyncModal(payload);
      window.electronAPI.notifyOverlayModalOpened('subsync');
    });
  });
  window.electronAPI.onKikuFieldGroupingRequest(
    (data: { original: KikuDuplicateCardInfo; duplicate: KikuDuplicateCardInfo }) => {
      runGuarded('kiku:field-grouping-open', () => {
        prepareForKikuFieldGroupingOpen({
          closeLookupWindow: () => keyboardHandlers.closeLookupWindow(),
          pausePlayback: () => {
            window.electronAPI.sendMpvCommand(['set_property', 'pause', 'yes']);
          },
        });
        kikuModal.openKikuFieldGroupingModal(data);
        window.electronAPI.notifyOverlayModalOpened('kiku');
      });
    },
  );
}

function registerKeyboardCommandHandlers(): void {
  window.electronAPI.onSessionNumericSelectionStart((payload) => {
    runGuarded('session:numeric-selection-start', () => {
      keyboardHandlers.beginSessionNumericSelection(payload.actionId, payload.timeoutMs);
    });
  });

  window.electronAPI.onKeyboardModeToggleRequested(() => {
    runGuarded('keyboard-mode-toggle:requested', () => {
      keyboardHandlers.handleKeyboardModeToggleRequested();
    });
  });

  window.electronAPI.onLookupWindowToggleRequested(() => {
    runGuarded('lookup-window-toggle:requested', () => {
      keyboardHandlers.handleLookupWindowToggleRequested();
    });
  });

  window.electronAPI.onSubtitleSidebarToggle(() => {
    runGuardedAsync('subtitle-sidebar:toggle', async () => {
      await subtitleSidebarModal.toggleSubtitleSidebarModal();
    });
  });

  window.electronAPI.onPrimarySubtitleBarToggle(() => {
    runGuarded('primary-subtitle-bar:toggle', () => {
      keyboardHandlers.togglePrimarySubtitleBarVisibility();
    });
  });
}

function runGuarded(action: string, fn: () => void): void {
  try {
    fn();
  } catch (error) {
    recovery.handleError(error, { source: 'callback', action });
  }
}

function runGuardedAsync(action: string, fn: () => Promise<void> | void): void {
  Promise.resolve()
    .then(fn)
    .catch((error) => {
      recovery.handleError(error, { source: 'callback', action });
    });
}

registerModalOpenHandlers();
registerKeyboardCommandHandlers();
registerYomitanLookupListener(window, () => {
  runGuarded('yomitan:lookup', () => {
    window.electronAPI.recordYomitanLookup(subtitleSidebarModal.getSubtitleSidebarMiningContext());
  });
});

async function init(): Promise<void> {
  document.body.classList.add(`layer-${ctx.platform.overlayLayer}`);
  if (ctx.platform.isMacOSPlatform) {
    document.body.classList.add('platform-macos');
  }
  if (ctx.platform.isWindowsPlatform) {
    document.body.classList.add('platform-windows');
  }
  if (ctx.platform.shouldToggleMouseIgnore) {
    syncOverlayMouseIgnoreState(ctx);
  }

  window.electronAPI.onOverlayPointerRecoveryRequested(() => {
    runGuarded('overlay:pointer-recovery', () => {
      if (!ctx.platform.isMacOSPlatform || !ctx.platform.shouldToggleMouseIgnore) {
        return;
      }
      if (isAnyModalOpen()) {
        return;
      }
      mouseHandlers.restorePointerInteractionState();
    });
  });

  await keyboardHandlers.setupMpvInputForwarding();

  const initialSubtitleStyle = await window.electronAPI.getSubtitleStyle();
  subtitleRenderer.applySubtitleStyle(initialSubtitleStyle);
  subtitleRenderer.updatePrimarySubMode(initialSubtitleStyle?.primaryDefaultMode ?? 'visible');
  positioning.applyStoredSubtitlePosition(
    await window.electronAPI.getSubtitlePosition(),
    'startup',
  );
  measurementReporter.schedule();

  window.electronAPI.onSubtitlePosition((position: SubtitlePosition | null) => {
    runGuarded('subtitle-position:update', () => {
      positioning.applyStoredSubtitlePosition(position, 'media-change');
      measurementReporter.schedule();
      if (ctx.state.playlistBrowserModalOpen) {
        runGuardedAsync('playlist-browser:refresh-on-media-change', async () => {
          await playlistBrowserModal.refreshSnapshot();
        });
      }
    });
  });

  let initialSubtitle: SubtitleData | string = '';
  try {
    initialSubtitle = await window.electronAPI.getCurrentSubtitle();
  } catch {
    initialSubtitle = await window.electronAPI.getCurrentSubtitleRaw();
  }
  lastSubtitlePreview = truncateForErrorLog(getSubtitleTextForPreview(initialSubtitle));
  keyboardHandlers.handleSubtitleContentUpdated();
  subtitleRenderer.renderSubtitle(initialSubtitle);
  positioning.applyYPercent(positioning.getCurrentYPercent());
  measurementReporter.emitNow();

  window.electronAPI.onSubtitle((data: SubtitleData) => {
    runGuarded('subtitle:update', () => {
      lastSubtitlePreview = truncateForErrorLog(getSubtitleTextForPreview(data));
      keyboardHandlers.handleSubtitleContentUpdated();
      subtitleRenderer.renderSubtitle(data);
      positioning.applyYPercent(positioning.getCurrentYPercent());
      measurementReporter.emitNow();
      subtitleSidebarModal.handleSubtitleUpdated(data);
      measurementReporter.schedule();
    });
  });

  window.electronAPI.onSecondarySub((text: string) => {
    runGuarded('secondary-subtitle:update', () => {
      lastSecondarySubtitlePreview = truncateForErrorLog(text);
      subtitleRenderer.renderSecondarySub(text);
      measurementReporter.schedule();
    });
  });
  window.electronAPI.onSecondarySubMode((mode: SecondarySubMode) => {
    runGuarded('secondary-subtitle-mode:update', () => {
      subtitleRenderer.updateSecondarySubMode(mode);
      measurementReporter.schedule();
    });
  });

  subtitleRenderer.updateSecondarySubMode(await window.electronAPI.getSecondarySubMode());
  subtitleRenderer.renderSecondarySub(await window.electronAPI.getCurrentSecondarySub());
  measurementReporter.schedule();

  ctx.dom.subtitleContainer.addEventListener('mouseenter', mouseHandlers.handlePrimaryMouseEnter);
  ctx.dom.subtitleContainer.addEventListener('mouseleave', mouseHandlers.handlePrimaryMouseLeave);
  ctx.dom.secondarySubContainer.addEventListener(
    'mouseenter',
    mouseHandlers.handleSecondaryMouseEnter,
  );
  ctx.dom.secondarySubContainer.addEventListener(
    'mouseleave',
    mouseHandlers.handleSecondaryMouseLeave,
  );

  mouseHandlers.setupResizeHandler();
  mouseHandlers.setupPointerTracking();
  mouseHandlers.setupSelectionObserver();
  mouseHandlers.setupYomitanObserver();
  setupDragDropToMpvQueue();
  window.addEventListener('resize', () => {
    measurementReporter.schedule();
  });

  jimakuModal.wireDomEvents();
  youtubePickerModal.wireDomEvents();
  playlistBrowserModal.wireDomEvents();
  kikuModal.wireDomEvents();
  runtimeOptionsModal.wireDomEvents();
  subsyncModal.wireDomEvents();
  controllerSelectModal.wireDomEvents();
  controllerDebugModal.wireDomEvents();
  sessionHelpModal.wireDomEvents();
  subtitleSidebarModal.wireDomEvents();
  characterDictionaryModal.wireDomEvents();
  window.addEventListener('beforeunload', () => {
    subtitleSidebarModal.disposeDomEvents();
  });

  window.electronAPI.onRuntimeOptionsChanged((options: RuntimeOptionState[]) => {
    runGuarded('runtime-options:changed', () => {
      runtimeOptionsModal.updateRuntimeOptions(options);
    });
  });
  window.electronAPI.onConfigHotReload((payload: ConfigHotReloadPayload) => {
    runGuarded('config:hot-reload', () => {
      keyboardHandlers.updateSessionBindings(payload.sessionBindings);
      void keyboardHandlers.refreshConfiguredShortcuts();
      subtitleRenderer.applySubtitleStyle(payload.subtitleStyle);
      subtitleRenderer.updatePrimarySubMode(payload.primarySubMode);
      subtitleRenderer.updateSecondarySubMode(payload.secondarySubMode);
      ctx.state.subtitleSidebarConfig = payload.subtitleSidebar;
      ctx.state.subtitleSidebarToggleKey = payload.subtitleSidebar.toggleKey;
      ctx.state.subtitleSidebarPauseVideoOnHover = payload.subtitleSidebar.pauseVideoOnHover;
      ctx.state.subtitleSidebarAutoScroll = payload.subtitleSidebar.autoScroll;
      void subtitleSidebarModal.refreshSubtitleSidebarSnapshot();
      measurementReporter.schedule();
    });
  });
  window.electronAPI.onMiningImage((payload: MiningImagePayload) => {
    runGuarded('mining:image', () => {
      showMiningImageToast(payload.image);
    });
  });
  mouseHandlers.setupDragging();
  try {
    ctx.state.controllerConfig = await window.electronAPI.getControllerConfig();
  } catch (error) {
    console.error('Failed to load controller config.', error);
    ctx.state.controllerConfig = null;
  }
  startControllerPolling();

  await subtitleSidebarModal.refreshSubtitleSidebarSnapshot();
  await subtitleSidebarModal.autoOpenSubtitleSidebarOnStartup();

  measurementReporter.emitNow();
}

function setupDragDropToMpvQueue(): void {
  let dragDepth = 0;

  const setDropInteractive = (): void => {
    ctx.dom.overlay.classList.add('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }
  };

  const clearDropInteractive = (): void => {
    dragDepth = 0;
    if (isAnyModalOpen() || ctx.state.isOverSubtitle) {
      return;
    }
    ctx.dom.overlay.classList.remove('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
    }
  };

  document.addEventListener('dragenter', (event: DragEvent) => {
    if (!event.dataTransfer) return;
    dragDepth += 1;
    setDropInteractive();
  });

  document.addEventListener('dragover', (event: DragEvent) => {
    if (dragDepth <= 0 || !event.dataTransfer) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  });

  document.addEventListener('dragleave', () => {
    if (dragDepth <= 0) return;
    dragDepth -= 1;
    if (dragDepth === 0) {
      clearDropInteractive();
    }
  });

  document.addEventListener('drop', (event: DragEvent) => {
    if (!event.dataTransfer) return;
    event.preventDefault();

    const droppedVideoPaths = collectDroppedVideoPaths(event.dataTransfer);
    const droppedSubtitlePaths = collectDroppedSubtitlePaths(event.dataTransfer);
    const appendDroppedVideos = event.shiftKey;
    const loadCommands = buildMpvLoadfileCommands(droppedVideoPaths, appendDroppedVideos);
    const subtitleCommands = buildMpvSubtitleAddCommands(droppedSubtitlePaths);
    for (const command of loadCommands) {
      window.electronAPI.sendMpvCommand(command);
    }
    for (const command of subtitleCommands) {
      window.electronAPI.sendMpvCommand(command);
    }
    const osdParts: string[] = [];
    if (loadCommands.length > 0) {
      const action = appendDroppedVideos ? 'Queued' : 'Loaded';
      osdParts.push(`${action} ${loadCommands.length} file${loadCommands.length === 1 ? '' : 's'}`);
    }
    if (subtitleCommands.length > 0) {
      osdParts.push(
        `Loaded ${subtitleCommands.length} subtitle file${subtitleCommands.length === 1 ? '' : 's'}`,
      );
    }
    if (osdParts.length > 0) {
      window.electronAPI.sendMpvCommand(['show-text', osdParts.join(' | '), '1500']);
    }

    clearDropInteractive();
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    runGuardedAsync('bootstrap:init', init);
  });
} else {
  runGuardedAsync('bootstrap:init', init);
}
