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
} from '../types';
import { createKeyboardHandlers } from './handlers/keyboard.js';
import { createMouseHandlers } from './handlers/mouse.js';
import { createJimakuModal } from './modals/jimaku.js';
import { createKikuModal } from './modals/kiku.js';
import { createSessionHelpModal } from './modals/session-help.js';
import { createRuntimeOptionsModal } from './modals/runtime-options.js';
import { createSubsyncModal } from './modals/subsync.js';
import { createPositioningController } from './positioning.js';
import { createOverlayContentMeasurementReporter } from './overlay-content-measurement.js';
import { createRendererState } from './state.js';
import { createSubtitleRenderer } from './subtitle-render.js';
import {
  createRendererRecoveryController,
  registerRendererGlobalErrorHandlers,
} from './error-recovery.js';
import { resolveRendererDom } from './utils/dom.js';
import { resolvePlatformInfo } from './utils/platform.js';
import {
  buildMpvLoadfileCommands,
  collectDroppedVideoPaths,
} from '../core/services/overlay-drop.js';

const ctx = {
  dom: resolveRendererDom(),
  platform: resolvePlatformInfo(),
  state: createRendererState(),
};

function isAnySettingsModalOpen(): boolean {
  return (
    ctx.state.runtimeOptionsModalOpen ||
    ctx.state.subsyncModalOpen ||
    ctx.state.kikuModalOpen ||
    ctx.state.jimakuModalOpen ||
    ctx.state.sessionHelpModalOpen
  );
}

function isAnyModalOpen(): boolean {
  return (
    ctx.state.jimakuModalOpen ||
    ctx.state.kikuModalOpen ||
    ctx.state.runtimeOptionsModalOpen ||
    ctx.state.subsyncModalOpen ||
    ctx.state.sessionHelpModalOpen
  );
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
const subsyncModal = createSubsyncModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const sessionHelpModal = createSessionHelpModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const kikuModal = createKikuModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const jimakuModal = createJimakuModal(ctx, {
  modalStateReader: { isAnyModalOpen },
  syncSettingsModalSubtitleSuppression,
});
const keyboardHandlers = createKeyboardHandlers(ctx, {
  handleRuntimeOptionsKeydown: runtimeOptionsModal.handleRuntimeOptionsKeydown,
  handleSubsyncKeydown: subsyncModal.handleSubsyncKeydown,
  handleKikuKeydown: kikuModal.handleKikuKeydown,
  handleJimakuKeydown: jimakuModal.handleJimakuKeydown,
  handleSessionHelpKeydown: sessionHelpModal.handleSessionHelpKeydown,
  openSessionHelpModal: sessionHelpModal.openSessionHelpModal,
  appendClipboardVideoToQueue: () => {
    void window.electronAPI.appendClipboardVideoToQueue();
  },
  getPlaybackPaused: () => window.electronAPI.getPlaybackPaused(),
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

let lastSubtitlePreview = '';
let lastSecondarySubtitlePreview = '';
let overlayErrorToastTimeout: ReturnType<typeof setTimeout> | null = null;

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
  if (ctx.state.jimakuModalOpen) return 'jimaku';
  if (ctx.state.kikuModalOpen) return 'kiku';
  if (ctx.state.runtimeOptionsModalOpen) return 'runtime-options';
  if (ctx.state.subsyncModalOpen) return 'subsync';
  if (ctx.state.sessionHelpModalOpen) return 'session-help';
  return null;
}

function dismissActiveUiAfterError(): void {
  if (ctx.state.jimakuModalOpen) {
    jimakuModal.closeJimakuModal();
  }
  if (ctx.state.runtimeOptionsModalOpen) {
    runtimeOptionsModal.closeRuntimeOptionsModal();
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
    runGuardedAsync('runtime-options:open', async () => {
      try {
        await runtimeOptionsModal.openRuntimeOptionsModal();
        window.electronAPI.notifyOverlayModalOpened('runtime-options');
      } catch {
        runtimeOptionsModal.setRuntimeOptionsStatus('Failed to load runtime options', true);
        window.electronAPI.notifyOverlayModalClosed('runtime-options');
        syncSettingsModalSubtitleSuppression();
      }
    });
  });
  window.electronAPI.onOpenJimaku(() => {
    runGuarded('jimaku:open', () => {
      jimakuModal.openJimakuModal();
      window.electronAPI.notifyOverlayModalOpened('jimaku');
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
        kikuModal.openKikuFieldGroupingModal(data);
        window.electronAPI.notifyOverlayModalOpened('kiku');
      });
    },
  );
}

function registerKeyboardCommandHandlers(): void {
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

async function init(): Promise<void> {
  document.body.classList.add(`layer-${ctx.platform.overlayLayer}`);
  if (ctx.platform.isMacOSPlatform) {
    document.body.classList.add('platform-macos');
  }

  window.electronAPI.onSubtitle((data: SubtitleData) => {
    runGuarded('subtitle:update', () => {
      lastSubtitlePreview = truncateForErrorLog(getSubtitleTextForPreview(data));
      subtitleRenderer.renderSubtitle(data);
      measurementReporter.schedule();
    });
  });

  window.electronAPI.onSubtitlePosition((position: SubtitlePosition | null) => {
    runGuarded('subtitle-position:update', () => {
      positioning.applyStoredSubtitlePosition(position, 'media-change');
      measurementReporter.schedule();
    });
  });

  let initialSubtitle: SubtitleData | string = '';
  try {
    initialSubtitle = await window.electronAPI.getCurrentSubtitle();
  } catch {
    initialSubtitle = await window.electronAPI.getCurrentSubtitleRaw();
  }
  lastSubtitlePreview = truncateForErrorLog(getSubtitleTextForPreview(initialSubtitle));
  subtitleRenderer.renderSubtitle(initialSubtitle);
  measurementReporter.schedule();

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

  ctx.dom.subtitleContainer.addEventListener('mouseenter', mouseHandlers.handleMouseEnter);
  ctx.dom.subtitleContainer.addEventListener('mouseleave', mouseHandlers.handleMouseLeave);
  ctx.dom.secondarySubContainer.addEventListener('mouseenter', mouseHandlers.handleMouseEnter);
  ctx.dom.secondarySubContainer.addEventListener('mouseleave', mouseHandlers.handleMouseLeave);

  mouseHandlers.setupResizeHandler();
  mouseHandlers.setupSelectionObserver();
  mouseHandlers.setupYomitanObserver();
  setupDragDropToMpvQueue();
  window.addEventListener('resize', () => {
    measurementReporter.schedule();
  });

  jimakuModal.wireDomEvents();
  kikuModal.wireDomEvents();
  runtimeOptionsModal.wireDomEvents();
  subsyncModal.wireDomEvents();
  sessionHelpModal.wireDomEvents();

  window.electronAPI.onRuntimeOptionsChanged((options: RuntimeOptionState[]) => {
    runGuarded('runtime-options:changed', () => {
      runtimeOptionsModal.updateRuntimeOptions(options);
    });
  });
  window.electronAPI.onConfigHotReload((payload: ConfigHotReloadPayload) => {
    runGuarded('config:hot-reload', () => {
      keyboardHandlers.updateKeybindings(payload.keybindings);
      subtitleRenderer.applySubtitleStyle(payload.subtitleStyle);
      subtitleRenderer.updateSecondarySubMode(payload.secondarySubMode);
      measurementReporter.schedule();
    });
  });
  mouseHandlers.setupDragging();

  await keyboardHandlers.setupMpvInputForwarding();

  subtitleRenderer.applySubtitleStyle(await window.electronAPI.getSubtitleStyle());

  positioning.applyStoredSubtitlePosition(
    await window.electronAPI.getSubtitlePosition(),
    'startup',
  );
  measurementReporter.schedule();

  if (ctx.platform.shouldToggleMouseIgnore) {
    window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
  }

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

    const droppedPaths = collectDroppedVideoPaths(event.dataTransfer);
    const loadCommands = buildMpvLoadfileCommands(droppedPaths, event.shiftKey);
    for (const command of loadCommands) {
      window.electronAPI.sendMpvCommand(command);
    }
    if (loadCommands.length > 0) {
      const action = event.shiftKey ? 'Queued' : 'Loaded';
      window.electronAPI.sendMpvCommand([
        'show-text',
        `${action} ${loadCommands.length} file${loadCommands.length === 1 ? '' : 's'}`,
        '1500',
      ]);
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
