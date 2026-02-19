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
  MpvSubtitleRenderMetrics,
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
import { resolveRendererDom } from './utils/dom.js';
import { resolvePlatformInfo } from './utils/platform.js';
import { buildMpvLoadfileCommands, collectDroppedVideoPaths } from '../core/services/overlay-drop.js';

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
const positioning = createPositioningController(ctx, {
  modalStateReader: { isAnySettingsModalOpen },
  applySubtitleFontSize: subtitleRenderer.applySubtitleFontSize,
});
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
  saveInvisiblePositionEdit: positioning.saveInvisiblePositionEdit,
  cancelInvisiblePositionEdit: positioning.cancelInvisiblePositionEdit,
  setInvisiblePositionEditMode: positioning.setInvisiblePositionEditMode,
  applyInvisibleSubtitleOffsetPosition: positioning.applyInvisibleSubtitleOffsetPosition,
  updateInvisiblePositionEditHud: positioning.updateInvisiblePositionEditHud,
  appendClipboardVideoToQueue: () => {
    void window.electronAPI.appendClipboardVideoToQueue();
  },
});
const mouseHandlers = createMouseHandlers(ctx, {
  modalStateReader: { isAnySettingsModalOpen, isAnyModalOpen },
  applyInvisibleSubtitleLayoutFromMpvMetrics:
    positioning.applyInvisibleSubtitleLayoutFromMpvMetrics,
  applyYPercent: positioning.applyYPercent,
  getCurrentYPercent: positioning.getCurrentYPercent,
  persistSubtitlePositionPatch: positioning.persistSubtitlePositionPatch,
});

async function init(): Promise<void> {
  document.body.classList.add(`layer-${ctx.platform.overlayLayer}`);

  window.electronAPI.onSubtitle((data: SubtitleData) => {
    subtitleRenderer.renderSubtitle(data);
    measurementReporter.schedule();
  });

  window.electronAPI.onSubtitlePosition((position: SubtitlePosition | null) => {
    if (ctx.platform.isInvisibleLayer) {
      positioning.applyInvisibleStoredSubtitlePosition(position, 'media-change');
    } else {
      positioning.applyStoredSubtitlePosition(position, 'media-change');
    }
    measurementReporter.schedule();
  });

  if (ctx.platform.isInvisibleLayer) {
    window.electronAPI.onMpvSubtitleRenderMetrics((metrics: MpvSubtitleRenderMetrics) => {
      positioning.applyInvisibleSubtitleLayoutFromMpvMetrics(metrics, 'event');
      measurementReporter.schedule();
    });
    window.electronAPI.onOverlayDebugVisualization((enabled: boolean) => {
      document.body.classList.toggle('debug-invisible-visualization', enabled);
    });
  }

  const initialSubtitle = await window.electronAPI.getCurrentSubtitleRaw();
  subtitleRenderer.renderSubtitle(initialSubtitle);
  measurementReporter.schedule();

  window.electronAPI.onSecondarySub((text: string) => {
    subtitleRenderer.renderSecondarySub(text);
    measurementReporter.schedule();
  });
  window.electronAPI.onSecondarySubMode((mode: SecondarySubMode) => {
    subtitleRenderer.updateSecondarySubMode(mode);
    measurementReporter.schedule();
  });

  subtitleRenderer.updateSecondarySubMode(await window.electronAPI.getSecondarySubMode());
  subtitleRenderer.renderSecondarySub(await window.electronAPI.getCurrentSecondarySub());
  measurementReporter.schedule();

  const hoverTarget = ctx.platform.isInvisibleLayer
    ? ctx.dom.subtitleRoot
    : ctx.dom.subtitleContainer;
  hoverTarget.addEventListener('mouseenter', mouseHandlers.handleMouseEnter);
  hoverTarget.addEventListener('mouseleave', mouseHandlers.handleMouseLeave);
  ctx.dom.secondarySubContainer.addEventListener('mouseenter', mouseHandlers.handleMouseEnter);
  ctx.dom.secondarySubContainer.addEventListener('mouseleave', mouseHandlers.handleMouseLeave);

  mouseHandlers.setupInvisibleHoverSelection();
  positioning.setupInvisiblePositionEditHud();
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
    runtimeOptionsModal.updateRuntimeOptions(options);
  });
  window.electronAPI.onConfigHotReload((payload: ConfigHotReloadPayload) => {
    keyboardHandlers.updateKeybindings(payload.keybindings);
    subtitleRenderer.applySubtitleStyle(payload.subtitleStyle);
    subtitleRenderer.updateSecondarySubMode(payload.secondarySubMode);
    measurementReporter.schedule();
  });
  window.electronAPI.onOpenRuntimeOptions(() => {
    runtimeOptionsModal.openRuntimeOptionsModal().catch(() => {
      runtimeOptionsModal.setRuntimeOptionsStatus('Failed to load runtime options', true);
      window.electronAPI.notifyOverlayModalClosed('runtime-options');
      syncSettingsModalSubtitleSuppression();
    });
  });
  window.electronAPI.onOpenJimaku(() => {
    jimakuModal.openJimakuModal();
  });
  window.electronAPI.onSubsyncManualOpen((payload: SubsyncManualPayload) => {
    subsyncModal.openSubsyncModal(payload);
  });
  window.electronAPI.onKikuFieldGroupingRequest(
    (data: { original: KikuDuplicateCardInfo; duplicate: KikuDuplicateCardInfo }) => {
      kikuModal.openKikuFieldGroupingModal(data);
    },
  );

  if (!ctx.platform.isInvisibleLayer) {
    mouseHandlers.setupDragging();
  }

  await keyboardHandlers.setupMpvInputForwarding();

  subtitleRenderer.applySubtitleStyle(await window.electronAPI.getSubtitleStyle());

  if (ctx.platform.isInvisibleLayer) {
    positioning.applyInvisibleStoredSubtitlePosition(
      await window.electronAPI.getSubtitlePosition(),
      'startup',
    );
    positioning.applyInvisibleSubtitleLayoutFromMpvMetrics(
      await window.electronAPI.getMpvSubtitleRenderMetrics(),
      'startup',
    );
  } else {
    positioning.applyStoredSubtitlePosition(
      await window.electronAPI.getSubtitlePosition(),
      'startup',
    );
    measurementReporter.schedule();
  }

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
    if (isAnyModalOpen() || ctx.state.isOverSubtitle || ctx.state.invisiblePositionEditMode) {
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
  document.addEventListener('DOMContentLoaded', init);
} else {
  void init();
}
