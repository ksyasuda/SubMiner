import type { PluginPanelContent, PluginPanelHandle } from '../../core/services/plugins/types';
import type { OverlayHostedModal } from '../../shared/ipc/contracts';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { PluginPanelSnapshot } from '../../types/runtime';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

const PLUGIN_PANEL_MODAL: OverlayHostedModal = 'plugin-panel';
const PLUGIN_PANEL_OPEN_TIMEOUT_MS = 1500;

interface PanelState {
  panelId: number;
  content: PluginPanelContent;
  controller: AbortController;
}

/** Sends the open request, retrying once like the other overlay-hosted modals. */
export async function openPluginPanelModal(
  deps: Parameters<typeof openOverlayHostedModal>[0] & Parameters<typeof retryOverlayModalOpen>[0],
  getSnapshot: () => PluginPanelSnapshot,
): Promise<boolean> {
  return await retryOverlayModalOpen(deps, {
    modal: PLUGIN_PANEL_MODAL,
    timeoutMs: PLUGIN_PANEL_OPEN_TIMEOUT_MS,
    retryWarning:
      'Plugin panel did not acknowledge modal open on first attempt; retrying dedicated modal window.',
    sendOpen: () =>
      openOverlayHostedModal(deps, {
        channel: IPC_CHANNELS.event.pluginPanelOpen,
        modal: PLUGIN_PANEL_MODAL,
        payload: getSnapshot(),
        preferModalWindow: true,
      }),
  });
}

/**
 * Owns the single plugin panel. Opening a panel replaces the previous one, and each handle's
 * signal aborts when its panel goes away so plugins can stop the work feeding it.
 */
export function createPluginPanelRuntime(deps: {
  openModal: (getSnapshot: () => PluginPanelSnapshot) => Promise<boolean>;
  /** Sends to every overlay window; renderers ignore panel ids they are not showing. */
  broadcast: (channel: string, payload: unknown) => void;
  onOpenFailed: () => void;
}) {
  let nextPanelId = 0;
  let active: PanelState | null = null;

  const snapshot = (state: PanelState): PluginPanelSnapshot => ({
    panelId: state.panelId,
    ...state.content,
  });

  const finish = (state: PanelState): void => {
    if (active === state) active = null;
    state.controller.abort();
  };

  return {
    open(content: PluginPanelContent): PluginPanelHandle {
      if (active) finish(active);
      const state: PanelState = {
        panelId: ++nextPanelId,
        content: { ...content },
        controller: new AbortController(),
      };
      active = state;

      const handleOpenFailure = (): void => {
        if (state.controller.signal.aborted) return;
        finish(state);
        deps.onOpenFailed();
      };
      deps
        .openModal(() => snapshot(state))
        .then((opened) => {
          if (!opened) {
            handleOpenFailure();
            return;
          }
          // Updates sent before the renderer acknowledged the open were dropped there.
          if (!state.controller.signal.aborted) {
            deps.broadcast(IPC_CHANNELS.event.pluginPanelUpdate, snapshot(state));
          }
        })
        .catch(handleOpenFailure);

      return {
        signal: state.controller.signal,
        get closed() {
          return state.controller.signal.aborted;
        },
        update: (next) => {
          if (state.controller.signal.aborted) return;
          state.content = { ...state.content, ...next };
          deps.broadcast(IPC_CHANNELS.event.pluginPanelUpdate, snapshot(state));
        },
        close: () => {
          if (state.controller.signal.aborted) return;
          finish(state);
          deps.broadcast(IPC_CHANNELS.event.pluginPanelClose, state.panelId);
        },
      };
    },

    /**
     * The overlay closed `panelId`. Returns false for a late notice about a panel that a newer one
     * already replaced, so the caller skips the modal teardown that would hide the newer panel.
     */
    handleModalClosed(panelId: number): boolean {
      if (!active) return true;
      if (active.panelId !== panelId) return false;
      finish(active);
      return true;
    },
  };
}
