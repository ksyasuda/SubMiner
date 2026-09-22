import type { IpcMain, WebContents } from 'electron';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  parseSubtitleSelectionRequest,
  type SubtitleSelectionState,
} from '../../shared/subtitle-selection';
import { openOverlayHostedModal, retryOverlayModalOpen } from './overlay-hosted-modal-open';

interface SelectionMpvClient {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
  request: (command: unknown[]) => Promise<{ error?: string }>;
}

export function openSubtitleSelectionModal(
  deps: Parameters<typeof openOverlayHostedModal>[0] & Parameters<typeof retryOverlayModalOpen>[0],
): Promise<boolean> {
  return retryOverlayModalOpen(deps, {
    modal: 'subtitle-selection',
    timeoutMs: 1500,
    retryWarning: 'Subtitle selection modal did not acknowledge opening; retrying.',
    sendOpen: () =>
      openOverlayHostedModal(deps, {
        channel: IPC_CHANNELS.event.subtitleSelectionOpen,
        modal: 'subtitle-selection',
        preferModalWindow: true,
      }),
  });
}

export function createSubtitleSelectionRuntime(deps: {
  isEnabled: () => boolean;
  getMpvClient: () => SelectionMpvClient | null;
}) {
  function getClient(): SelectionMpvClient {
    if (!deps.isEnabled()) throw new Error('Enable subtitle selection in Settings first.');
    const client = deps.getMpvClient();
    if (!client?.connected) throw new Error('Connect to mpv first.');
    return client;
  }

  async function readState(client: SelectionMpvClient): Promise<SubtitleSelectionState> {
    const mediaPath = await client.requestProperty('path');
    if (typeof mediaPath !== 'string' || !mediaPath) throw new Error('Open a video first.');
    const [rawTracks, primary, secondary] = await Promise.all([
      client.requestProperty('track-list'),
      client.requestProperty('sid'),
      client.requestProperty('secondary-sid'),
    ]);
    const tracks: SubtitleSelectionState['tracks'] = [];
    const candidates: unknown[] = Array.isArray(rawTracks) ? rawTracks : [];
    for (const track of candidates) {
      if (
        typeof track !== 'object' ||
        track === null ||
        !('type' in track) ||
        track.type !== 'sub' ||
        !('id' in track) ||
        typeof track.id !== 'number' ||
        !Number.isSafeInteger(track.id) ||
        track.id <= 0
      )
        continue;
      const details = [
        'title' in track ? track.title : undefined,
        'lang' in track ? track.lang : undefined,
        'codec' in track ? track.codec : undefined,
      ].filter((value): value is string => typeof value === 'string' && value.length > 0);
      if ('external' in track && track.external === true) details.push('external');
      tracks.push({ id: track.id, label: `#${track.id} · ${details.join(' · ') || 'Subtitle'}` });
    }
    if ((await client.requestProperty('path')) !== mediaPath)
      throw new Error('The video changed. Reopen subtitle selection.');
    const selected = (value: unknown): number | null =>
      tracks.find((track) => track.id === value)?.id ?? null;
    return { mediaPath, tracks, primary: selected(primary), secondary: selected(secondary) };
  }

  async function apply(value: unknown): Promise<void> {
    const selection = parseSubtitleSelectionRequest(value);
    const client = getClient();
    const current = await readState(client);
    if (current.mediaPath !== selection.mediaPath)
      throw new Error('The video changed. Reopen subtitle selection.');
    for (const id of [selection.primary, selection.secondary]) {
      if (id !== null && !current.tracks.some((track) => track.id === id))
        throw new Error('A selected track is no longer available. Reopen subtitle selection.');
    }
    const set = async (property: string, id: number | null): Promise<void> => {
      const response = await client.request(['set_property', property, id ?? 'no']);
      if (response.error && response.error !== 'success') throw new Error(response.error);
    };
    // Clear secondary first so swapping the two tracks works in mpv.
    await set('secondary-sid', null);
    await set('sid', selection.primary);
    await set('secondary-sid', selection.secondary);
  }

  return { getState: async () => readState(getClient()), apply };
}

export function registerSubtitleSelectionIpc(deps: {
  ipc: Pick<IpcMain, 'handle'>;
  isAllowedSender: (sender: WebContents) => boolean;
  runtime: ReturnType<typeof createSubtitleSelectionRuntime>;
}): void {
  deps.ipc.handle(IPC_CHANNELS.request.getSubtitleSelection, (event) => {
    if (!deps.isAllowedSender(event.sender))
      throw new Error('Subtitle selection requires the overlay.');
    return deps.runtime.getState();
  });
  deps.ipc.handle(IPC_CHANNELS.request.applySubtitleSelection, (event, value: unknown) => {
    if (!deps.isAllowedSender(event.sender))
      throw new Error('Subtitle selection requires the overlay.');
    return deps.runtime.apply(value);
  });
}
