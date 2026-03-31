import type {
  PlaylistBrowserMutationResult,
  PlaylistBrowserQueueItem,
  PlaylistBrowserSnapshot,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';
import {
  renderPlaylistBrowserDirectoryRow,
  renderPlaylistBrowserPlaylistRow,
} from './playlist-browser-renderer.js';

function clampIndex(index: number, length: number): number {
  if (length <= 0) return 0;
  return Math.min(Math.max(index, 0), length - 1);
}

function buildDefaultStatus(snapshot: PlaylistBrowserSnapshot): string {
  const directoryCount = snapshot.directoryItems.length;
  const playlistCount = snapshot.playlistItems.length;
  if (!snapshot.directoryAvailable) {
    return `${snapshot.directoryStatus} ${playlistCount > 0 ? `· ${playlistCount} queued` : ''}`.trim();
  }
  return `${directoryCount} sibling videos · ${playlistCount} queued`;
}

export function createPlaylistBrowserModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function setStatus(message: string, isError = false): void {
    ctx.state.playlistBrowserStatus = message;
    ctx.dom.playlistBrowserStatus.textContent = message;
    ctx.dom.playlistBrowserStatus.classList.toggle('error', isError);
  }

  function getSnapshot(): PlaylistBrowserSnapshot | null {
    return ctx.state.playlistBrowserSnapshot;
  }

  function resetSnapshotUi(): void {
    ctx.state.playlistBrowserSnapshot = null;
    ctx.state.playlistBrowserStatus = '';
    ctx.state.playlistBrowserSelectedDirectoryIndex = 0;
    ctx.state.playlistBrowserSelectedPlaylistIndex = 0;
    ctx.dom.playlistBrowserTitle.textContent = 'Playlist Browser';
    ctx.dom.playlistBrowserDirectoryList.replaceChildren();
    ctx.dom.playlistBrowserPlaylistList.replaceChildren();
    ctx.dom.playlistBrowserStatus.textContent = '';
    ctx.dom.playlistBrowserStatus.classList.remove('error');
  }

  function syncSelection(snapshot: PlaylistBrowserSnapshot): void {
    const directoryIndex = snapshot.directoryItems.findIndex((item) => item.isCurrentFile);
    const playlistIndex =
      snapshot.playingIndex ?? snapshot.playlistItems.findIndex((item) => item.current || item.playing);
    ctx.state.playlistBrowserSelectedDirectoryIndex = clampIndex(
      directoryIndex >= 0 ? directoryIndex : 0,
      snapshot.directoryItems.length,
    );
    ctx.state.playlistBrowserSelectedPlaylistIndex = clampIndex(
      playlistIndex >= 0 ? playlistIndex : 0,
      snapshot.playlistItems.length,
    );
  }

  function render(): void {
    const snapshot = getSnapshot();
    if (!snapshot) {
      ctx.dom.playlistBrowserDirectoryList.replaceChildren();
      ctx.dom.playlistBrowserPlaylistList.replaceChildren();
      return;
    }

    ctx.dom.playlistBrowserTitle.textContent = snapshot.directoryPath ?? 'Playlist Browser';
    ctx.dom.playlistBrowserStatus.textContent =
      ctx.state.playlistBrowserStatus || buildDefaultStatus(snapshot);
    ctx.dom.playlistBrowserDirectoryList.replaceChildren(
      ...snapshot.directoryItems.map((item, index) =>
        renderPlaylistBrowserDirectoryRow(ctx, item, index, {
          appendDirectoryItem,
          movePlaylistItem,
          playPlaylistItem,
          removePlaylistItem,
          render,
        }),
      ),
    );
    ctx.dom.playlistBrowserPlaylistList.replaceChildren(
      ...snapshot.playlistItems.map((item, index) =>
        renderPlaylistBrowserPlaylistRow(ctx, item, index, {
          appendDirectoryItem,
          movePlaylistItem,
          playPlaylistItem,
          removePlaylistItem,
          render,
        }),
      ),
    );
  }

  function applySnapshot(snapshot: PlaylistBrowserSnapshot): void {
    ctx.state.playlistBrowserSnapshot = snapshot;
    syncSelection(snapshot);
    render();
  }

  async function refreshSnapshot(): Promise<void> {
    try {
      const snapshot = await window.electronAPI.getPlaylistBrowserSnapshot();
      ctx.state.playlistBrowserStatus = '';
      applySnapshot(snapshot);
      setStatus(
        buildDefaultStatus(snapshot),
        !snapshot.directoryAvailable && snapshot.directoryStatus.length > 0,
      );
    } catch (error) {
      resetSnapshotUi();
      setStatus(error instanceof Error ? error.message : String(error), true);
    }
  }

  async function handleMutation(
    action: Promise<PlaylistBrowserMutationResult>,
    fallbackMessage: string,
  ): Promise<void> {
    const result = await action;
    if (!result.ok) {
      setStatus(result.message, true);
      return;
    }
    setStatus(result.message || fallbackMessage, false);
    if (result.snapshot) {
      applySnapshot(result.snapshot);
      return;
    }
    await refreshSnapshot();
  }

  async function appendDirectoryItem(filePath: string): Promise<void> {
    await handleMutation(window.electronAPI.appendPlaylistBrowserFile(filePath), 'Queued file');
  }

  async function playPlaylistItem(index: number): Promise<void> {
    const result = await window.electronAPI.playPlaylistBrowserIndex(index);
    if (!result.ok) {
      setStatus(result.message, true);
      return;
    }
    closePlaylistBrowserModal();
  }

  async function removePlaylistItem(index: number): Promise<void> {
    await handleMutation(window.electronAPI.removePlaylistBrowserIndex(index), 'Removed queue item');
  }

  async function movePlaylistItem(index: number, direction: 1 | -1): Promise<void> {
    await handleMutation(
      window.electronAPI.movePlaylistBrowserIndex(index, direction),
      'Moved queue item',
    );
  }

  async function openPlaylistBrowserModal(): Promise<void> {
    if (ctx.state.playlistBrowserModalOpen) {
      await refreshSnapshot();
      return;
    }
    if (options.modalStateReader.isAnyModalOpen()) {
      return;
    }

    ctx.state.playlistBrowserModalOpen = true;
    ctx.state.playlistBrowserActivePane = 'playlist';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.playlistBrowserModal.classList.remove('hidden');
    ctx.dom.playlistBrowserModal.setAttribute('aria-hidden', 'false');
    window.electronAPI.notifyOverlayModalOpened('playlist-browser');
    await refreshSnapshot();
  }

  function closePlaylistBrowserModal(): void {
    if (!ctx.state.playlistBrowserModalOpen) return;
    ctx.state.playlistBrowserModalOpen = false;
    resetSnapshotUi();
    ctx.dom.playlistBrowserModal.classList.add('hidden');
    ctx.dom.playlistBrowserModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('playlist-browser');
    options.syncSettingsModalSubtitleSuppression();
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function moveSelection(delta: number): void {
    const snapshot = getSnapshot();
    if (!snapshot) return;
    if (ctx.state.playlistBrowserActivePane === 'directory') {
      ctx.state.playlistBrowserSelectedDirectoryIndex = clampIndex(
        ctx.state.playlistBrowserSelectedDirectoryIndex + delta,
        snapshot.directoryItems.length,
      );
    } else {
      ctx.state.playlistBrowserSelectedPlaylistIndex = clampIndex(
        ctx.state.playlistBrowserSelectedPlaylistIndex + delta,
        snapshot.playlistItems.length,
      );
    }
    render();
  }

  function jumpSelection(target: 'start' | 'end'): void {
    const snapshot = getSnapshot();
    if (!snapshot) return;
    const length =
      ctx.state.playlistBrowserActivePane === 'directory'
        ? snapshot.directoryItems.length
        : snapshot.playlistItems.length;
    const nextIndex = target === 'start' ? 0 : Math.max(0, length - 1);
    if (ctx.state.playlistBrowserActivePane === 'directory') {
      ctx.state.playlistBrowserSelectedDirectoryIndex = nextIndex;
    } else {
      ctx.state.playlistBrowserSelectedPlaylistIndex = nextIndex;
    }
    render();
  }

  function activateSelection(): void {
    const snapshot = getSnapshot();
    if (!snapshot) return;
    if (ctx.state.playlistBrowserActivePane === 'directory') {
      const item = snapshot.directoryItems[ctx.state.playlistBrowserSelectedDirectoryIndex];
      if (item) {
        void appendDirectoryItem(item.path);
      }
      return;
    }

    const item = snapshot.playlistItems[ctx.state.playlistBrowserSelectedPlaylistIndex];
    if (item) {
      void playPlaylistItem(item.index);
    }
  }

  function handlePlaylistBrowserKeydown(event: KeyboardEvent): boolean {
    if (!ctx.state.playlistBrowserModalOpen) return false;

    if (event.key === 'Escape') {
      event.preventDefault();
      closePlaylistBrowserModal();
      return true;
    }
    if (event.key === 'Tab') {
      event.preventDefault();
      ctx.state.playlistBrowserActivePane =
        ctx.state.playlistBrowserActivePane === 'directory' ? 'playlist' : 'directory';
      render();
      return true;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      jumpSelection('start');
      return true;
    }
    if (event.key === 'End') {
      event.preventDefault();
      jumpSelection('end');
      return true;
    }
    if (event.key === 'ArrowUp' && (event.ctrlKey || event.metaKey)) {
      if (ctx.state.playlistBrowserActivePane === 'playlist') {
        event.preventDefault();
        const item = getSnapshot()?.playlistItems[ctx.state.playlistBrowserSelectedPlaylistIndex];
        if (item) {
          void movePlaylistItem(item.index, -1);
        }
        return true;
      }
    }
    if (event.key === 'ArrowDown' && (event.ctrlKey || event.metaKey)) {
      if (ctx.state.playlistBrowserActivePane === 'playlist') {
        event.preventDefault();
        const item = getSnapshot()?.playlistItems[ctx.state.playlistBrowserSelectedPlaylistIndex];
        if (item) {
          void movePlaylistItem(item.index, 1);
        }
        return true;
      }
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      moveSelection(-1);
      return true;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      moveSelection(1);
      return true;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      activateSelection();
      return true;
    }
    if (event.key === 'Delete' || event.key === 'Backspace') {
      if (ctx.state.playlistBrowserActivePane === 'playlist') {
        event.preventDefault();
        const item = getSnapshot()?.playlistItems[ctx.state.playlistBrowserSelectedPlaylistIndex];
        if (item) {
          void removePlaylistItem(item.index);
        }
        return true;
      }
    }

    return false;
  }

  function wireDomEvents(): void {
    ctx.dom.playlistBrowserClose.addEventListener('click', () => {
      closePlaylistBrowserModal();
    });
  }

  return {
    openPlaylistBrowserModal,
    closePlaylistBrowserModal,
    handlePlaylistBrowserKeydown,
    refreshSnapshot,
    wireDomEvents,
  };
}
