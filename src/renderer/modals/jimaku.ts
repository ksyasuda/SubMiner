import type {
  JimakuApiResponse,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuMediaInfo,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';
import { i18n } from '../../i18n/index.js';

export function createJimakuModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function setJimakuStatus(message: string, isError = false): void {
    ctx.dom.jimakuStatus.textContent = message;
    ctx.dom.jimakuStatus.style.color = isError
      ? 'rgba(255, 120, 120, 0.95)'
      : 'rgba(255, 255, 255, 0.8)';
  }

  function resetJimakuLists(): void {
    ctx.state.jimakuEntries = [];
    ctx.state.jimakuFiles = [];
    ctx.state.selectedEntryIndex = 0;
    ctx.state.selectedFileIndex = 0;
    ctx.state.currentEntryId = null;

    ctx.dom.jimakuEntriesList.innerHTML = '';
    ctx.dom.jimakuFilesList.innerHTML = '';
    ctx.dom.jimakuEntriesSection.classList.add('hidden');
    ctx.dom.jimakuFilesSection.classList.add('hidden');
    ctx.dom.jimakuBroadenButton.classList.add('hidden');
  }

  function formatEntryLabel(entry: JimakuEntry): string {
    if (entry.english_name && entry.english_name !== entry.name) {
      return `${entry.name} / ${entry.english_name}`;
    }
    return entry.name;
  }

  function renderEntries(): void {
    ctx.dom.jimakuEntriesList.innerHTML = '';
    if (ctx.state.jimakuEntries.length === 0) {
      ctx.dom.jimakuEntriesSection.classList.add('hidden');
      return;
    }

    ctx.dom.jimakuEntriesSection.classList.remove('hidden');
    ctx.state.jimakuEntries.forEach((entry, index) => {
      const li = document.createElement('li');
      li.textContent = formatEntryLabel(entry);

      if (entry.japanese_name) {
        const sub = document.createElement('div');
        sub.className = 'jimaku-subtext';
        sub.textContent = entry.japanese_name;
        li.appendChild(sub);
      }

      if (index === ctx.state.selectedEntryIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        selectEntry(index);
      });

      ctx.dom.jimakuEntriesList.appendChild(li);
    });
  }

  function formatBytes(size: number): string {
    if (!Number.isFinite(size)) return '';
    const units = ['B', 'KB', 'MB', 'GB'];
    let value = size;
    let idx = 0;
    while (value >= 1024 && idx < units.length - 1) {
      value /= 1024;
      idx += 1;
    }
    return `${value.toFixed(value >= 10 || idx === 0 ? 0 : 1)} ${units[idx]}`;
  }

  function renderFiles(): void {
    ctx.dom.jimakuFilesList.innerHTML = '';
    if (ctx.state.jimakuFiles.length === 0) {
      ctx.dom.jimakuFilesSection.classList.add('hidden');
      return;
    }

    ctx.dom.jimakuFilesSection.classList.remove('hidden');
    ctx.state.jimakuFiles.forEach((file, index) => {
      const li = document.createElement('li');
      li.textContent = file.name;

      const sub = document.createElement('div');
      sub.className = 'jimaku-subtext';
      sub.textContent = `${formatBytes(file.size)} • ${file.last_modified}`;
      li.appendChild(sub);

      if (index === ctx.state.selectedFileIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        void selectFile(index);
      });

      ctx.dom.jimakuFilesList.appendChild(li);
    });
  }

  function getSearchQuery(): { query: string; episode: number | null } {
    const title = ctx.dom.jimakuTitleInput.value.trim();
    const episode = ctx.dom.jimakuEpisodeInput.value
      ? Number.parseInt(ctx.dom.jimakuEpisodeInput.value, 10)
      : null;
    return { query: title, episode: Number.isFinite(episode) ? episode : null };
  }

  async function performJimakuSearch(): Promise<void> {
    const { query, episode } = getSearchQuery();
    if (!query) {
      setJimakuStatus(i18n.t('jimaku.enterTitle'), true);
      return;
    }

    resetJimakuLists();
    setJimakuStatus(i18n.t('jimaku.searching'));
    ctx.state.currentEpisodeFilter = episode;

    const response: JimakuApiResponse<JimakuEntry[]> = await window.electronAPI.jimakuSearchEntries(
      { query },
    );
    if (!response.ok) {
      const retry = response.error.retryAfter
        ? ` Retry after ${response.error.retryAfter.toFixed(1)}s.`
        : '';
      setJimakuStatus(`${response.error.error}${retry}`, true);
      return;
    }

    ctx.state.jimakuEntries = response.data;
    ctx.state.selectedEntryIndex = 0;

    if (ctx.state.jimakuEntries.length === 0) {
      setJimakuStatus(i18n.t('jimaku.noEntries'));
      return;
    }

    setJimakuStatus(i18n.t('jimaku.selectEntry'));
    renderEntries();
    if (ctx.state.jimakuEntries.length === 1) {
      void selectEntry(0);
    }
  }

  async function loadFiles(entryId: number, episode: number | null): Promise<void> {
    setJimakuStatus(i18n.t('jimaku.loadingFiles'));
    ctx.state.jimakuFiles = [];
    ctx.state.selectedFileIndex = 0;

    ctx.dom.jimakuFilesList.innerHTML = '';
    ctx.dom.jimakuFilesSection.classList.add('hidden');

    const response: JimakuApiResponse<JimakuFileEntry[]> = await window.electronAPI.jimakuListFiles(
      {
        entryId,
        episode,
      },
    );
    if (!response.ok) {
      const retry = response.error.retryAfter
        ? ` Retry after ${response.error.retryAfter.toFixed(1)}s.`
        : '';
      setJimakuStatus(`${response.error.error}${retry}`, true);
      return;
    }

    ctx.state.jimakuFiles = response.data;
    if (ctx.state.jimakuFiles.length === 0) {
      if (episode !== null) {
        setJimakuStatus(i18n.t('jimaku.noFiles'));
        ctx.dom.jimakuBroadenButton.classList.remove('hidden');
      } else {
        setJimakuStatus(i18n.t('jimaku.noFilesAll'));
      }
      return;
    }

    ctx.dom.jimakuBroadenButton.classList.add('hidden');
    setJimakuStatus(i18n.t('jimaku.selectFile'));
    renderFiles();
    if (ctx.state.jimakuFiles.length === 1) {
      await selectFile(0);
    }
  }

  function selectEntry(index: number): void {
    if (index < 0 || index >= ctx.state.jimakuEntries.length) return;

    ctx.state.selectedEntryIndex = index;
    ctx.state.currentEntryId = ctx.state.jimakuEntries[index]!.id;
    renderEntries();

    if (ctx.state.currentEntryId !== null) {
      void loadFiles(ctx.state.currentEntryId, ctx.state.currentEpisodeFilter);
    }
  }

  async function selectFile(index: number): Promise<void> {
    if (index < 0 || index >= ctx.state.jimakuFiles.length) return;

    ctx.state.selectedFileIndex = index;
    renderFiles();

    if (ctx.state.currentEntryId === null) {
      setJimakuStatus(i18n.t('jimaku.selectEntryFirst'), true);
      return;
    }

    const file = ctx.state.jimakuFiles[index]!;
    setJimakuStatus('Downloading subtitle...');

    const result: JimakuDownloadResult = await window.electronAPI.jimakuDownloadFile({
      entryId: ctx.state.currentEntryId,
      url: file.url,
      name: file.name,
    });

    if (result.ok) {
      setJimakuStatus(`Downloaded and loaded: ${result.path}`);
      closeJimakuModal();
      return;
    }

    const retry = result.error.retryAfter
      ? ` Retry after ${result.error.retryAfter.toFixed(1)}s.`
      : '';
    setJimakuStatus(`${result.error.error}${retry}`, true);
  }

  function isTextInputFocused(): boolean {
    const active = document.activeElement;
    if (!active) return false;
    const tag = active.tagName.toLowerCase();
    return tag === 'input' || tag === 'textarea';
  }

  function openJimakuModal(): void {
    if (ctx.state.jimakuModalOpen) return;

    ctx.state.jimakuModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.jimakuModal.classList.remove('hidden');
    ctx.dom.jimakuModal.setAttribute('aria-hidden', 'false');

    setJimakuStatus(i18n.t('jimaku.loadingMedia'));
    resetJimakuLists();

    window.electronAPI
      .getJimakuMediaInfo()
      .then((info: JimakuMediaInfo) => {
        ctx.dom.jimakuTitleInput.value = info.title || '';
        ctx.dom.jimakuSeasonInput.value = info.season ? String(info.season) : '';
        ctx.dom.jimakuEpisodeInput.value = info.episode ? String(info.episode) : '';
        ctx.state.currentEpisodeFilter = info.episode ?? null;

        if (info.confidence === 'high' && info.title && info.episode) {
          void performJimakuSearch();
        } else if (info.title) {
          setJimakuStatus(i18n.t('jimaku.readyHint'));
        } else {
          setJimakuStatus(i18n.t('jimaku.readyHint'));
        }
      })
      .catch(() => {
        setJimakuStatus(i18n.t('jimaku.failedMedia'), true);
      });
  }

  function closeJimakuModal(): void {
    if (!ctx.state.jimakuModalOpen) return;

    ctx.state.jimakuModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.jimakuModal.classList.add('hidden');
    ctx.dom.jimakuModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('jimaku');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    resetJimakuLists();
  }

  function handleJimakuKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeJimakuModal();
      return true;
    }

    if (isTextInputFocused()) {
      if (e.key === 'Enter') {
        e.preventDefault();
        void performJimakuSearch();
      }
      return true;
    }

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (ctx.state.jimakuFiles.length > 0) {
        ctx.state.selectedFileIndex = Math.min(
          ctx.state.jimakuFiles.length - 1,
          ctx.state.selectedFileIndex + 1,
        );
        renderFiles();
      } else if (ctx.state.jimakuEntries.length > 0) {
        ctx.state.selectedEntryIndex = Math.min(
          ctx.state.jimakuEntries.length - 1,
          ctx.state.selectedEntryIndex + 1,
        );
        renderEntries();
      }
      return true;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (ctx.state.jimakuFiles.length > 0) {
        ctx.state.selectedFileIndex = Math.max(0, ctx.state.selectedFileIndex - 1);
        renderFiles();
      } else if (ctx.state.jimakuEntries.length > 0) {
        ctx.state.selectedEntryIndex = Math.max(0, ctx.state.selectedEntryIndex - 1);
        renderEntries();
      }
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (ctx.state.jimakuFiles.length > 0) {
        void selectFile(ctx.state.selectedFileIndex);
      } else if (ctx.state.jimakuEntries.length > 0) {
        selectEntry(ctx.state.selectedEntryIndex);
      } else {
        void performJimakuSearch();
      }
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.jimakuSearchButton.addEventListener('click', () => {
      void performJimakuSearch();
    });
    ctx.dom.jimakuCloseButton.addEventListener('click', () => {
      closeJimakuModal();
    });
    ctx.dom.jimakuBroadenButton.addEventListener('click', () => {
      if (ctx.state.currentEntryId !== null) {
        ctx.dom.jimakuBroadenButton.classList.add('hidden');
        void loadFiles(ctx.state.currentEntryId, null);
      }
    });
  }

  return {
    closeJimakuModal,
    handleJimakuKeydown,
    openJimakuModal,
    wireDomEvents,
  };
}
