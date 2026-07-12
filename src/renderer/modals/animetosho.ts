import type {
  AnimetoshoApiResponse,
  AnimetoshoDownloadResult,
  AnimetoshoEntry,
  AnimetoshoSubtitleFile,
  JimakuMediaInfo,
} from '../../types';
import {
  animetoshoTrackMatchesLanguages,
  describeAnimetoshoTabLanguages,
  normalizeAnimetoshoLangCode,
} from '../../animetosho/lang.js';
import type { ModalStateReader, RendererContext } from '../context';

export function createAnimetoshoModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function setAnimetoshoStatus(message: string, isError = false): void {
    ctx.dom.animetoshoStatus.textContent = message;
    ctx.dom.animetoshoStatus.style.color = isError
      ? 'rgba(255, 120, 120, 0.95)'
      : 'rgba(255, 255, 255, 0.8)';
  }

  function resetAnimetoshoLists(): void {
    ctx.state.animetoshoEntries = [];
    ctx.state.animetoshoFiles = [];
    ctx.state.selectedAnimetoshoEntryIndex = 0;
    ctx.state.selectedAnimetoshoFileIndex = 0;
    ctx.state.currentAnimetoshoEntryId = null;

    ctx.dom.animetoshoEntriesList.innerHTML = '';
    ctx.dom.animetoshoFilesList.innerHTML = '';
    ctx.dom.animetoshoEntriesSection.classList.add('hidden');
    ctx.dom.animetoshoFilesSection.classList.add('hidden');
  }

  // Defaults to English until the configured secondarySub languages arrive.
  let secondaryLanguages: string[] = ['en'];

  function secondaryTabLabel(): string {
    return describeAnimetoshoTabLanguages(secondaryLanguages);
  }

  function isJapaneseTrack(file: AnimetoshoSubtitleFile): boolean {
    return normalizeAnimetoshoLangCode(file.lang) === 'ja';
  }

  function getVisibleFiles(): AnimetoshoSubtitleFile[] {
    if (ctx.state.animetoshoActiveTab === 'ja') {
      return ctx.state.animetoshoFiles.filter(isJapaneseTrack);
    }
    return ctx.state.animetoshoFiles.filter(
      (file) =>
        !isJapaneseTrack(file) && animetoshoTrackMatchesLanguages(file.lang, secondaryLanguages),
    );
  }

  function renderTabs(): void {
    if (ctx.state.animetoshoActiveTab === 'ja') {
      ctx.dom.animetoshoTabEnglishButton.classList.remove('active');
      ctx.dom.animetoshoTabJapaneseButton.classList.add('active');
    } else {
      ctx.dom.animetoshoTabEnglishButton.classList.add('active');
      ctx.dom.animetoshoTabJapaneseButton.classList.remove('active');
    }
  }

  function describeEmptyTab(): string {
    const hiddenCount = ctx.state.animetoshoFiles.length;
    if (ctx.state.animetoshoActiveTab === 'ja') {
      return hiddenCount > 0
        ? `No Japanese tracks in this release. Switch to the ${secondaryTabLabel()} tab.`
        : 'No Japanese tracks in this release.';
    }
    return hiddenCount > 0
      ? `No ${secondaryTabLabel()} tracks in this release. Switch to the Japanese tab.`
      : `No ${secondaryTabLabel()} tracks in this release.`;
  }

  function setActiveTab(tab: 'en' | 'ja'): void {
    if (ctx.state.animetoshoActiveTab === tab) return;
    ctx.state.animetoshoActiveTab = tab;
    ctx.state.selectedAnimetoshoFileIndex = 0;
    renderTabs();

    if (ctx.state.animetoshoFiles.length === 0) return;
    renderFiles();
    if (getVisibleFiles().length === 0) {
      setAnimetoshoStatus(describeEmptyTab());
    } else {
      setAnimetoshoStatus('Select a subtitle track.');
    }
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

  function renderEntries(): void {
    ctx.dom.animetoshoEntriesList.innerHTML = '';
    if (ctx.state.animetoshoEntries.length === 0) {
      ctx.dom.animetoshoEntriesSection.classList.add('hidden');
      return;
    }

    ctx.dom.animetoshoEntriesSection.classList.remove('hidden');
    ctx.state.animetoshoEntries.forEach((entry, index) => {
      const li = document.createElement('li');
      li.textContent = entry.title;

      const details: string[] = [];
      if (entry.totalSize !== null) details.push(formatBytes(entry.totalSize));
      if (entry.numFiles !== null) {
        details.push(`${entry.numFiles} file${entry.numFiles === 1 ? '' : 's'}`);
      }
      if (details.length > 0) {
        const sub = document.createElement('div');
        sub.className = 'jimaku-subtext';
        sub.textContent = details.join(' • ');
        li.appendChild(sub);
      }

      if (index === ctx.state.selectedAnimetoshoEntryIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        selectEntry(index);
      });

      ctx.dom.animetoshoEntriesList.appendChild(li);
    });
  }

  function renderFiles(): void {
    ctx.dom.animetoshoFilesList.innerHTML = '';
    const visibleFiles = getVisibleFiles();
    if (visibleFiles.length === 0) {
      ctx.dom.animetoshoFilesSection.classList.add('hidden');
      return;
    }

    ctx.dom.animetoshoFilesSection.classList.remove('hidden');
    visibleFiles.forEach((file, index) => {
      const li = document.createElement('li');
      li.textContent = file.filename;

      const details: string[] = [];
      if (file.lang) details.push(file.lang);
      if (file.trackName) details.push(file.trackName);
      details.push(formatBytes(file.size));
      const sub = document.createElement('div');
      sub.className = 'jimaku-subtext';
      sub.textContent = details.filter(Boolean).join(' • ');
      li.appendChild(sub);

      if (index === ctx.state.selectedAnimetoshoFileIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        void selectFile(index);
      });

      ctx.dom.animetoshoFilesList.appendChild(li);
    });
  }

  function getSearchQuery(): string {
    const title = ctx.dom.animetoshoTitleInput.value.trim();
    if (!title) return '';
    const episodeValue = ctx.dom.animetoshoEpisodeInput.value
      ? Number.parseInt(ctx.dom.animetoshoEpisodeInput.value, 10)
      : null;
    if (episodeValue !== null && Number.isFinite(episodeValue)) {
      return `${title} ${String(episodeValue).padStart(2, '0')}`;
    }
    return title;
  }

  async function performAnimetoshoSearch(): Promise<void> {
    const query = getSearchQuery();
    if (!query) {
      setAnimetoshoStatus('Enter a title before searching.', true);
      return;
    }

    resetAnimetoshoLists();
    setAnimetoshoStatus('Searching Animetosho...');

    const response: AnimetoshoApiResponse<AnimetoshoEntry[]> =
      await window.electronAPI.animetoshoSearchEntries({ query });
    if (!response.ok) {
      setAnimetoshoStatus(response.error.error, true);
      return;
    }

    ctx.state.animetoshoEntries = response.data;
    ctx.state.selectedAnimetoshoEntryIndex = 0;

    if (ctx.state.animetoshoEntries.length === 0) {
      setAnimetoshoStatus('No releases found.');
      return;
    }

    setAnimetoshoStatus('Select a release.');
    renderEntries();
    if (ctx.state.animetoshoEntries.length === 1) {
      selectEntry(0);
    }
  }

  async function loadFiles(entryId: number): Promise<void> {
    setAnimetoshoStatus('Loading subtitle tracks...');
    ctx.state.animetoshoFiles = [];
    ctx.state.selectedAnimetoshoFileIndex = 0;

    ctx.dom.animetoshoFilesList.innerHTML = '';
    ctx.dom.animetoshoFilesSection.classList.add('hidden');

    const response: AnimetoshoApiResponse<AnimetoshoSubtitleFile[]> =
      await window.electronAPI.animetoshoListFiles({ entryId });
    if (!response.ok) {
      setAnimetoshoStatus(response.error.error, true);
      return;
    }

    ctx.state.animetoshoFiles = response.data;
    if (ctx.state.animetoshoFiles.length === 0) {
      const entry = ctx.state.animetoshoEntries.find((candidate) => candidate.id === entryId);
      // The feed API omits per-file attachment data for multi-file torrents.
      if (entry && entry.numFiles !== null && entry.numFiles > 1) {
        setAnimetoshoStatus(
          'Batch releases are not supported. Pick a single-episode release instead.',
        );
      } else {
        setAnimetoshoStatus('No text subtitle tracks in this release. Try another one.');
      }
      return;
    }

    const visibleFiles = getVisibleFiles();
    if (visibleFiles.length === 0) {
      setAnimetoshoStatus(describeEmptyTab());
      return;
    }

    setAnimetoshoStatus('Select a subtitle track.');
    renderFiles();
    if (visibleFiles.length === 1) {
      await selectFile(0);
    }
  }

  function selectEntry(index: number): void {
    if (index < 0 || index >= ctx.state.animetoshoEntries.length) return;

    ctx.state.selectedAnimetoshoEntryIndex = index;
    ctx.state.currentAnimetoshoEntryId = ctx.state.animetoshoEntries[index]!.id;
    renderEntries();

    if (ctx.state.currentAnimetoshoEntryId !== null) {
      void loadFiles(ctx.state.currentAnimetoshoEntryId);
    }
  }

  async function selectFile(index: number): Promise<void> {
    const visibleFiles = getVisibleFiles();
    if (index < 0 || index >= visibleFiles.length) return;

    ctx.state.selectedAnimetoshoFileIndex = index;
    renderFiles();

    if (ctx.state.currentAnimetoshoEntryId === null) {
      setAnimetoshoStatus('Select a release first.', true);
      return;
    }

    const file = visibleFiles[index]!;
    setAnimetoshoStatus('Downloading subtitle...');

    const result: AnimetoshoDownloadResult = await window.electronAPI.animetoshoDownloadFile({
      entryId: ctx.state.currentAnimetoshoEntryId,
      url: file.url,
      name: file.filename,
      lang: file.lang,
    });

    if (result.ok) {
      setAnimetoshoStatus(`Downloaded and loaded: ${result.path}`);
      closeAnimetoshoModal();
      return;
    }

    setAnimetoshoStatus(result.error.error, true);
  }

  function isTextInputFocused(): boolean {
    const active = document.activeElement;
    if (!active) return false;
    const tag = active.tagName.toLowerCase();
    return tag === 'input' || tag === 'textarea';
  }

  async function loadSecondaryLanguages(): Promise<void> {
    try {
      const languages = await window.electronAPI.animetoshoGetSecondaryLanguages();
      secondaryLanguages = languages.length > 0 ? languages : ['en'];
    } catch {
      secondaryLanguages = ['en'];
    }
    ctx.dom.animetoshoTabEnglishButton.textContent = secondaryTabLabel();
    // Tracks may already be on screen if the languages arrived late.
    if (ctx.state.animetoshoFiles.length > 0) {
      renderFiles();
    }
  }

  function openAnimetoshoModal(): void {
    if (ctx.state.animetoshoModalOpen) return;

    ctx.state.animetoshoModalOpen = true;
    ctx.state.animetoshoActiveTab = 'en';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.animetoshoModal.classList.remove('hidden');
    ctx.dom.animetoshoModal.setAttribute('aria-hidden', 'false');

    setAnimetoshoStatus('Loading media info...');
    resetAnimetoshoLists();
    renderTabs();

    const secondaryLanguagesReady = loadSecondaryLanguages();

    window.electronAPI
      .getJimakuMediaInfo()
      .then(async (info: JimakuMediaInfo) => {
        ctx.dom.animetoshoTitleInput.value = info.title || '';
        ctx.dom.animetoshoEpisodeInput.value = info.episode ? String(info.episode) : '';

        if (info.confidence === 'high' && info.title && info.episode) {
          await secondaryLanguagesReady;
          void performAnimetoshoSearch();
        } else if (info.title) {
          setAnimetoshoStatus('Check title/episode and press Search.');
        } else {
          setAnimetoshoStatus('Enter title/episode and press Search.');
        }
      })
      .catch(() => {
        setAnimetoshoStatus('Failed to load media info.', true);
      });
  }

  function closeAnimetoshoModal(): void {
    if (!ctx.state.animetoshoModalOpen) return;

    ctx.state.animetoshoModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.animetoshoModal.classList.add('hidden');
    ctx.dom.animetoshoModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('animetosho');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    resetAnimetoshoLists();
  }

  function handleAnimetoshoKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeAnimetoshoModal();
      return true;
    }

    if (isTextInputFocused()) {
      if (e.key === 'Enter') {
        e.preventDefault();
        void performAnimetoshoSearch();
      }
      return true;
    }

    if (e.key === 'ArrowLeft') {
      e.preventDefault();
      setActiveTab('en');
      return true;
    }

    if (e.key === 'ArrowRight') {
      e.preventDefault();
      setActiveTab('ja');
      return true;
    }

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      const visibleFiles = getVisibleFiles();
      if (visibleFiles.length > 0) {
        ctx.state.selectedAnimetoshoFileIndex = Math.min(
          visibleFiles.length - 1,
          ctx.state.selectedAnimetoshoFileIndex + 1,
        );
        renderFiles();
      } else if (ctx.state.animetoshoEntries.length > 0) {
        ctx.state.selectedAnimetoshoEntryIndex = Math.min(
          ctx.state.animetoshoEntries.length - 1,
          ctx.state.selectedAnimetoshoEntryIndex + 1,
        );
        renderEntries();
      }
      return true;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (getVisibleFiles().length > 0) {
        ctx.state.selectedAnimetoshoFileIndex = Math.max(
          0,
          ctx.state.selectedAnimetoshoFileIndex - 1,
        );
        renderFiles();
      } else if (ctx.state.animetoshoEntries.length > 0) {
        ctx.state.selectedAnimetoshoEntryIndex = Math.max(
          0,
          ctx.state.selectedAnimetoshoEntryIndex - 1,
        );
        renderEntries();
      }
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (getVisibleFiles().length > 0) {
        void selectFile(ctx.state.selectedAnimetoshoFileIndex);
      } else if (ctx.state.animetoshoEntries.length > 0) {
        selectEntry(ctx.state.selectedAnimetoshoEntryIndex);
      } else {
        void performAnimetoshoSearch();
      }
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.animetoshoSearchButton.addEventListener('click', () => {
      void performAnimetoshoSearch();
    });
    ctx.dom.animetoshoCloseButton.addEventListener('click', () => {
      closeAnimetoshoModal();
    });
    ctx.dom.animetoshoTabEnglishButton.addEventListener('click', () => {
      setActiveTab('en');
    });
    ctx.dom.animetoshoTabJapaneseButton.addEventListener('click', () => {
      setActiveTab('ja');
    });
  }

  return {
    closeAnimetoshoModal,
    handleAnimetoshoKeydown,
    openAnimetoshoModal,
    wireDomEvents,
  };
}
