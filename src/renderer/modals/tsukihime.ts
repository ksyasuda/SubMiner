import type {
  TsukihimeApiResponse,
  TsukihimeDownloadResult,
  TsukihimeEntry,
  TsukihimeSubtitleFile,
  JimakuMediaInfo,
} from '../../types';
import {
  tsukihimeTrackMatchesLanguages,
  describeTsukihimeTabLanguages,
  normalizeTsukihimeLangCode,
} from '../../tsukihime/lang.js';
import type { ModalStateReader, RendererContext } from '../context';

export function createTsukihimeModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  function setTsukihimeStatus(message: string, isError = false): void {
    ctx.dom.tsukihimeStatus.textContent = message;
    ctx.dom.tsukihimeStatus.style.color = isError
      ? 'rgba(255, 120, 120, 0.95)'
      : 'rgba(255, 255, 255, 0.8)';
  }

  function resetTsukihimeLists(): void {
    ctx.state.tsukihimeEntries = [];
    ctx.state.tsukihimeFiles = [];
    ctx.state.selectedTsukihimeEntryIndex = 0;
    ctx.state.selectedTsukihimeFileIndex = 0;
    ctx.state.currentTsukihimeEntryId = null;

    ctx.dom.tsukihimeEntriesList.innerHTML = '';
    ctx.dom.tsukihimeFilesList.innerHTML = '';
    ctx.dom.tsukihimeEntriesSection.classList.add('hidden');
    ctx.dom.tsukihimeFilesSection.classList.add('hidden');
  }

  // Defaults to English until the configured secondarySub languages arrive.
  let secondaryLanguages: string[] = ['en'];

  function secondaryTabLabel(): string {
    return describeTsukihimeTabLanguages(secondaryLanguages);
  }

  function isJapaneseTrack(file: TsukihimeSubtitleFile): boolean {
    return normalizeTsukihimeLangCode(file.lang) === 'ja';
  }

  function getVisibleFiles(): TsukihimeSubtitleFile[] {
    if (ctx.state.tsukihimeActiveTab === 'ja') {
      return ctx.state.tsukihimeFiles.filter(isJapaneseTrack);
    }
    return ctx.state.tsukihimeFiles.filter(
      (file) =>
        !isJapaneseTrack(file) && tsukihimeTrackMatchesLanguages(file.lang, secondaryLanguages),
    );
  }

  function renderTabs(): void {
    if (ctx.state.tsukihimeActiveTab === 'ja') {
      ctx.dom.tsukihimeTabEnglishButton.classList.remove('active');
      ctx.dom.tsukihimeTabJapaneseButton.classList.add('active');
    } else {
      ctx.dom.tsukihimeTabEnglishButton.classList.add('active');
      ctx.dom.tsukihimeTabJapaneseButton.classList.remove('active');
    }
  }

  function describeEmptyTab(): string {
    const hiddenCount = ctx.state.tsukihimeFiles.length;
    if (ctx.state.tsukihimeActiveTab === 'ja') {
      return hiddenCount > 0
        ? `No Japanese tracks in this release. Switch to the ${secondaryTabLabel()} tab.`
        : 'No Japanese tracks in this release.';
    }
    return hiddenCount > 0
      ? `No ${secondaryTabLabel()} tracks in this release. Switch to the Japanese tab.`
      : `No ${secondaryTabLabel()} tracks in this release.`;
  }

  function setActiveTab(tab: 'en' | 'ja'): void {
    if (ctx.state.tsukihimeActiveTab === tab) return;
    ctx.state.tsukihimeActiveTab = tab;
    ctx.state.selectedTsukihimeFileIndex = 0;
    renderTabs();

    if (ctx.state.tsukihimeFiles.length === 0) return;
    renderFiles();
    if (getVisibleFiles().length === 0) {
      setTsukihimeStatus(describeEmptyTab());
    } else {
      setTsukihimeStatus('Select a subtitle track.');
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
    ctx.dom.tsukihimeEntriesList.innerHTML = '';
    if (ctx.state.tsukihimeEntries.length === 0) {
      ctx.dom.tsukihimeEntriesSection.classList.add('hidden');
      return;
    }

    ctx.dom.tsukihimeEntriesSection.classList.remove('hidden');
    ctx.state.tsukihimeEntries.forEach((entry, index) => {
      const li = document.createElement('li');
      li.textContent = entry.title;

      const details: string[] = [];
      if (entry.totalSize !== null) details.push(formatBytes(entry.totalSize));
      if (entry.numFiles !== null) {
        details.push(`${entry.numFiles} file${entry.numFiles === 1 ? '' : 's'}`);
      }
      if (entry.sublangs?.length) details.push(`subs: ${entry.sublangs.join(', ')}`);
      if (details.length > 0) {
        const sub = document.createElement('div');
        sub.className = 'jimaku-subtext';
        sub.textContent = details.join(' • ');
        li.appendChild(sub);
      }

      if (index === ctx.state.selectedTsukihimeEntryIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        selectEntry(index);
      });

      ctx.dom.tsukihimeEntriesList.appendChild(li);
    });
  }

  function renderFiles(): void {
    ctx.dom.tsukihimeFilesList.innerHTML = '';
    const visibleFiles = getVisibleFiles();
    if (visibleFiles.length === 0) {
      ctx.dom.tsukihimeFilesSection.classList.add('hidden');
      return;
    }

    ctx.dom.tsukihimeFilesSection.classList.remove('hidden');
    visibleFiles.forEach((file, index) => {
      const li = document.createElement('li');
      li.textContent = file.filename;

      const details: string[] = [];
      if (file.lang) details.push(file.lang);
      if (file.trackName) details.push(file.trackName);
      // TsukiHime does not report attachment sizes; skip a meaningless "0 B".
      if (file.size > 0) details.push(formatBytes(file.size));
      const sub = document.createElement('div');
      sub.className = 'jimaku-subtext';
      sub.textContent = details.filter(Boolean).join(' • ');
      li.appendChild(sub);

      if (index === ctx.state.selectedTsukihimeFileIndex) {
        li.classList.add('active');
      }

      li.addEventListener('click', () => {
        void selectFile(index);
      });

      ctx.dom.tsukihimeFilesList.appendChild(li);
    });
  }

  function getSearchQuery(): string {
    const title = ctx.dom.tsukihimeTitleInput.value.trim();
    if (!title) return '';
    const episodeValue = ctx.dom.tsukihimeEpisodeInput.value
      ? Number.parseInt(ctx.dom.tsukihimeEpisodeInput.value, 10)
      : null;
    if (episodeValue !== null && Number.isFinite(episodeValue)) {
      return `${title} ${String(episodeValue).padStart(2, '0')}`;
    }
    return title;
  }

  async function performTsukihimeSearch(): Promise<void> {
    const query = getSearchQuery();
    if (!query) {
      setTsukihimeStatus('Enter a title before searching.', true);
      return;
    }

    resetTsukihimeLists();
    setTsukihimeStatus('Searching TsukiHime...');

    const response: TsukihimeApiResponse<TsukihimeEntry[]> =
      await window.electronAPI.tsukihimeSearchEntries({ query });
    if (!response.ok) {
      setTsukihimeStatus(response.error.error, true);
      return;
    }

    ctx.state.tsukihimeEntries = response.data;
    ctx.state.selectedTsukihimeEntryIndex = 0;

    if (ctx.state.tsukihimeEntries.length === 0) {
      setTsukihimeStatus('No releases found.');
      return;
    }

    setTsukihimeStatus('Select a release.');
    renderEntries();
    if (ctx.state.tsukihimeEntries.length === 1) {
      selectEntry(0);
    }
  }

  async function loadFiles(entryId: number): Promise<void> {
    setTsukihimeStatus('Loading subtitle tracks...');
    ctx.state.tsukihimeFiles = [];
    ctx.state.selectedTsukihimeFileIndex = 0;

    ctx.dom.tsukihimeFilesList.innerHTML = '';
    ctx.dom.tsukihimeFilesSection.classList.add('hidden');

    const response: TsukihimeApiResponse<TsukihimeSubtitleFile[]> =
      await window.electronAPI.tsukihimeListFiles({ entryId });
    // The user may have picked another release while this was in flight.
    if (ctx.state.currentTsukihimeEntryId !== entryId) return;
    if (!response.ok) {
      setTsukihimeStatus(response.error.error, true);
      return;
    }

    ctx.state.tsukihimeFiles = response.data;
    if (ctx.state.tsukihimeFiles.length === 0) {
      const entry = ctx.state.tsukihimeEntries.find((candidate) => candidate.id === entryId);
      // The feed API omits per-file attachment data for multi-file torrents.
      if (entry && entry.numFiles !== null && entry.numFiles > 1) {
        setTsukihimeStatus(
          'Batch releases are not supported. Pick a single-episode release instead.',
        );
      } else {
        setTsukihimeStatus('No text subtitle tracks in this release. Try another one.');
      }
      return;
    }

    const visibleFiles = getVisibleFiles();
    if (visibleFiles.length === 0) {
      setTsukihimeStatus(describeEmptyTab());
      return;
    }

    setTsukihimeStatus('Select a subtitle track.');
    renderFiles();
    if (visibleFiles.length === 1) {
      await selectFile(0);
    }
  }

  function selectEntry(index: number): void {
    if (index < 0 || index >= ctx.state.tsukihimeEntries.length) return;

    ctx.state.selectedTsukihimeEntryIndex = index;
    ctx.state.currentTsukihimeEntryId = ctx.state.tsukihimeEntries[index]!.id;
    renderEntries();

    if (ctx.state.currentTsukihimeEntryId !== null) {
      void loadFiles(ctx.state.currentTsukihimeEntryId);
    }
  }

  async function selectFile(index: number): Promise<void> {
    const visibleFiles = getVisibleFiles();
    if (index < 0 || index >= visibleFiles.length) return;

    ctx.state.selectedTsukihimeFileIndex = index;
    renderFiles();

    if (ctx.state.currentTsukihimeEntryId === null) {
      setTsukihimeStatus('Select a release first.', true);
      return;
    }

    const file = visibleFiles[index]!;
    setTsukihimeStatus('Downloading subtitle...');

    const result: TsukihimeDownloadResult = await window.electronAPI.tsukihimeDownloadFile({
      entryId: ctx.state.currentTsukihimeEntryId,
      url: file.url,
      name: file.filename,
      lang: file.lang,
    });

    if (result.ok) {
      setTsukihimeStatus(`Downloaded and loaded: ${result.path}`);
      closeTsukihimeModal();
      return;
    }

    setTsukihimeStatus(result.error.error, true);
  }

  function isTextInputFocused(): boolean {
    const active = document.activeElement;
    if (!active) return false;
    const tag = active.tagName.toLowerCase();
    return tag === 'input' || tag === 'textarea';
  }

  async function loadSecondaryLanguages(): Promise<void> {
    try {
      const languages = await window.electronAPI.tsukihimeGetSecondaryLanguages();
      secondaryLanguages = languages.length > 0 ? languages : ['en'];
    } catch {
      secondaryLanguages = ['en'];
    }
    ctx.dom.tsukihimeTabEnglishButton.textContent = secondaryTabLabel();
    // Tracks may already be on screen if the languages arrived late.
    if (ctx.state.tsukihimeFiles.length > 0) {
      renderFiles();
    }
  }

  function openTsukihimeModal(): void {
    if (ctx.state.tsukihimeModalOpen) return;

    ctx.state.tsukihimeModalOpen = true;
    ctx.state.tsukihimeActiveTab = 'en';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.tsukihimeModal.classList.remove('hidden');
    ctx.dom.tsukihimeModal.setAttribute('aria-hidden', 'false');

    setTsukihimeStatus('Loading media info...');
    resetTsukihimeLists();
    renderTabs();

    const secondaryLanguagesReady = loadSecondaryLanguages();

    window.electronAPI
      .getJimakuMediaInfo()
      .then(async (info: JimakuMediaInfo) => {
        ctx.dom.tsukihimeTitleInput.value = info.title || '';
        ctx.dom.tsukihimeEpisodeInput.value = info.episode ? String(info.episode) : '';

        if (info.confidence === 'high' && info.title && info.episode) {
          await secondaryLanguagesReady;
          void performTsukihimeSearch();
        } else if (info.title) {
          setTsukihimeStatus('Check title/episode and press Search.');
        } else {
          setTsukihimeStatus('Enter title/episode and press Search.');
        }
      })
      .catch(() => {
        setTsukihimeStatus('Failed to load media info.', true);
      });
  }

  function closeTsukihimeModal(): void {
    if (!ctx.state.tsukihimeModalOpen) return;

    ctx.state.tsukihimeModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.tsukihimeModal.classList.add('hidden');
    ctx.dom.tsukihimeModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('tsukihime');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }

    resetTsukihimeLists();
  }

  function handleTsukihimeKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeTsukihimeModal();
      return true;
    }

    if (isTextInputFocused()) {
      if (e.key === 'Enter') {
        e.preventDefault();
        void performTsukihimeSearch();
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
        ctx.state.selectedTsukihimeFileIndex = Math.min(
          visibleFiles.length - 1,
          ctx.state.selectedTsukihimeFileIndex + 1,
        );
        renderFiles();
      } else if (ctx.state.tsukihimeEntries.length > 0) {
        ctx.state.selectedTsukihimeEntryIndex = Math.min(
          ctx.state.tsukihimeEntries.length - 1,
          ctx.state.selectedTsukihimeEntryIndex + 1,
        );
        renderEntries();
      }
      return true;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (getVisibleFiles().length > 0) {
        ctx.state.selectedTsukihimeFileIndex = Math.max(
          0,
          ctx.state.selectedTsukihimeFileIndex - 1,
        );
        renderFiles();
      } else if (ctx.state.tsukihimeEntries.length > 0) {
        ctx.state.selectedTsukihimeEntryIndex = Math.max(
          0,
          ctx.state.selectedTsukihimeEntryIndex - 1,
        );
        renderEntries();
      }
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (getVisibleFiles().length > 0) {
        void selectFile(ctx.state.selectedTsukihimeFileIndex);
      } else if (ctx.state.tsukihimeEntries.length > 0) {
        selectEntry(ctx.state.selectedTsukihimeEntryIndex);
      } else {
        void performTsukihimeSearch();
      }
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.tsukihimeSearchButton.addEventListener('click', () => {
      void performTsukihimeSearch();
    });
    ctx.dom.tsukihimeCloseButton.addEventListener('click', () => {
      closeTsukihimeModal();
    });
    ctx.dom.tsukihimeTabEnglishButton.addEventListener('click', () => {
      setActiveTab('en');
    });
    ctx.dom.tsukihimeTabJapaneseButton.addEventListener('click', () => {
      setActiveTab('ja');
    });
  }

  return {
    closeTsukihimeModal,
    handleTsukihimeKeydown,
    openTsukihimeModal,
    selectTsukihimeEntry: selectEntry,
    wireDomEvents,
  };
}
