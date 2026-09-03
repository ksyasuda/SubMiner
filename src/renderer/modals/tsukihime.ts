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

  // Defaults to English until the configured secondary languages arrive.
  let secondaryLanguages: string[] = ['en'];
  // Both tab filters read the configured languages, so a search must wait for
  // them rather than filtering against the English fallback.
  let secondaryLanguagesReady: Promise<void> = Promise.resolve();
  let activeDownloadToken = 0;
  // Bumped by every new search and by closing the modal, so results that
  // arrive late cannot repopulate a reopened modal or a newer search.
  let activeSearchToken = 0;

  function secondaryTabLabel(): string {
    return describeTsukihimeTabLanguages(secondaryLanguages);
  }

  function isPrimaryTrack(file: TsukihimeSubtitleFile): boolean {
    return normalizeTsukihimeLangCode(file.lang) === 'ja';
  }

  function getVisibleFiles(): TsukihimeSubtitleFile[] {
    if (ctx.state.tsukihimeActiveTab === 'primary') {
      return ctx.state.tsukihimeFiles.filter(isPrimaryTrack);
    }
    return ctx.state.tsukihimeFiles.filter(
      (file) =>
        !isPrimaryTrack(file) && tsukihimeTrackMatchesLanguages(file.lang, secondaryLanguages),
    );
  }

  // Releases are filtered by the languages the search index reports for
  // them. Most releases carry no Japanese track, so the primary tab hides
  // them outright. A release with no language data cannot be classified and
  // stays on the secondary tab, mirroring how unlabeled tracks are handled.
  function entryMatchesTab(entry: TsukihimeEntry, tab: 'secondary' | 'primary'): boolean {
    if (tab === 'primary') {
      return entry.sublangs.some((lang) => normalizeTsukihimeLangCode(lang) === 'ja');
    }
    if (entry.sublangs.length === 0) return true;
    return entry.sublangs.some(
      (lang) =>
        normalizeTsukihimeLangCode(lang) !== 'ja' &&
        tsukihimeTrackMatchesLanguages(lang, secondaryLanguages),
    );
  }

  function getVisibleEntries(): TsukihimeEntry[] {
    return ctx.state.tsukihimeEntries.filter((entry) =>
      entryMatchesTab(entry, ctx.state.tsukihimeActiveTab),
    );
  }

  function describeEmptyReleases(): string {
    const otherTab = ctx.state.tsukihimeActiveTab === 'primary' ? 'secondary' : 'primary';
    const otherTabHasReleases = ctx.state.tsukihimeEntries.some((entry) =>
      entryMatchesTab(entry, otherTab),
    );
    const language = ctx.state.tsukihimeActiveTab === 'primary' ? 'Japanese' : secondaryTabLabel();
    const otherLabel = otherTab === 'primary' ? 'Japanese' : secondaryTabLabel();
    return otherTabHasReleases
      ? `No releases with ${language} subtitles. Switch to the ${otherLabel} tab.`
      : `No releases with ${language} subtitles.`;
  }

  function clearFiles(): void {
    ctx.state.tsukihimeFiles = [];
    ctx.state.selectedTsukihimeFileIndex = 0;
    ctx.dom.tsukihimeFilesList.innerHTML = '';
    ctx.dom.tsukihimeFilesSection.classList.add('hidden');
  }

  function renderTabs(): void {
    const primaryActive = ctx.state.tsukihimeActiveTab === 'primary';
    ctx.dom.tsukihimeTabSecondaryButton.setAttribute(
      'aria-selected',
      primaryActive ? 'false' : 'true',
    );
    ctx.dom.tsukihimeTabPrimaryButton.setAttribute(
      'aria-selected',
      primaryActive ? 'true' : 'false',
    );
    if (primaryActive) {
      ctx.dom.tsukihimeTabSecondaryButton.classList.remove('active');
      ctx.dom.tsukihimeTabPrimaryButton.classList.add('active');
    } else {
      ctx.dom.tsukihimeTabSecondaryButton.classList.add('active');
      ctx.dom.tsukihimeTabPrimaryButton.classList.remove('active');
    }
  }

  function describeEmptyTab(): string {
    const hiddenCount = ctx.state.tsukihimeFiles.length;
    if (ctx.state.tsukihimeActiveTab === 'primary') {
      return hiddenCount > 0
        ? `No Japanese tracks in this release. Switch to the ${secondaryTabLabel()} tab.`
        : 'No Japanese tracks in this release.';
    }
    return hiddenCount > 0
      ? `No ${secondaryTabLabel()} tracks in this release. Switch to the Japanese tab.`
      : `No ${secondaryTabLabel()} tracks in this release.`;
  }

  function setActiveTab(tab: 'secondary' | 'primary'): void {
    if (ctx.state.tsukihimeActiveTab === tab) return;
    ctx.state.tsukihimeActiveTab = tab;
    ctx.state.selectedTsukihimeFileIndex = 0;
    renderTabs();

    const currentEntry = ctx.state.tsukihimeEntries.find(
      (entry) => entry.id === ctx.state.currentTsukihimeEntryId,
    );
    if (currentEntry && !entryMatchesTab(currentEntry, tab)) {
      // The selected release is hidden on this tab; drop its tracks so the
      // list matches what the tab claims to show.
      ctx.state.currentTsukihimeEntryId = null;
      ctx.state.selectedTsukihimeEntryIndex = 0;
      clearFiles();
      renderEntries();
      setTsukihimeStatus(
        getVisibleEntries().length === 0 ? describeEmptyReleases() : 'Select a release.',
      );
      return;
    }

    const visibleEntries = getVisibleEntries();
    ctx.state.selectedTsukihimeEntryIndex = currentEntry ? visibleEntries.indexOf(currentEntry) : 0;
    renderEntries();

    if (ctx.state.tsukihimeFiles.length > 0) {
      renderFiles();
      setTsukihimeStatus(
        getVisibleFiles().length === 0 ? describeEmptyTab() : 'Select a subtitle track.',
      );
      return;
    }
    if (!currentEntry && ctx.state.tsukihimeEntries.length > 0) {
      setTsukihimeStatus(
        visibleEntries.length === 0 ? describeEmptyReleases() : 'Select a release.',
      );
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
    const visibleEntries = getVisibleEntries();
    if (visibleEntries.length === 0) {
      ctx.dom.tsukihimeEntriesSection.classList.add('hidden');
      return;
    }

    ctx.dom.tsukihimeEntriesSection.classList.remove('hidden');
    visibleEntries.forEach((entry, index) => {
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

  function readNumericInput(input: HTMLInputElement): number | null {
    if (!input.value) return null;
    const parsed = Number.parseInt(input.value, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  function getSearchQuery(): string {
    const title = ctx.dom.tsukihimeTitleInput.value.trim();
    if (!title) return '';
    const season = readNumericInput(ctx.dom.tsukihimeSeasonInput);
    const episode = readNumericInput(ctx.dom.tsukihimeEpisodeInput);
    // Season 1 is left out on purpose: releases of a first season almost never
    // put "S01" in the name, so adding it narrows the search to nothing.
    const parts = [title];
    if (season !== null && season > 1) {
      parts.push(`S${String(season).padStart(2, '0')}`);
    }
    if (episode !== null) {
      parts.push(String(episode).padStart(2, '0'));
    }
    return parts.join(' ');
  }

  async function performTsukihimeSearch(): Promise<void> {
    const query = getSearchQuery();
    if (!query) {
      setTsukihimeStatus('Enter a title before searching.', true);
      return;
    }

    const searchToken = ++activeSearchToken;
    resetTsukihimeLists();
    setTsukihimeStatus('Searching TsukiHime...');
    await secondaryLanguagesReady;
    if (searchToken !== activeSearchToken) return;

    const response: TsukihimeApiResponse<TsukihimeEntry[]> =
      await window.electronAPI.tsukihimeSearchEntries({ query });
    if (searchToken !== activeSearchToken) return;
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

    const visibleEntries = getVisibleEntries();
    if (visibleEntries.length === 0) {
      setTsukihimeStatus(describeEmptyReleases());
      return;
    }

    setTsukihimeStatus('Select a release.');
    renderEntries();
    if (visibleEntries.length === 1) {
      selectEntry(0);
    }
  }

  async function loadFiles(entryId: number): Promise<void> {
    setTsukihimeStatus('Loading subtitle tracks...');
    clearFiles();

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

  // `index` addresses the entries visible on the active tab, not the full
  // search result list.
  function selectEntry(index: number): void {
    const visibleEntries = getVisibleEntries();
    if (index < 0 || index >= visibleEntries.length) return;

    ctx.state.selectedTsukihimeEntryIndex = index;
    ctx.state.currentTsukihimeEntryId = visibleEntries[index]!.id;
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

    const entryId = ctx.state.currentTsukihimeEntryId;
    const file = visibleFiles[index]!;
    const downloadToken = ++activeDownloadToken;
    setTsukihimeStatus('Downloading subtitle...');

    const result: TsukihimeDownloadResult = await window.electronAPI.tsukihimeDownloadFile({
      entryId,
      url: file.url,
      name: file.filename,
      lang: file.lang,
    });
    if (downloadToken !== activeDownloadToken || ctx.state.currentTsukihimeEntryId !== entryId) {
      return;
    }

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
      const configuredSecondary = await window.electronAPI.tsukihimeGetSecondaryLanguages();
      secondaryLanguages = configuredSecondary.length > 0 ? configuredSecondary : ['en'];
    } catch {
      secondaryLanguages = ['en'];
    }
    ctx.dom.tsukihimeTabSecondaryButton.textContent = secondaryTabLabel();
    ctx.dom.tsukihimeTabPrimaryButton.textContent = 'Japanese';
    // Tracks may already be on screen if the languages arrived late.
    if (ctx.state.tsukihimeFiles.length > 0) {
      renderFiles();
    }
  }

  function openTsukihimeModal(): void {
    if (ctx.state.tsukihimeModalOpen) return;

    ctx.state.tsukihimeModalOpen = true;
    ctx.state.tsukihimeActiveTab = 'secondary';
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.tsukihimeModal.classList.remove('hidden');
    ctx.dom.tsukihimeModal.setAttribute('aria-hidden', 'false');

    setTsukihimeStatus('Loading media info...');
    resetTsukihimeLists();
    renderTabs();

    secondaryLanguagesReady = loadSecondaryLanguages();

    window.electronAPI
      .getJimakuMediaInfo()
      .then((info: JimakuMediaInfo) => {
        ctx.dom.tsukihimeTitleInput.value = info.title || '';
        ctx.dom.tsukihimeSeasonInput.value = info.season ? String(info.season) : '';
        ctx.dom.tsukihimeEpisodeInput.value = info.episode ? String(info.episode) : '';

        if (info.confidence === 'high' && info.title && info.episode) {
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

    activeDownloadToken += 1;
    activeSearchToken += 1;
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
      setActiveTab('secondary');
      return true;
    }

    if (e.key === 'ArrowRight') {
      e.preventDefault();
      setActiveTab('primary');
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
      } else if (getVisibleEntries().length > 0) {
        ctx.state.selectedTsukihimeEntryIndex = Math.min(
          getVisibleEntries().length - 1,
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
      } else if (getVisibleEntries().length > 0) {
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
      } else if (getVisibleEntries().length > 0) {
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
    ctx.dom.tsukihimeTabSecondaryButton.addEventListener('click', () => {
      setActiveTab('secondary');
    });
    ctx.dom.tsukihimeTabPrimaryButton.addEventListener('click', () => {
      setActiveTab('primary');
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
