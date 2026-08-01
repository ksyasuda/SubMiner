import { describe, el } from './dom';
import { sourceOptionLabel, summarizeSearch } from './format';
import { applySearchUpdate, idleSearchProgress, summarizeProgress } from './search-progress';
import { createExtensionsPanel } from './extensions-panel';
import { renderPreferences, renderPreferencesUnavailable } from './preferences-fields';
import { ALL_SOURCES_ID } from '../types/anime-browser';
import type {
  AnimeBrowserAPI,
  AnimeBrowserBridgeState,
  AnimeBrowserEntry,
  AnimeBrowserEpisode,
  AnimeBrowserSource,
} from '../types/anime-browser';

declare global {
  interface Window {
    animeBrowserAPI: AnimeBrowserAPI;
  }
}

const api = window.animeBrowserAPI;

const searchForm = el<HTMLFormElement>('search-form');
const searchInput = el<HTMLInputElement>('search-input');
const searchButton = el<HTMLButtonElement>('search-button');
const sourceSelect = el<HTMLSelectElement>('source-select');
const grid = el<HTMLDivElement>('grid');
const gridEmpty = el<HTMLParagraphElement>('grid-empty');
const results = el<HTMLElement>('results');
const detail = el<HTMLElement>('detail');
const detailBack = el<HTMLButtonElement>('detail-back');
const detailCover = el<HTMLImageElement>('detail-cover');
const detailTitle = el<HTMLHeadingElement>('detail-title');
const detailChips = el<HTMLDivElement>('detail-chips');
const detailDescription = el<HTMLParagraphElement>('detail-description');
const episodes = el<HTMLOListElement>('episodes');
const episodesCount = el<HTMLSpanElement>('episodes-count');
const banner = el<HTMLDivElement>('bridge-banner');
const bannerMessage = el<HTMLSpanElement>('bridge-message');
const bannerMeter = el<HTMLSpanElement>('bridge-meter');
const bannerMeterFill = el<HTMLElement>('bridge-meter-fill');
const statusMessage = el<HTMLSpanElement>('status-message');
const browseTab = el<HTMLButtonElement>('tab-browse');
const extensionsTab = el<HTMLButtonElement>('tab-extensions');
const settingsTab = el<HTMLButtonElement>('tab-settings');
const layout = el<HTMLElement>('layout');
const extensionsPanel = el<HTMLElement>('extensions');
const settingsPanel = el<HTMLElement>('settings');
const settingsFields = el<HTMLDivElement>('settings-fields');
const settingsTitle = el<HTMLHeadingElement>('settings-title');

/** The anime the detail page is showing, with the source that produced it. */
let selectedAnime: { url: string; title: string; sourceId: string } | null = null;

/** Where the results grid was scrolled to before the detail page covered it. */
let resultsScrollTop = 0;

/* ---------- tabs ---------- */

type View = 'browse' | 'extensions' | 'settings';

let currentView: View = 'browse';

/**
 * Show one view at a time. The panels used to sit above the results, which left
 * a long extension list scrolling inside a sliver of the window; as tabs each
 * one gets the whole content region.
 */
function setView(view: View): void {
  currentView = view;
  layout.classList.toggle('hidden', view !== 'browse');
  extensionsPanel.classList.toggle('hidden', view !== 'extensions');
  settingsPanel.classList.toggle('hidden', view !== 'settings');
  browseTab.setAttribute('aria-selected', String(view === 'browse'));
  extensionsTab.setAttribute('aria-selected', String(view === 'extensions'));
  settingsTab.setAttribute('aria-selected', String(view === 'settings'));
}

function setStatus(message: string, tone: 'info' | 'ok' | 'error' = 'info'): void {
  statusMessage.textContent = message;
  statusMessage.parentElement?.setAttribute('data-tone', tone);
}

const BRIDGE_LABELS: Record<AnimeBrowserBridgeState['stage'], string> = {
  idle: 'Starting the extension bridge',
  locating: 'Looking up the extension bridge release',
  downloading: 'Downloading the extension bridge',
  verifying: 'Verifying the download',
  extracting: 'Unpacking the extension bridge',
  starting: 'Starting the extension bridge',
  ready: 'Bridge ready',
  failed: 'Bridge failed to start',
};

const BUSY_STAGES = new Set([
  'idle',
  'locating',
  'downloading',
  'verifying',
  'extracting',
  'starting',
]);

function renderBridgeState(state: AnimeBrowserBridgeState): void {
  const busy = BUSY_STAGES.has(state.stage);
  banner.dataset.stage = state.stage;
  banner.dataset.busy = String(busy);

  // Once ready with nothing to report, the banner has nothing to say.
  const hide = state.stage === 'ready' && state.message === null;
  banner.classList.toggle('hidden', hide);
  bannerMessage.textContent = state.message ?? BRIDGE_LABELS[state.stage];

  const showMeter = state.progress !== null;
  bannerMeter.classList.toggle('hidden', !showMeter);
  if (state.progress !== null) {
    bannerMeterFill.style.width = `${Math.round(state.progress * 100)}%`;
  }

  const ready = state.stage === 'ready';
  searchInput.disabled = !ready;
  searchButton.disabled = !ready;
}

/* ---------- source picker ---------- */

/**
 * The picker lists every installed source, plus an "All sources" entry that
 * searches them together. That entry only earns its place with more than one
 * source installed.
 */
function renderSources(sources: AnimeBrowserSource[], selectedId: string | null): void {
  const options: HTMLOptionElement[] = [];

  if (sources.length > 1) {
    const all = document.createElement('option');
    all.value = ALL_SOURCES_ID;
    all.textContent = `All sources (${sources.length})`;
    all.selected = selectedId === ALL_SOURCES_ID;
    options.push(all);
  }

  for (const source of sources) {
    const option = document.createElement('option');
    option.value = source.id;
    option.textContent = sourceOptionLabel(source);
    option.selected = source.id === selectedId;
    options.push(option);
  }

  sourceSelect.replaceChildren(...options);
  sourceSelect.disabled = options.length <= 1;
}

function searchingAllSources(): boolean {
  return sourceSelect.value === ALL_SOURCES_ID;
}

/* ---------- results ---------- */

function createCard(entry: AnimeBrowserEntry, showSource: boolean): HTMLButtonElement {
  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'card';

  const art = document.createElement('div');
  art.className = 'card-art';
  if (entry.thumbnailUrl) {
    const img = document.createElement('img');
    img.src = entry.thumbnailUrl;
    img.alt = '';
    img.loading = 'lazy';
    // A dead cover should fall back to the initial, not a broken-image icon.
    img.addEventListener('error', () => {
      img.remove();
      art.classList.add('is-empty');
    });
    art.append(img);
  } else {
    art.classList.add('is-empty');
  }
  art.dataset.initial = entry.title.trim().charAt(0).toUpperCase() || '?';

  if (showSource) {
    const badge = document.createElement('span');
    badge.className = 'card-source';
    badge.textContent = entry.sourceName;
    art.append(badge);
  }

  const title = document.createElement('div');
  title.className = 'card-title';
  title.textContent = entry.title;

  card.append(art, title);
  card.title = showSource ? `${entry.title} — ${entry.sourceName}` : entry.title;
  card.addEventListener('click', () => {
    void openDetail(entry);
  });
  return card;
}

function renderEntries(entries: AnimeBrowserEntry[], emptyMessage: string): void {
  // Which source a cover came from only matters when they are mixed together.
  const showSource = searchingAllSources();
  grid.replaceChildren(...entries.map((entry) => createCard(entry, showSource)));

  const empty = entries.length === 0;
  gridEmpty.classList.toggle('hidden', !empty);
  gridEmpty.textContent = emptyMessage;
}

/** Streamed results land at the end of the grid, in arrival order. */
function appendEntries(entries: AnimeBrowserEntry[]): void {
  const showSource = searchingAllSources();
  grid.append(...entries.map((entry) => createCard(entry, showSource)));
}

function formatEpisodeIndex(episode: AnimeBrowserEpisode, fallbackIndex: number): string {
  const value = episode.number ?? fallbackIndex;
  return Number.isInteger(value) ? String(value).padStart(2, '0') : value.toFixed(1);
}

function renderEpisodes(list: AnimeBrowserEpisode[]): void {
  episodesCount.textContent = list.length === 0 ? '' : `${list.length}`;
  episodes.replaceChildren(
    ...list.map((episode, index) => {
      const item = document.createElement('li');
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'cue';

      const cueIndex = document.createElement('span');
      cueIndex.className = 'cue-index';
      cueIndex.textContent = formatEpisodeIndex(episode, list.length - index);

      const name = document.createElement('span');
      name.className = 'cue-name';
      name.textContent = episode.name;
      if (episode.uploadedAt !== null) {
        const sub = document.createElement('span');
        sub.className = 'cue-sub';
        sub.textContent = new Date(episode.uploadedAt).toISOString().slice(0, 10);
        name.append(sub);
      }

      button.append(cueIndex, name);
      button.addEventListener('click', () => {
        void playEpisode(button, episode);
      });
      item.append(button);
      return item;
    }),
  );
}

/**
 * The detail page replaces the results grid rather than squeezing in beside
 * it. The grid stays in the DOM with its scroll position remembered, so Back
 * returns to the same results without re-running the search.
 */
async function openDetail(entry: AnimeBrowserEntry): Promise<void> {
  selectedAnime = { url: entry.url, title: entry.title, sourceId: entry.sourceId };
  resultsScrollTop = results.scrollTop;
  results.classList.add('hidden');
  detail.classList.remove('hidden');
  detail.scrollTop = 0;
  detailTitle.textContent = entry.title;
  detailDescription.textContent = 'Loading…';
  detailChips.replaceChildren();
  episodes.replaceChildren();
  episodesCount.textContent = '';
  detailCover.src = entry.thumbnailUrl ?? '';

  try {
    // Always ask the entry's own source: after an all-sources search the
    // picker's selection says nothing about where this cover came from.
    const [details, episodeList] = await Promise.all([
      api.getDetails(entry.url, entry.sourceId),
      api.getEpisodes(entry.url, entry.sourceId),
    ]);

    detailTitle.textContent = details.title;
    detailDescription.textContent = details.description ?? 'No description from this source.';
    if (details.thumbnailUrl) detailCover.src = details.thumbnailUrl;

    const chips: HTMLSpanElement[] = [];
    const source = document.createElement('span');
    source.className = 'chip source';
    source.textContent = entry.sourceName;
    chips.push(source);
    if (details.status !== 'unknown') {
      const status = document.createElement('span');
      status.className = 'chip status';
      status.textContent = details.status.replace(/-/g, ' ');
      chips.push(status);
    }
    for (const genre of details.genres.slice(0, 6)) {
      const chip = document.createElement('span');
      chip.className = 'chip';
      chip.textContent = genre;
      chips.push(chip);
    }
    detailChips.replaceChildren(...chips);

    renderEpisodes(episodeList);
    setStatus(`${details.title} · ${episodeList.length} episodes`);
  } catch (error) {
    detailDescription.textContent = '';
    setStatus(describe(error), 'error');
  }
}

function closeDetail(): void {
  detail.classList.add('hidden');
  results.classList.remove('hidden');
  results.scrollTop = resultsScrollTop;
  selectedAnime = null;
}

async function playEpisode(button: HTMLButtonElement, episode: AnimeBrowserEpisode): Promise<void> {
  if (!selectedAnime) return;

  for (const other of episodes.querySelectorAll<HTMLButtonElement>('.cue')) {
    other.removeAttribute('data-state');
  }
  button.dataset.state = 'loading';
  setStatus(`Resolving ${episode.name}…`);

  const result = await api.playEpisode({
    sourceId: selectedAnime.sourceId,
    animeUrl: selectedAnime.url,
    animeTitle: selectedAnime.title,
    episodeUrl: episode.url,
    episodeName: episode.name,
  });

  if (result.ok) {
    button.dataset.state = 'playing';
    setStatus(
      result.quality ? `Playing ${episode.name} · ${result.quality}` : `Playing ${episode.name}`,
      'ok',
    );
  } else {
    button.removeAttribute('data-state');
    setStatus(result.error ?? 'Could not play that episode.', 'error');
  }
}

/* ---------- streamed search ---------- */

/**
 * Results are pushed per source while the search call is still pending, so a
 * fast source is on screen before a slow one answers. The grid is built from
 * those pushes; the awaited result then settles the final status line (and
 * backfills the grid if no push ever arrived).
 */
let progress = idleSearchProgress();

/** Orders runSearch calls so a slow search cannot finish over a newer one. */
let searchRequest = 0;

api.onSearchUpdate((update) => {
  const applied = applySearchUpdate(progress, update);
  if (!applied) return; // A superseded search; let it run out quietly.
  progress = applied.progress;

  if (applied.started) {
    grid.replaceChildren();
    gridEmpty.classList.add('hidden');
    return;
  }
  if (applied.entries.length > 0) appendEntries(applied.entries);
  if (!progress.done) {
    setStatus(summarizeProgress(progress), progress.failures.length > 0 ? 'error' : 'info');
  }
});

async function runSearch(query: string): Promise<void> {
  const request = ++searchRequest;
  // A new search means new results; leave the detail page for them.
  if (!detail.classList.contains('hidden')) closeDetail();
  setStatus(query ? `Searching for “${query}”…` : 'Loading popular…');
  grid.replaceChildren();
  gridEmpty.classList.add('hidden');

  try {
    const result = query ? await api.search(query) : await api.getPopular();
    // A newer search owns the grid now; this one's result is history.
    if (request !== searchRequest) return;

    // Every source failing is an error, not an empty result set.
    if (result.entries.length === 0 && result.failures.length > 0) {
      const first = result.failures[0];
      renderEntries([], `${first?.sourceName}: ${first?.error}`);
      setStatus(summarizeSearch(result), 'error');
      return;
    }

    // The stream already filled the grid; only render from the result when no
    // update arrived (covers a host without streaming wired up).
    if (grid.childElementCount === 0 || result.entries.length === 0) {
      renderEntries(
        result.entries,
        query ? `Nothing found for “${query}”.` : 'This source returned nothing.',
      );
    }
    setStatus(summarizeSearch(result), result.failures.length > 0 ? 'error' : 'info');
  } catch (error) {
    if (request !== searchRequest) return;
    renderEntries([], describe(error));
    setStatus(describe(error), 'error');
  }
}

/* ---------- source settings ---------- */

async function openSettings(): Promise<void> {
  const option = sourceSelect.selectedOptions[0];
  settingsTitle.textContent = option ? `${option.textContent} settings` : 'Source settings';
  settingsFields.replaceChildren();
  setView('settings');

  // Settings belong to one extension, so there is nothing coherent to show for
  // "All sources".
  if (searchingAllSources()) {
    renderPreferencesUnavailable(
      settingsFields,
      'Pick a single source in the Source picker to edit its settings.',
    );
    return;
  }

  const sourceId = sourceSelect.value;
  try {
    renderPreferences(settingsFields, await api.getPreferences(sourceId), (key, value) =>
      api.setPreference(sourceId, key, value),
    );
  } catch (error) {
    setStatus(describe(error), 'error');
    setView('browse');
  }
}

/* ---------- wiring ---------- */

async function refreshSources(): Promise<void> {
  const snapshot = await api.getSnapshot();
  renderSources(snapshot.sources, snapshot.selectedSourceId);
}

const extensions = createExtensionsPanel({ api, setStatus, onSourcesChanged: refreshSources });

browseTab.addEventListener('click', () => setView('browse'));

settingsTab.addEventListener('click', () => void openSettings());

extensionsTab.addEventListener('click', () => {
  setView('extensions');
  void extensions.refresh();
});

searchForm.addEventListener('submit', (event) => {
  event.preventDefault();
  // Searching is a browse action, whichever tab it was typed from.
  setView('browse');
  void runSearch(searchInput.value.trim());
});

sourceSelect.addEventListener('change', () => {
  void (async () => {
    await api.selectSource(sourceSelect.value);
    // Settings belong to the source, so reload them rather than showing stale fields.
    if (currentView === 'settings') await openSettings();
    await runSearch(searchInput.value.trim());
  })();
});

detailBack.addEventListener('click', closeDetail);

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && !detail.classList.contains('hidden')) {
    closeDetail();
  }
});

api.onBridgeState(renderBridgeState);

void (async () => {
  renderBridgeState({ stage: 'idle', progress: null, message: null });
  try {
    const state = await api.ensureBridge();
    renderBridgeState(state);

    const snapshot = await api.getSnapshot();
    renderSources(snapshot.sources, snapshot.selectedSourceId);

    if (state.stage === 'ready' && snapshot.sources.length > 0) {
      searchInput.focus();
      await runSearch('');
    } else if (state.stage === 'ready') {
      setStatus(state.message ?? 'No extensions installed.', 'error');
    } else {
      setStatus(state.message ?? 'The extension bridge is not available.', 'error');
    }
  } catch (error) {
    // Without this the window keeps the "starting" banner up forever, with the
    // search box disabled and nothing saying why.
    renderBridgeState({ stage: 'failed', progress: null, message: describe(error) });
    setStatus(describe(error), 'error');
  }
})();
