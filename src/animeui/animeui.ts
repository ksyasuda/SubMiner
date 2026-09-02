import { describe, el } from './dom';
import { sourceOptionLabel, summarizeSearch } from './format';
import { applySearchUpdate, idleSearchProgress, summarizeProgress } from './search-progress';
import { createExtensionsPanel } from './extensions-panel';
import { createDetailPanel } from './detail-panel';
import { renderPreferences, renderPreferencesUnavailable } from './preferences-fields';
import {
  beginBrowse,
  beginNextPage,
  createBrowseState,
  failBrowse,
  finishBrowse,
  soleBrowseRequest,
  takeUnseenEntries,
} from './browse-state';
import type { BrowseRequest } from './browse-state';
import { ALL_SOURCES_ID } from '../types/anime-browser';
import type {
  AnimeBrowserAPI,
  AnimeBrowserBridgeState,
  AnimeBrowserEntry,
  AnimeBrowserSource,
} from '../types/anime-browser';
import {
  ANIME_BROWSER_CLOSE_MESSAGE,
  createAnimeBrowserKeydownMessage,
} from '../shared/anime-browser-embed';

const embeddedInOverlay =
  new URLSearchParams(window.location.search).get('embedded') === 'overlay-modal';
if (embeddedInOverlay) document.documentElement.dataset.embedded = 'overlay-modal';

function resolveAnimeBrowserAPI(): AnimeBrowserAPI {
  const local = (window as Window & { animeBrowserAPI?: AnimeBrowserAPI }).animeBrowserAPI;
  if (local) return local;
  try {
    const parent = (window.parent as Window & { animeBrowserAPI?: AnimeBrowserAPI })
      .animeBrowserAPI;
    if (parent) return parent;
  } catch {
    // The embedded document is expected to be same-origin. Fall through to a
    // direct error if Chromium blocks that relationship unexpectedly.
  }
  throw new Error('Anime Browser API is unavailable.');
}

const api = resolveAnimeBrowserAPI();

const searchForm = el<HTMLFormElement>('search-form');
const searchInput = el<HTMLInputElement>('search-input');
const searchButton = el<HTMLButtonElement>('search-button');
const sourceSelect = el<HTMLSelectElement>('source-select');
const grid = el<HTMLDivElement>('grid');
const gridEmpty = el<HTMLParagraphElement>('grid-empty');
const loadMoreButton = el<HTMLButtonElement>('load-more');
const banner = el<HTMLDivElement>('bridge-banner');
const bannerMessage = el<HTMLSpanElement>('bridge-message');
const bannerMeter = el<HTMLSpanElement>('bridge-meter');
const bannerMeterFill = el<HTMLElement>('bridge-meter-fill');
const bannerUpdate = el<HTMLButtonElement>('bridge-update');
const statusMessage = el<HTMLSpanElement>('status-message');
const browseTab = el<HTMLButtonElement>('tab-browse');
const extensionsTab = el<HTMLButtonElement>('tab-extensions');
const settingsTab = el<HTMLButtonElement>('tab-settings');
const layout = el<HTMLElement>('layout');
const extensionsPanel = el<HTMLElement>('extensions');
const settingsPanel = el<HTMLElement>('settings');
const settingsFields = el<HTMLDivElement>('settings-fields');
const settingsTitle = el<HTMLHeadingElement>('settings-title');

/** Last source accepted by the main process, used to roll back a rejected change. */
let selectedSourceId: string | null = null;

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

const detailPanel = createDetailPanel({ api, setStatus });

if (embeddedInOverlay) {
  const overlayOrigin = window.location.origin;
  // Keyboard events do not bubble out of an iframe. Forward the physical key
  // chord so the overlay can apply the user's configured Anime Browser binding.
  document.addEventListener(
    'keydown',
    (event) => {
      if (event.key === 'Escape' && detailPanel.isOpen()) return;
      window.parent.postMessage(createAnimeBrowserKeydownMessage(event), overlayOrigin);
    },
    true,
  );
  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape' || event.defaultPrevented || detailPanel.isOpen()) return;
    event.preventDefault();
    window.parent.postMessage(ANIME_BROWSER_CLOSE_MESSAGE, overlayOrigin);
  });
}

const BRIDGE_LABELS: Record<AnimeBrowserBridgeState['stage'], string> = {
  idle: 'Starting the extension bridge',
  locating: 'Looking up the extension bridge release',
  downloading: 'Downloading the extension bridge',
  extracting: 'Unpacking the extension bridge',
  starting: 'Starting the extension bridge',
  ready: 'Bridge ready',
  failed: 'Bridge failed to start',
};

const BUSY_STAGES = new Set(['idle', 'locating', 'downloading', 'extracting', 'starting']);

function renderBridgeState(state: AnimeBrowserBridgeState): void {
  const busy = BUSY_STAGES.has(state.stage);
  banner.dataset.stage = state.stage;
  banner.dataset.busy = String(busy);

  // An update is only offered from a running bridge; mid-start it would race
  // the start it interrupts.
  const update = state.stage === 'ready' ? (state.install?.updateAvailable ?? null) : null;
  // Once ready with nothing to report, the banner has nothing to say.
  const hide = state.stage === 'ready' && state.message === null && update === null;
  banner.classList.toggle('hidden', hide);
  bannerMessage.textContent =
    state.message ??
    (update === null
      ? BRIDGE_LABELS[state.stage]
      : `Extension bridge ${state.install?.version ?? 'of unknown version'} is installed; ${update} is available.`);
  bannerUpdate.classList.toggle('hidden', update === null);
  bannerUpdate.textContent = update === null ? '' : `Update to ${update}`;
  bannerUpdate.disabled = busy;

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
  selectedSourceId = selectedId;
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
    void detailPanel.open(entry);
  });
  return card;
}

/** Drops every card and the empty-state message so a fresh page can build up. */
function resetGrid(): void {
  seenEntries.clear();
  grid.replaceChildren();
  gridEmpty.classList.add('hidden');
}

function renderEntries(entries: AnimeBrowserEntry[], emptyMessage: string): void {
  // Which source a cover came from only matters when they are mixed together.
  resetGrid();
  appendEntries(entries);

  const empty = grid.childElementCount === 0;
  gridEmpty.classList.toggle('hidden', !empty);
  gridEmpty.textContent = emptyMessage;
}

/** Streamed results land at the end of the grid, in arrival order. */
function appendEntries(entries: AnimeBrowserEntry[]): void {
  const showSource = searchingAllSources();
  const unseen = takeUnseenEntries(entries, seenEntries);
  grid.append(...unseen.map((entry) => createCard(entry, showSource)));
  if (unseen.length > 0) gridEmpty.classList.add('hidden');
}

/* ---------- streamed search ---------- */

/**
 * Results are pushed per source while the search call is still pending, so a
 * fast source is on screen before a slow one answers. The grid is built from
 * those pushes; the awaited result then settles the final status line (and
 * backfills the grid if no push ever arrived).
 */
let progress = idleSearchProgress();

let browseState = createBrowseState();
const inFlightBrowses = new Map<number, BrowseRequest>();
let activeStreamRequestId = 0;
const seenEntries = new Set<string>();

function renderLoadMore(): void {
  const loadingNextPage = browseState.loading && browseState.page > 1;
  loadMoreButton.classList.toggle('hidden', !browseState.hasNextPage && !loadingNextPage);
  loadMoreButton.disabled = browseState.loading;
  loadMoreButton.textContent = loadingNextPage ? 'Loading…' : 'Load more';
}

api.onSearchUpdate((update) => {
  const applied = applySearchUpdate(progress, update);
  if (!applied) return; // A superseded search; let it run out quietly.
  progress = applied.progress;

  if (applied.started) {
    // The runtime token does not carry the renderer request id. When calls
    // overlap their starts may arrive in either order, so stream only when the
    // association is unambiguous; the final response still backfills the grid.
    const request = soleBrowseRequest(inFlightBrowses);
    activeStreamRequestId = request?.id ?? 0;
    const append = request?.append === true;
    if (!append) resetGrid();
    return;
  }
  if (activeStreamRequestId !== browseState.requestId) return;
  if (applied.entries.length > 0) appendEntries(applied.entries);
  if (!progress.done) {
    setStatus(summarizeProgress(progress), progress.failures.length > 0 ? 'error' : 'info');
  }
});

async function runBrowse(request: BrowseRequest): Promise<void> {
  inFlightBrowses.set(request.id, request);
  renderLoadMore();

  try {
    const result = request.query
      ? await api.search(request.query, request.page)
      : await api.getPopular(request.page);
    if (request.id !== browseState.requestId) return;
    browseState = finishBrowse(browseState, request.id, result.hasNextPage);
    renderLoadMore();

    // Every source failing is an error, not an empty result set.
    if (result.entries.length === 0 && result.failures.length > 0) {
      const first = result.failures[0];
      if (!request.append) renderEntries([], `${first?.sourceName}: ${first?.error}`);
      setStatus(summarizeSearch(result), 'error');
      return;
    }

    // The final response backfills a host without streaming. Entries already
    // pushed by the stream are filtered out, including sources that repeat a
    // final page while another source still has more.
    appendEntries(result.entries);
    if (!request.append && grid.childElementCount === 0) {
      renderEntries(
        [],
        request.query ? `Nothing found for “${request.query}”.` : 'This source returned nothing.',
      );
    }
    setStatus(summarizeSearch(result), result.failures.length > 0 ? 'error' : 'info');
  } catch (error) {
    if (request.id !== browseState.requestId) return;
    browseState = failBrowse(browseState, request);
    renderLoadMore();
    if (!request.append) renderEntries([], describe(error));
    setStatus(describe(error), 'error');
  } finally {
    inFlightBrowses.delete(request.id);
  }
}

async function runSearch(query: string): Promise<void> {
  const started = beginBrowse(browseState, query);
  browseState = started.state;
  // A new search means new results; leave the detail page for them.
  if (detailPanel.isOpen()) detailPanel.close();
  setStatus(query ? `Searching for “${query}”…` : 'Loading popular…');
  resetGrid();
  await runBrowse(started.request);
}

async function loadNextPage(): Promise<void> {
  const started = beginNextPage(browseState);
  if (!started) return;
  browseState = started.state;
  setStatus(`Loading page ${started.request.page}…`);
  await runBrowse(started.request);
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
    const requestedSourceId = sourceSelect.value;
    try {
      await api.selectSource(requestedSourceId);
      selectedSourceId = requestedSourceId;
      // Settings belong to the source, so reload them rather than showing stale fields.
      if (currentView === 'settings') await openSettings();
      await runSearch(searchInput.value.trim());
    } catch (error) {
      if (selectedSourceId !== null) sourceSelect.value = selectedSourceId;
      setStatus(describe(error), 'error');
    }
  })();
});

loadMoreButton.addEventListener('click', () => void loadNextPage());

bannerUpdate.addEventListener('click', () => {
  void (async () => {
    bannerUpdate.disabled = true;
    try {
      const state = await api.updateBridge();
      renderBridgeState(state);
      if (state.stage === 'ready' && state.install?.updateAvailable === null) {
        setStatus(
          `Extension bridge updated to ${state.install.version ?? 'the pinned release'}.`,
          'ok',
        );
      } else if (state.message !== null) {
        setStatus(state.message, 'error');
      }
      // The bridge restarted, so the source list is fresh from disk.
      const snapshot = await api.getSnapshot();
      renderSources(snapshot.sources, snapshot.selectedSourceId);
      if (currentView === 'extensions') await extensions.refresh();
    } catch (error) {
      setStatus(describe(error), 'error');
      bannerUpdate.disabled = false;
    }
  })();
});

api.onBridgeState(renderBridgeState);

// The queue changes without this window asking: it advances by itself when an
// episode ends, whether or not anyone is looking at the browser.
api.onQueueState((state) => detailPanel.setQueue(state));
let receivedPlaybackStateEvent = false;
api.onPlaybackState((state) => {
  receivedPlaybackStateEvent = true;
  detailPanel.setPlaybackState(state);
});

void (async () => {
  renderBridgeState({ stage: 'idle', progress: null, message: null, install: null });
  // A queue survives the window being closed and reopened, so start from what
  // the main process already holds rather than from empty.
  void api.getQueue().then(
    (state) => detailPanel.setQueue(state),
    () => {
      // A live queue event can still populate the panel after a failed snapshot request.
    },
  );
  void api.getPlaybackState().then(
    (state) => {
      if (!receivedPlaybackStateEvent) detailPanel.setPlaybackState(state);
    },
    () => {
      // A live playback event can still populate the cue after a failed snapshot request.
    },
  );
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
    renderBridgeState({
      stage: 'failed',
      progress: null,
      message: describe(error),
      install: null,
    });
    setStatus(describe(error), 'error');
  }
})();
