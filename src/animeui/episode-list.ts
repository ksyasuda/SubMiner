import { capture, LatestRequest, safeUploadDate } from './browse-state';
import { describe, el } from './dom';
import { closeContextMenu, showContextMenu, type ContextMenuItem } from './context-menu';
import { filterEpisodes } from './episode-filter';
import { describeMarkCount, episodesInScope } from './episode-marks';
import { describeQueuePosition } from './episode-queue';
import { createEpisodeQueueControls } from './episode-queue-controls';
import type { AnimeBrowserAPI, AnimeBrowserEpisode } from '../types/anime-browser';

export interface SelectedAnime {
  url: string;
  title: string;
  sourceId: string;
}

interface EpisodeListOptions {
  api: AnimeBrowserAPI;
  setStatus: (message: string, tone?: 'info' | 'ok' | 'error') => void;
  /** The anime the list belongs to, or null once the detail page is closed. */
  selectedAnime: () => SelectedAnime | null;
}

/**
 * An episode, the number the filter matches on, and the number the rail shows.
 *
 * They are separate on purpose: the rail falls back to the listing position so
 * every row has an index, but a special the source gave no number to must not
 * become findable under a number it does not have.
 */
interface ListedEpisode {
  episode: AnimeBrowserEpisode;
  number: number | null;
  displayIndex: number;
  name: string;
}

export function createEpisodeList({ api, setStatus, selectedAnime }: EpisodeListOptions) {
  const episodes = el<HTMLOListElement>('episodes');
  const episodesCount = el<HTMLSpanElement>('episodes-count');
  const episodesWatched = el<HTMLSpanElement>('episodes-watched');
  const filterInput = el<HTMLInputElement>('episodes-filter');

  let listed: ListedEpisode[] = [];
  /** Episode urls the stats history reports as already watched. */
  let watched = new Set<string>();
  /**
   * Which episode is resolving or playing. Kept here rather than only on the
   * button, so filtering mid-playback repaints the cue instead of dropping it.
   */
  let cueState: { url: string; state: 'loading' | 'playing' } | null = null;
  const watchStateRequests = new LatestRequest();
  /**
   * Mark writes carry their own token: a background refresh starting mid-write
   * must not make the write look superseded and drop its repaint on the floor.
   */
  const markWrites = new LatestRequest();
  const playbacks = new LatestRequest();

  function formatEpisodeIndex(item: ListedEpisode): string {
    const value = item.number ?? item.displayIndex;
    return Number.isInteger(value) ? String(value).padStart(2, '0') : value.toFixed(1);
  }

  /** Push `cueState` onto whichever rows are on screen right now. */
  function applyCueState(): void {
    for (const row of episodes.querySelectorAll<HTMLButtonElement>('.cue')) {
      if (cueState && row.dataset.episodeUrl === cueState.url) row.dataset.state = cueState.state;
      else row.removeAttribute('data-state');
    }
  }

  const queue = createEpisodeQueueControls({
    api,
    setStatus,
    selectedAnime,
    playEpisode: (episode) => playEpisode(episode),
    onChange: () => paint(),
    onAdvance: (entry) => {
      // The queue started this one, so the cue this window was holding belongs
      // to an episode that has finished. The new one only earns the cue when it
      // is an episode of the anime on screen.
      const anime = selectedAnime();
      const mine =
        anime !== null && anime.sourceId === entry.sourceId && anime.url === entry.animeUrl;
      cueState = mine ? { url: entry.episodeUrl, state: 'playing' } : null;
      applyCueState();
      setStatus(`Queue started ${entry.episodeName}.`, 'ok');
    },
  });

  async function playEpisode(episode: AnimeBrowserEpisode) {
    const anime = selectedAnime();
    if (!anime) return;

    // Only the newest click owns the cue states and the status line; an earlier
    // episode resolving late must not overwrite them.
    const playback = playbacks.begin();
    cueState = { url: episode.url, state: 'loading' };
    applyCueState();
    setStatus(`Resolving ${episode.name}…`);

    const attempt = await capture(() =>
      api.playEpisode({
        sourceId: anime.sourceId,
        animeUrl: anime.url,
        animeTitle: anime.title,
        episodeUrl: episode.url,
        episodeName: episode.name,
        episodeNumber: episode.number,
      }),
    );

    if (!playbacks.isCurrent(playback)) return;

    if (!attempt.ok) {
      cueState = null;
      applyCueState();
      setStatus(describe(attempt.error), 'error');
      return;
    }

    const result = attempt.value;
    if (result.ok) {
      cueState = { url: episode.url, state: 'playing' };
      applyCueState();
      setStatus(
        result.quality ? `Playing ${episode.name} · ${result.quality}` : `Playing ${episode.name}`,
        'ok',
      );
    } else {
      cueState = null;
      applyCueState();
      setStatus(result.error ?? 'Could not play that episode.', 'error');
    }
  }

  function createRow(item: ListedEpisode): HTMLLIElement {
    const { episode } = item;
    const row = document.createElement('li');
    row.className = 'cue-row';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'cue';
    button.dataset.episodeUrl = episode.url;
    if (watched.has(episode.url)) button.dataset.watched = 'true';
    if (cueState?.url === episode.url) button.dataset.state = cueState.state;

    const cueIndex = document.createElement('span');
    cueIndex.className = 'cue-index';
    cueIndex.textContent = formatEpisodeIndex(item);

    const name = document.createElement('span');
    name.className = 'cue-name';
    name.textContent = episode.name;
    // Inline after the title rather than off at the row's far edge, where it
    // reads as belonging to no episode in particular.
    if (watched.has(episode.url)) {
      const mark = document.createElement('span');
      mark.className = 'cue-watched';
      mark.textContent = '✓ watched';
      name.append(mark);
    }
    const position = queue.positionOf(episode);
    if (position !== undefined) {
      button.dataset.queued = 'true';
      // On the row too, so the actions can stay visible without asking the CSS
      // to look inside for a queued cue.
      row.dataset.queued = 'true';
      const mark = document.createElement('span');
      mark.className = 'cue-queued';
      mark.textContent = describeQueuePosition(position);
      name.append(mark);
    }
    if (episode.uploadedAt !== null) {
      const uploaded = safeUploadDate(episode.uploadedAt);
      if (uploaded) {
        const sub = document.createElement('span');
        sub.className = 'cue-sub';
        sub.textContent = uploaded;
        name.append(sub);
      }
    }

    button.append(cueIndex, name);
    button.addEventListener('click', () => void playEpisode(episode));
    button.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      openRowMenu(event, item);
    });

    // The row itself plays, which leaves the other choice — after this one, not
    // instead of it — with nothing to click. These two spell both out, and a
    // right-click on the row says the same thing in words.
    const actions = document.createElement('div');
    actions.className = 'cue-actions';
    actions.append(
      createRowAction('Play', `Play ${episode.name} now`, () => void playEpisode(episode)),
      createRowAction(
        position === undefined ? 'Queue' : 'Queued',
        position === undefined
          ? `Play ${episode.name} after the current episode`
          : `Take ${episode.name} out of the queue`,
        () => void queue.toggle(episode),
        position !== undefined,
      ),
    );

    row.append(button, actions);
    return row;
  }

  function createRowAction(
    label: string,
    title: string,
    onClick: () => void,
    active = false,
  ): HTMLButtonElement {
    const action = document.createElement('button');
    action.type = 'button';
    action.className = 'cue-action';
    action.textContent = label;
    action.title = title;
    action.setAttribute('aria-label', title);
    if (active) action.dataset.active = 'true';
    action.addEventListener('click', onClick);
    return action;
  }

  /**
   * The right-click menu on one episode: this episode, or this one and every
   * episode listed below it, which for a newest-first source is everything
   * older.
   */
  function openRowMenu(event: MouseEvent, item: ListedEpisode): void {
    const index = listed.indexOf(item);
    if (index < 0) return;
    const isWatched = watched.has(item.episode.url);
    const below = episodesInScope(listed, index, 'below');
    const items: ContextMenuItem[] = [
      {
        label: 'Play now',
        onSelect: () => void playEpisode(item.episode),
      },
      {
        label: queue.isQueued(item.episode) ? 'Remove from queue' : 'Play after current',
        onSelect: () => void queue.toggle(item.episode),
      },
      {
        label: isWatched ? 'Mark unwatched' : 'Mark watched',
        separated: true,
        onSelect: () => void applyMark([item], !isWatched),
      },
    ];

    // The oldest episode has nothing below it, so the span entries would only
    // repeat the single one above them.
    if (below.length > 1) {
      const count = below.length - 1;
      items.push(
        {
          label: `Mark this and ${count} below watched`,
          separated: true,
          onSelect: () => void applyMark(below, true),
        },
        {
          label: `Mark this and ${count} below unwatched`,
          onSelect: () => void applyMark(below, false),
        },
      );
    }

    showContextMenu(event.clientX, event.clientY, items);
  }

  /**
   * Write the mark, then repaint from what the store reports rather than from
   * what was asked for: an episode the write could not record must not show a
   * mark that is not there.
   */
  async function applyMark(items: ListedEpisode[], mark: boolean): Promise<void> {
    const anime = selectedAnime();
    if (!anime || items.length === 0) return;

    const request = markWrites.begin();
    const attempt = await capture(() =>
      api.setWatched({
        sourceId: anime.sourceId,
        animeUrl: anime.url,
        animeTitle: anime.title,
        watched: mark,
        episodes: items.map((item) => ({
          episodeUrl: item.episode.url,
          episodeName: item.episode.name,
          episodeNumber: item.episode.number,
        })),
      }),
    );
    if (!markWrites.isCurrent(request)) return;

    if (!attempt.ok) {
      setStatus(describe(attempt.error), 'error');
      return;
    }

    const marked = new Set(
      attempt.value.filter((state) => state.watched).map((state) => state.episodeUrl),
    );
    // Counted from what came back, not from what was asked for, so the status
    // line cannot claim more than the store actually recorded.
    let changed = 0;
    for (const item of items) {
      const isMarked = marked.has(item.episode.url);
      if (isMarked) watched.add(item.episode.url);
      else watched.delete(item.episode.url);
      if (isMarked === mark) changed += 1;
    }
    paint();
    // A refresh that read the store before this write landed would repaint stale
    // marks over the ones above, so re-read from the store to settle the list.
    void refreshWatchState();

    // Nothing came back marked when marking is what was asked for: the write
    // had nowhere to land, which is what a disabled stats history looks like.
    if (mark && marked.size === 0) {
      setStatus('Watch marks need immersion tracking enabled.', 'error');
      return;
    }
    setStatus(describeMarkCount(changed, mark), 'ok');
  }

  /** Repaints the rows the current filter leaves, and the two counters. */
  function paint(): void {
    const query = filterInput.value;
    const visible = filterEpisodes(listed, query);
    episodes.replaceChildren(...visible.map(createRow));

    if (listed.length === 0) {
      episodesCount.textContent = '';
    } else {
      episodesCount.textContent =
        visible.length === listed.length
          ? `${listed.length}`
          : `${visible.length} of ${listed.length}`;
    }

    const watchedCount = listed.filter((item) => watched.has(item.episode.url)).length;
    episodesWatched.textContent = watchedCount > 0 ? `${watchedCount} watched` : '';
    episodesWatched.classList.toggle('hidden', watchedCount === 0);
    filterInput.classList.toggle('hidden', listed.length === 0);
  }

  /**
   * Ask the stats history which of these episodes are already watched.
   *
   * Playback marks an episode watched partway through a session, so this also
   * runs when the window comes back to the front: an episode finished in mpv
   * while the browser sat behind it shows its mark on the way back.
   */
  async function refreshWatchState(): Promise<void> {
    const anime = selectedAnime();
    if (!anime || listed.length === 0) return;

    const request = watchStateRequests.begin();
    const attempt = await capture(() =>
      api.getWatchState({
        sourceId: anime.sourceId,
        animeUrl: anime.url,
        episodeUrls: listed.map((item) => item.episode.url),
      }),
    );
    if (!watchStateRequests.isCurrent(request)) return;
    // Watch marks are decoration; a failed lookup leaves the list as it is.
    if (!attempt.ok) return;

    watched = new Set(
      attempt.value.filter((state) => state.watched).map((state) => state.episodeUrl),
    );
    paint();
  }

  function render(list: AnimeBrowserEpisode[]): void {
    listed = list.map((episode, index) => ({
      episode,
      number: episode.number,
      displayIndex: list.length - index,
      name: episode.name,
    }));
    paint();
    void refreshWatchState();
  }

  function clear(): void {
    // A menu opened against the list that is going away has nothing left to act on.
    closeContextMenu();
    watchStateRequests.cancel();
    markWrites.cancel();
    playbacks.cancel();
    listed = [];
    watched = new Set();
    cueState = null;
    filterInput.value = '';
    paint();
  }

  filterInput.addEventListener('input', paint);
  // Escape inside the filter clears it instead of closing the detail page.
  filterInput.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape' || filterInput.value === '') return;
    event.stopPropagation();
    filterInput.value = '';
    paint();
  });
  window.addEventListener('focus', () => void refreshWatchState());

  return { render, clear, refreshWatchState, setQueue: queue.setState };
}
