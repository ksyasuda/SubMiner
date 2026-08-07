import { capture } from './browse-state';
import { describe, el } from './dom';
import { describeQueuePosition, queueKey, queuePositions } from './episode-queue';
import type { SelectedAnime } from './episode-list';
import type {
  AnimeBrowserAPI,
  AnimeBrowserEpisode,
  AnimeBrowserQueueEntry,
  AnimeBrowserQueueState,
} from '../types/anime-browser';

interface EpisodeQueueControlsOptions {
  api: AnimeBrowserAPI;
  setStatus: (message: string, tone?: 'info' | 'ok' | 'error') => void;
  /** The anime the open detail page belongs to, or null once it is closed. */
  selectedAnime: () => SelectedAnime | null;
  /** Play now, with the cue states and status line the episode list owns. */
  playEpisode: (episode: AnimeBrowserEpisode) => Promise<void>;
  /** Repaint the rows; their badges and buttons read from this queue. */
  onChange: () => void;
  /** The queue started this episode by itself, so nothing else is playing now. */
  onAdvance: (entry: AnimeBrowserQueueEntry) => void;
}

/**
 * The play queue as this window sees it.
 *
 * The queue itself lives in the main process — it has to, since it advances
 * when an episode ends whether or not this window is open. Everything here is a
 * view of what it pushes back, plus the two calls that change it.
 */
export function createEpisodeQueueControls(options: EpisodeQueueControlsOptions) {
  const { api, setStatus, selectedAnime, playEpisode, onChange, onAdvance } = options;
  const queuedLabel = el<HTMLSpanElement>('episodes-queued');
  const clearButton = el<HTMLButtonElement>('episodes-queue-clear');

  let queue: AnimeBrowserQueueState = {
    entries: [],
    lastError: null,
    advances: 0,
    lastStarted: null,
  };
  let positions = new Map<string, number>();

  function positionOf(episode: AnimeBrowserEpisode): number | undefined {
    const anime = selectedAnime();
    return anime ? positions.get(queueKey(anime.sourceId, episode.url)) : undefined;
  }

  function isQueued(episode: AnimeBrowserEpisode): boolean {
    return positionOf(episode) !== undefined;
  }

  /** The queue as the main process now holds it, from a call or from a push. */
  function setState(next: AnimeBrowserQueueState): void {
    const previousError = queue.lastError;
    const advanced = next.advances > queue.advances;
    queue = next;
    positions = queuePositions(next.entries);
    paintHeader();
    onChange();
    // The episode this window had marked playing has ended, whether or not the
    // one that replaced it is in the list on screen.
    if (advanced && next.lastStarted) onAdvance(next.lastStarted);
    // An advance that failed happened with nobody watching this window, so it
    // is reported once rather than on every repaint after it.
    if (next.lastError && next.lastError !== previousError) {
      setStatus(`Queue stopped — ${next.lastError}`, 'error');
    }
  }

  /**
   * The queue spans every anime, so it is counted whole even when none of its
   * episodes are in the list on screen.
   */
  function paintHeader(): void {
    const queued = queue.entries.length;
    queuedLabel.textContent = queued > 0 ? `Queue · ${queued}` : '';
    queuedLabel.classList.toggle('hidden', queued === 0);
    clearButton.classList.toggle('hidden', queued === 0);
  }

  /** Queue the episode, or take it back out when it is already in line. */
  async function toggle(episode: AnimeBrowserEpisode): Promise<void> {
    const anime = selectedAnime();
    if (!anime) return;

    const queued = isQueued(episode);
    // Queueing behind nothing would wait for an end that never comes, so with
    // an idle player the second option collapses into the first.
    if (!queued) {
      const playing = await capture(() => api.isPlaying());
      if (playing.ok && !playing.value) {
        setStatus('Nothing was playing, so this starts now.');
        await playEpisode(episode);
        return;
      }
    }

    const attempt = await capture(() =>
      queued
        ? api.dequeueEpisode(anime.sourceId, episode.url)
        : api.queueEpisode({
            sourceId: anime.sourceId,
            animeUrl: anime.url,
            animeTitle: anime.title,
            episodeUrl: episode.url,
            episodeName: episode.name,
            episodeNumber: episode.number,
          }),
    );
    if (!attempt.ok) {
      setStatus(describe(attempt.error), 'error');
      return;
    }

    // Applied from what the call returned rather than waited for on the push
    // channel, so the row settles on the click that changed it.
    setState(attempt.value);
    const position = positionOf(episode);
    setStatus(
      position === undefined
        ? `${episode.name} left the queue.`
        : `${episode.name} is ${describeQueuePosition(position)}.`,
      'ok',
    );
  }

  async function clear(): Promise<void> {
    const attempt = await capture(() => api.clearQueue());
    if (!attempt.ok) {
      setStatus(describe(attempt.error), 'error');
      return;
    }
    setState(attempt.value);
    setStatus('Queue cleared.', 'ok');
  }

  clearButton.addEventListener('click', () => void clear());

  return { positionOf, isQueued, setState, toggle };
}
