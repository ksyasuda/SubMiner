import { LatestRequest } from './browse-state';
import { describe, el } from './dom';
import { createEpisodeList, type SelectedAnime } from './episode-list';
import type {
  AnimeBrowserAPI,
  AnimeBrowserEntry,
  AnimeBrowserPlaybackState,
} from '../types/anime-browser';

interface DetailPanelOptions {
  api: AnimeBrowserAPI;
  setStatus: (message: string, tone?: 'info' | 'ok' | 'error') => void;
}

export function createDetailPanel({ api, setStatus }: DetailPanelOptions) {
  const results = el<HTMLElement>('results');
  const detail = el<HTMLElement>('detail');
  const detailBack = el<HTMLButtonElement>('detail-back');
  const detailCover = el<HTMLImageElement>('detail-cover');
  const detailTitle = el<HTMLHeadingElement>('detail-title');
  const detailChips = el<HTMLDivElement>('detail-chips');
  const detailDescription = el<HTMLParagraphElement>('detail-description');

  let selectedAnime: SelectedAnime | null = null;
  let resultsScrollTop = 0;
  const requests = new LatestRequest();
  const episodeList = createEpisodeList({
    api,
    setStatus,
    selectedAnime: () => selectedAnime,
  });

  async function open(entry: AnimeBrowserEntry): Promise<void> {
    const request = requests.begin();
    selectedAnime = { url: entry.url, title: entry.title, sourceId: entry.sourceId };
    resultsScrollTop = results.scrollTop;
    results.classList.add('hidden');
    detail.classList.remove('hidden');
    detail.scrollTop = 0;
    detailTitle.textContent = entry.title;
    detailDescription.textContent = 'Loading…';
    detailChips.replaceChildren();
    episodeList.clear();
    detailCover.src = entry.thumbnailUrl ?? '';
    setStatus(`Loading ${entry.title}…`);

    // Details and episodes are independent bridge calls. A source whose
    // details call fails (a bridge that rejects the extension's metadata, a
    // flaky page) can still list episodes, so a details failure only costs the
    // description and chips, not the episode list.
    const [detailsResult, episodesResult] = await Promise.allSettled([
      api.getDetails(entry.url, entry.sourceId),
      api.getEpisodes(entry.url, entry.sourceId),
    ]);
    if (!requests.isCurrent(request)) return;

    if (episodesResult.status === 'rejected') {
      detailDescription.textContent = '';
      setStatus(describe(episodesResult.reason), 'error');
      return;
    }
    const episodes = episodesResult.value;

    const chips: HTMLSpanElement[] = [];
    const source = document.createElement('span');
    source.className = 'chip source';
    source.textContent = entry.sourceName;
    chips.push(source);

    if (detailsResult.status === 'rejected') {
      detailDescription.textContent = 'Details unavailable from this source.';
      detailChips.replaceChildren(...chips);
      episodeList.render(episodes);
      setStatus(describe(detailsResult.reason), 'error');
      return;
    }

    const details = detailsResult.value;
    detailTitle.textContent = details.title;
    detailDescription.textContent = details.description ?? 'No description from this source.';
    if (details.thumbnailUrl) detailCover.src = details.thumbnailUrl;

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

    episodeList.render(episodes);
    setStatus(`${details.title} · ${episodes.length} episodes`);
  }

  function close(): void {
    requests.cancel();
    episodeList.clear();
    detail.classList.add('hidden');
    results.classList.remove('hidden');
    results.scrollTop = resultsScrollTop;
    selectedAnime = null;
  }

  detailBack.addEventListener('click', close);
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !detail.classList.contains('hidden')) {
      event.preventDefault();
      event.stopImmediatePropagation();
      close();
    }
  });

  return {
    open,
    close,
    isOpen: (): boolean => !detail.classList.contains('hidden'),
    /** The queue outlives the open anime, so it is pushed in from outside. */
    setQueue: episodeList.setQueue,
    setPlaybackState: (state: AnimeBrowserPlaybackState | null) =>
      episodeList.setPlaybackState(state),
  };
}
