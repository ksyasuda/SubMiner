import { capture, LatestRequest, safeUploadDate } from './browse-state';
import { describe, el } from './dom';
import type {
  AnimeBrowserAPI,
  AnimeBrowserEntry,
  AnimeBrowserEpisode,
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
  const episodes = el<HTMLOListElement>('episodes');
  const episodesCount = el<HTMLSpanElement>('episodes-count');

  let selectedAnime: { url: string; title: string; sourceId: string } | null = null;
  let resultsScrollTop = 0;
  const requests = new LatestRequest();

  function formatEpisodeIndex(episode: AnimeBrowserEpisode, fallbackIndex: number): string {
    const value = episode.number ?? fallbackIndex;
    return Number.isInteger(value) ? String(value).padStart(2, '0') : value.toFixed(1);
  }

  async function playEpisode(
    button: HTMLButtonElement,
    episode: AnimeBrowserEpisode,
  ): Promise<void> {
    const anime = selectedAnime;
    if (!anime) return;

    for (const other of episodes.querySelectorAll<HTMLButtonElement>('.cue')) {
      other.removeAttribute('data-state');
    }
    button.dataset.state = 'loading';
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

    if (!attempt.ok) {
      button.removeAttribute('data-state');
      setStatus(describe(attempt.error), 'error');
      return;
    }

    const result = attempt.value;
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
          const uploaded = safeUploadDate(episode.uploadedAt);
          if (uploaded) {
            const sub = document.createElement('span');
            sub.className = 'cue-sub';
            sub.textContent = uploaded;
            name.append(sub);
          }
        }

        button.append(cueIndex, name);
        button.addEventListener('click', () => void playEpisode(button, episode));
        item.append(button);
        return item;
      }),
    );
  }

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
    episodes.replaceChildren();
    episodesCount.textContent = '';
    detailCover.src = entry.thumbnailUrl ?? '';

    try {
      const [details, episodeList] = await Promise.all([
        api.getDetails(entry.url, entry.sourceId),
        api.getEpisodes(entry.url, entry.sourceId),
      ]);
      if (!requests.isCurrent(request)) return;

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
      if (!requests.isCurrent(request)) return;
      detailDescription.textContent = '';
      setStatus(describe(error), 'error');
    }
  }

  function close(): void {
    requests.cancel();
    detail.classList.add('hidden');
    results.classList.remove('hidden');
    results.scrollTop = resultsScrollTop;
    selectedAnime = null;
  }

  detailBack.addEventListener('click', close);
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !detail.classList.contains('hidden')) close();
  });

  return {
    open,
    close,
    isOpen: (): boolean => !detail.classList.contains('hidden'),
  };
}
