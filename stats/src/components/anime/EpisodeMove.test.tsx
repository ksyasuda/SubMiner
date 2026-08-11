import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { AnimeEpisode, AnimeLibraryItem, StatsMoveVideoResponse } from '../../types/stats';
import { EpisodeList } from './EpisodeList';

interface TestWindow extends Window {
  IS_REACT_ACT_ENVIRONMENT?: boolean;
}

function installDom(): () => void {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const previousHTMLElement = globalThis.HTMLElement;
  const previousIsReactActEnvironment = (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT;
  const window = new Window() as TestWindow;

  Object.defineProperty(globalThis, 'window', { value: window, configurable: true });
  Object.defineProperty(globalThis, 'document', { value: window.document, configurable: true });
  Object.defineProperty(globalThis, 'HTMLElement', {
    value: window.HTMLElement,
    configurable: true,
  });
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;

  return () => {
    Object.defineProperty(globalThis, 'window', { value: previousWindow, configurable: true });
    Object.defineProperty(globalThis, 'document', { value: previousDocument, configurable: true });
    Object.defineProperty(globalThis, 'HTMLElement', {
      value: previousHTMLElement,
      configurable: true,
    });
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = previousIsReactActEnvironment;
  };
}

function episode(videoId: number, title: string): AnimeEpisode {
  return {
    videoId,
    episode: videoId,
    season: null,
    durationMs: 1_440_000,
    endedMediaMs: null,
    watched: 0,
    canonicalTitle: title,
    totalSessions: 1,
    totalActiveMs: 1000,
    totalCards: 0,
    totalTokensSeen: 0,
    totalYomitanLookupCount: 0,
    lastWatchedMs: 1,
  };
}

function libraryItem(animeId: number, title: string): AnimeLibraryItem {
  return {
    animeId,
    canonicalTitle: title,
    anilistId: null,
    totalSessions: 1,
    totalActiveMs: 1000,
    totalCards: 0,
    totalTokensSeen: 0,
    episodeCount: 1,
    episodesTotal: null,
    lastWatchedMs: 1,
  };
}

function findButtonByTitle(container: Element, title: string): HTMLElement {
  const match = [...container.querySelectorAll('button')].find(
    (button) => button.getAttribute('title') === title,
  );
  assert.ok(match, `expected a button titled "${title}"`);
  return match as unknown as HTMLElement;
}

function findButtonByText(container: Element, text: string): HTMLElement {
  const match = [...container.querySelectorAll('button')].find((button) =>
    (button.textContent ?? '').includes(text),
  );
  assert.ok(match, `expected a "${text}" button`);
  return match as unknown as HTMLElement;
}

test('EpisodeList moves an episode to the library entry picked in the dialog', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    moveVideoToAnime: apiClient.moveVideoToAnime,
  };

  let moveCall: { videoId: number; animeId: number } | null = null;
  let movedResult: boolean | null = null;

  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Current Entry'),
    libraryItem(2, 'Real Series'),
  ]) as typeof apiClient.getAnimeLibrary;
  apiClient.moveVideoToAnime = (async (videoId: number, animeId: number) => {
    moveCall = { videoId, animeId };
    return {
      ok: true,
      animeId,
      previousAnimeId: 1,
      removedPreviousAnime: true,
    } satisfies StatsMoveVideoResponse;
  }) as typeof apiClient.moveVideoToAnime;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        <EpisodeList
          episodes={[episode(5, 'Stray Episode')]}
          animeId={1}
          onEpisodeMoved={(removedPreviousAnime) => {
            movedResult = removedPreviousAnime;
          }}
        />,
      );
    });

    await act(async () => {
      findButtonByTitle(container, 'Move to another library entry').click();
    });
    assert.match(container.textContent ?? '', /Move "Stray Episode" To/);
    // The entry the episode already belongs to is not offered as a target.
    assert.doesNotMatch(container.textContent ?? '', /Current Entry/);

    await act(async () => {
      findButtonByText(container, 'Real Series').click();
    });

    assert.deepEqual(moveCall, { videoId: 5, animeId: 2 });
    assert.equal(movedResult, true);
    // The row leaves this entry's list and the picker closes.
    assert.doesNotMatch(container.textContent ?? '', /Stray Episode/);

    await act(async () => {
      root.unmount();
    });
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});
