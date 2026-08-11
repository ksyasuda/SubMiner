import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { AnimeLibraryItem, StatsMergeAnimeResponse } from '../../types/stats';
import { AnimeTab } from './AnimeTab';

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

function libraryItem(animeId: number, title: string, episodeCount: number): AnimeLibraryItem {
  return {
    animeId,
    canonicalTitle: title,
    anilistId: null,
    totalSessions: 1,
    totalActiveMs: 1000,
    totalCards: 1,
    totalTokensSeen: 0,
    episodeCount,
    episodesTotal: null,
    lastWatchedMs: animeId,
  };
}

function findButton(container: Element, label: string): HTMLElement {
  const match = [...container.querySelectorAll('button')].find((button) =>
    (button.textContent ?? '').includes(label),
  );
  assert.ok(match, `expected a "${label}" button`);
  return match as unknown as HTMLElement;
}

/** Library cards only expose aria-pressed while selection mode is on. */
function cardButtons(container: Element): HTMLButtonElement[] {
  return [...container.querySelectorAll('button[aria-pressed]')] as unknown as HTMLButtonElement[];
}

function mergeButton(container: Element): HTMLButtonElement {
  const match = [...container.querySelectorAll('button')].find(
    (button) => (button.textContent ?? '').trim() === 'Merge Selected',
  );
  assert.ok(match, 'expected a "Merge Selected" button');
  return match as unknown as HTMLButtonElement;
}

test('AnimeTab merges the selected duplicate entries into the chosen keeper', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    mergeAnime: apiClient.mergeAnime,
  };

  // Two cards for one show, the split this feature exists to undo.
  let entries = [libraryItem(1, 'Show', 2), libraryItem(2, 'Show Season 1', 1)];
  let libraryFetches = 0;
  let mergeCall: { targetAnimeId: number; sourceAnimeIds: number[] } | null = null;

  apiClient.getAnimeLibrary = (async () => {
    libraryFetches += 1;
    return entries;
  }) as typeof apiClient.getAnimeLibrary;
  apiClient.mergeAnime = (async (targetAnimeId: number, sourceAnimeIds: number[]) => {
    mergeCall = { targetAnimeId, sourceAnimeIds };
    entries = [libraryItem(1, 'Show', 3)];
    return {
      ok: true,
      animeId: targetAnimeId,
      mergedAnimeIds: sourceAnimeIds,
      movedVideos: 1,
    } satisfies StatsMergeAnimeResponse;
  }) as typeof apiClient.mergeAnime;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<AnimeTab />);
    });
    assert.equal(libraryFetches, 1);

    await act(async () => {
      findButton(container, 'Select').click();
    });
    // Nothing to merge until at least two entries are picked.
    assert.equal(mergeButton(container).disabled, true);

    // Sorted by last watched, so the season-tagged duplicate comes first.
    const cards = cardButtons(container);
    assert.equal(cards.length, 2);
    assert.match(cards[0]?.textContent ?? '', /Show Season 1/);

    await act(async () => {
      cards[0]?.click();
    });
    assert.equal(mergeButton(container).disabled, true);

    await act(async () => {
      cardButtons(container)[1]?.click();
    });
    assert.equal(mergeButton(container).disabled, false);

    await act(async () => {
      mergeButton(container).click();
    });
    // The dialog defaults to the entry with the most episodes.
    assert.match(container.textContent ?? '', /Merge 2 Library Entries/);

    await act(async () => {
      findButton(container, 'Merge Entries').click();
    });

    assert.deepEqual(mergeCall, { targetAnimeId: 1, sourceAnimeIds: [2] });
    assert.equal(libraryFetches, 2);
    // Selection mode closes and the grid is back to a single card.
    assert.doesNotMatch(container.textContent ?? '', /Merge 2 Library Entries/);
    assert.doesNotMatch(container.textContent ?? '', /Show Season 1/);

    await act(async () => {
      root.unmount();
    });
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});
