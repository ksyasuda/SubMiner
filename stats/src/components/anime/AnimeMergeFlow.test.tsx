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

test('AnimeTab keeps a suggested duplicate visible until it is reviewed and merged', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
    dismissAnimeMergeRecommendation: apiClient.dismissAnimeMergeRecommendation,
    mergeAnime: apiClient.mergeAnime,
  };

  let entries = [libraryItem(1, 'Show', 2), libraryItem(2, 'Show Season 1', 1)];
  let recommendations = [{ recommendationId: 41, animeIds: [1, 2] }];
  let mergeCall: { targetAnimeId: number; sourceAnimeIds: number[] } | null = null;

  apiClient.getAnimeLibrary = (async () => entries) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => ({
    recommendations,
  })) as typeof apiClient.getAnimeMergeRecommendations;
  apiClient.dismissAnimeMergeRecommendation = (async () =>
    undefined) as typeof apiClient.dismissAnimeMergeRecommendation;
  apiClient.mergeAnime = (async (targetAnimeId: number, sourceAnimeIds: number[]) => {
    mergeCall = { targetAnimeId, sourceAnimeIds };
    entries = [libraryItem(targetAnimeId, 'Show', 3)];
    recommendations = [];
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

    assert.match(container.textContent ?? '', /Possible duplicate/);
    assert.match(container.textContent ?? '', /Show/);
    assert.match(container.textContent ?? '', /Show Season 1/);

    await act(async () => {
      findButton(container, 'Review merge').click();
    });
    assert.match(container.textContent ?? '', /Merge 2 Library Entries/);

    const keeper = [...container.querySelectorAll('button[aria-pressed]')].find((button) =>
      (button.textContent ?? '').includes('Show Season 1'),
    ) as HTMLButtonElement | undefined;
    assert.ok(keeper);
    await act(async () => {
      keeper.click();
    });
    await act(async () => {
      findButton(container, 'Merge Entries').click();
    });

    assert.deepEqual(mergeCall, { targetAnimeId: 2, sourceAnimeIds: [1] });
    assert.doesNotMatch(container.textContent ?? '', /Possible duplicate/);
    assert.doesNotMatch(container.textContent ?? '', /Show Season 1/);

    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab dismisses a false-positive duplicate recommendation', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
    dismissAnimeMergeRecommendation: apiClient.dismissAnimeMergeRecommendation,
  };
  let dismissedId: number | null = null;

  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show', 2),
    libraryItem(2, 'Different Show', 1),
  ]) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => ({
    recommendations: [{ recommendationId: 73, animeIds: [1, 2] }],
  })) as typeof apiClient.getAnimeMergeRecommendations;
  apiClient.dismissAnimeMergeRecommendation = (async (recommendationId: number) => {
    dismissedId = recommendationId;
  }) as typeof apiClient.dismissAnimeMergeRecommendation;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => root.render(<AnimeTab />));

    await act(async () => {
      findButton(container, 'Not duplicates').click();
    });

    assert.equal(dismissedId, 73);
    assert.doesNotMatch(container.textContent ?? '', /Possible duplicate/);
    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab refreshes the library and recommendations when the window regains focus', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
  };
  let entries = [libraryItem(1, 'Show', 2), libraryItem(2, 'Show Season 1', 1)];
  let libraryFetches = 0;
  let recommendationFetches = 0;
  apiClient.getAnimeLibrary = (async () => {
    libraryFetches += 1;
    return entries;
  }) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => {
    recommendationFetches += 1;
    return { recommendations: [] };
  }) as typeof apiClient.getAnimeMergeRecommendations;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => root.render(<AnimeTab />));

    entries = [libraryItem(1, 'Show', 3)];
    await act(async () => {
      window.dispatchEvent(new window.Event('focus'));
    });

    assert.equal(libraryFetches, 2);
    assert.equal(recommendationFetches, 2);
    assert.doesNotMatch(container.textContent ?? '', /Show Season 1/);
    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab keeps a recommendation visible through a transient refresh failure', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
  };
  let failRecommendations = false;
  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show', 2),
    libraryItem(2, 'Show Season 1', 1),
  ]) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => {
    if (failRecommendations) throw new Error('temporary failure');
    return { recommendations: [{ recommendationId: 41, animeIds: [1, 2] }] };
  }) as typeof apiClient.getAnimeMergeRecommendations;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => root.render(<AnimeTab />));
    assert.match(container.textContent ?? '', /Possible duplicate/);

    failRecommendations = true;
    await act(async () => window.dispatchEvent(new window.Event('focus')));

    assert.match(container.textContent ?? '', /Possible duplicate/);
    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab retains a recommendation and reports a failed dismissal', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
    dismissAnimeMergeRecommendation: apiClient.dismissAnimeMergeRecommendation,
  };
  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show', 2),
    libraryItem(2, 'Show Season 1', 1),
  ]) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => ({
    recommendations: [{ recommendationId: 41, animeIds: [1, 2] }],
  })) as typeof apiClient.getAnimeMergeRecommendations;
  apiClient.dismissAnimeMergeRecommendation = (async () => {
    throw new Error('offline');
  }) as typeof apiClient.dismissAnimeMergeRecommendation;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => root.render(<AnimeTab />));

    await act(async () => findButton(container, 'Not duplicates').click());

    assert.match(container.textContent ?? '', /Possible duplicate/);
    assert.match(container.textContent ?? '', /Could not dismiss this suggestion/);
    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});

test('AnimeTab keeps an open recommendation review stable during background refresh', async () => {
  const uninstallDom = installDom();
  const original = {
    getAnimeLibrary: apiClient.getAnimeLibrary,
    getAnimeMergeRecommendations: apiClient.getAnimeMergeRecommendations,
  };
  let recommendations = [{ recommendationId: 41, animeIds: [1, 2] as [number, number] }];
  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show', 2),
    libraryItem(2, 'Show Season 1', 1),
  ]) as typeof apiClient.getAnimeLibrary;
  apiClient.getAnimeMergeRecommendations = (async () => ({
    recommendations,
  })) as typeof apiClient.getAnimeMergeRecommendations;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => root.render(<AnimeTab />));
    await act(async () => findButton(container, 'Review merge').click());
    assert.match(container.textContent ?? '', /Merge 2 Library Entries/);

    recommendations = [];
    await act(async () => window.dispatchEvent(new window.Event('focus')));

    assert.match(container.textContent ?? '', /Merge 2 Library Entries/);
    assert.match(container.textContent ?? '', /Show Season 1/);
    await act(async () => root.unmount());
  } finally {
    Object.assign(apiClient, original);
    uninstallDom();
  }
});
