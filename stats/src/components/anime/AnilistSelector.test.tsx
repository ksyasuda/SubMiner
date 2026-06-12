import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import { AnilistSelector } from './AnilistSelector';

interface TestWindow extends Window {
  IS_REACT_ACT_ENVIRONMENT?: boolean;
}

function installDom(): () => void {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const previousHTMLElement = globalThis.HTMLElement;
  const previousISReactActEnvironment = (
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
    Object.defineProperty(globalThis, 'window', {
      value: previousWindow,
      configurable: true,
    });
    Object.defineProperty(globalThis, 'document', {
      value: previousDocument,
      configurable: true,
    });
    Object.defineProperty(globalThis, 'HTMLElement', {
      value: previousHTMLElement,
      configurable: true,
    });
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = previousISReactActEnvironment;
  };
}

function renderSelector(root: Root, props: { animeId: number; initialQuery: string }): void {
  root.render(
    <AnilistSelector
      animeId={props.animeId}
      initialQuery={props.initialQuery}
      onClose={() => {}}
      onLinked={() => {}}
    />,
  );
}

function inputValue(container: Element): string {
  const input = container.querySelector('input');
  assert.ok(input);
  return input.value;
}

function deferred<T>(): {
  promise: Promise<T>;
  resolve: (value: T) => void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

test('AnilistSelector resyncs normalized query and searches when the initial anime changes', async () => {
  const uninstallDom = installDom();
  const originalSearchAnilist = apiClient.searchAnilist;
  const secondSearch = deferred<Awaited<ReturnType<typeof apiClient.searchAnilist>>>();
  const searchCalls: string[] = [];

  apiClient.searchAnilist = (async (query: string) => {
    searchCalls.push(query);
    if (query === 'My Hero Academia') {
      return secondSearch.promise;
    }
    return [
      {
        id: 1,
        episodes: 1,
        season: null,
        seasonYear: null,
        description: null,
        coverImage: null,
        title: { romaji: 'First Result', english: null, native: null },
      },
    ];
  }) as typeof apiClient.searchAnilist;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      renderSelector(root, { animeId: 1, initialQuery: 'Sword Art Online Season 1' });
    });

    assert.equal(inputValue(container), 'Sword Art Online');
    assert.deepEqual(searchCalls, ['Sword Art Online']);
    assert.match(container.textContent ?? '', /First Result/);

    await act(async () => {
      renderSelector(root, { animeId: 2, initialQuery: 'My Hero Academia: Season 3' });
    });

    assert.equal(inputValue(container), 'My Hero Academia');
    assert.deepEqual(searchCalls, ['Sword Art Online', 'My Hero Academia']);
    assert.doesNotMatch(container.textContent ?? '', /First Result/);
    assert.match(container.textContent ?? '', /Searching/);

    await act(async () => {
      secondSearch.resolve([
        {
          id: 2,
          episodes: 2,
          season: null,
          seasonYear: null,
          description: null,
          coverImage: null,
          title: { romaji: 'Second Result', english: null, native: null },
        },
      ]);
      await secondSearch.promise;
    });

    assert.match(container.textContent ?? '', /Second Result/);

    await act(async () => {
      root.unmount();
    });
  } finally {
    apiClient.searchAnilist = originalSearchAnilist;
    uninstallDom();
  }
});
