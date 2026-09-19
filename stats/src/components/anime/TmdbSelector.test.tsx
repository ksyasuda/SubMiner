import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { StatsTmdbSearchResult } from '../../types/stats';
import { TmdbSelector } from './TmdbSelector';

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
    Object.defineProperty(globalThis, 'window', { value: previousWindow, configurable: true });
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

const HANZAWA: StatsTmdbSearchResult = {
  tmdbId: 61222,
  tmdbType: 'tv',
  title: 'Hanzawa Naoki',
  originalTitle: '半沢直樹',
  originalLanguage: 'ja',
  overview: null,
  posterUrl: null,
  year: 2013,
  isAnimation: false,
};

test('TmdbSelector searches the normalized title and links the picked result by id', async () => {
  const uninstallDom = installDom();
  const original = {
    searchTmdb: apiClient.searchTmdb,
    reassignAnimeTmdb: apiClient.reassignAnimeTmdb,
  };
  const searchCalls: string[] = [];
  const linkCalls: Array<[number, { tmdbId: number; tmdbType: string }]> = [];
  let linked = 0;
  apiClient.searchTmdb = (async (query: string) => {
    searchCalls.push(query);
    return [HANZAWA];
  }) as typeof apiClient.searchTmdb;
  apiClient.reassignAnimeTmdb = (async (animeId: number, info) => {
    linkCalls.push([animeId, info]);
  }) as typeof apiClient.reassignAnimeTmdb;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        <TmdbSelector
          animeId={9}
          initialQuery="Hanzawa Naoki Season 2"
          onClose={() => {}}
          onLinked={() => {
            linked += 1;
          }}
        />,
      );
    });

    assert.deepEqual(searchCalls, ['Hanzawa Naoki']);
    assert.match(container.textContent ?? '', /半沢直樹/);
    assert.match(container.textContent ?? '', /TV · 2013/);

    const pick = [...container.querySelectorAll('button')].find((button) =>
      /Select/.test(button.textContent ?? ''),
    );
    assert.ok(pick);
    await act(async () => {
      pick.click();
    });

    assert.deepEqual(linkCalls, [[9, { tmdbId: 61222, tmdbType: 'tv' }]]);
    assert.equal(linked, 1);

    await act(async () => {
      root.unmount();
    });
  } finally {
    apiClient.searchTmdb = original.searchTmdb;
    apiClient.reassignAnimeTmdb = original.reassignAnimeTmdb;
    uninstallDom();
  }
});

test('TmdbSelector explains a missing API key instead of showing "No results"', async () => {
  const uninstallDom = installDom();
  const originalSearch = apiClient.searchTmdb;
  apiClient.searchTmdb = (async () => {
    throw new Error('Stats API error: 503 {"error":"TMDB API key not configured."}');
  }) as typeof apiClient.searchTmdb;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        <TmdbSelector
          animeId={1}
          initialQuery="Hanzawa Naoki"
          onClose={() => {}}
          onLinked={() => {}}
        />,
      );
    });

    assert.match(container.textContent ?? '', /tmdb\.apiKey/);
    assert.doesNotMatch(container.textContent ?? '', /No results/);

    await act(async () => {
      root.unmount();
    });
  } finally {
    apiClient.searchTmdb = originalSearch;
    uninstallDom();
  }
});
