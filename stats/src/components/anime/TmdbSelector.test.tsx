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

test('TmdbSelector reports a failed link as a link problem, not a search failure', async () => {
  const uninstallDom = installDom();
  const original = {
    searchTmdb: apiClient.searchTmdb,
    reassignAnimeTmdb: apiClient.reassignAnimeTmdb,
  };
  apiClient.searchTmdb = (async () => [HANZAWA]) as typeof apiClient.searchTmdb;
  apiClient.reassignAnimeTmdb = (async () => {
    throw new Error('Stats API error: 404');
  }) as typeof apiClient.reassignAnimeTmdb;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        <TmdbSelector
          animeId={9}
          initialQuery="Hanzawa Naoki"
          onClose={() => {}}
          onLinked={() => {}}
        />,
      );
    });

    const pick = [...container.querySelectorAll('button')].find((button) =>
      /Select/.test(button.textContent ?? ''),
    );
    assert.ok(pick);
    await act(async () => {
      pick.click();
    });

    assert.match(container.textContent ?? '', /TMDB has no details for this title/);
    assert.doesNotMatch(container.textContent ?? '', /search failed/);
    // The results stay on screen so the user can pick another one.
    assert.match(container.textContent ?? '', /半沢直樹/);

    await act(async () => {
      root.unmount();
    });
  } finally {
    apiClient.searchTmdb = original.searchTmdb;
    apiClient.reassignAnimeTmdb = original.reassignAnimeTmdb;
    uninstallDom();
  }
});

test('TmdbSelector cannot be dismissed while a link is in flight', async () => {
  const uninstallDom = installDom();
  const original = {
    searchTmdb: apiClient.searchTmdb,
    reassignAnimeTmdb: apiClient.reassignAnimeTmdb,
  };
  let finishLink: () => void = () => {};
  let closed = 0;
  apiClient.searchTmdb = (async () => [HANZAWA]) as typeof apiClient.searchTmdb;
  apiClient.reassignAnimeTmdb = (() =>
    new Promise<void>((resolve) => {
      finishLink = resolve;
    })) as typeof apiClient.reassignAnimeTmdb;

  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        <TmdbSelector
          animeId={9}
          initialQuery="Hanzawa Naoki"
          onClose={() => {
            closed += 1;
          }}
          onLinked={() => {}}
        />,
      );
    });

    const pick = [...container.querySelectorAll('button')].find((button) =>
      /Select/.test(button.textContent ?? ''),
    );
    assert.ok(pick);
    await act(async () => {
      pick.click();
    });

    const close = [...container.querySelectorAll('button')].find((button) =>
      /✕/.test(button.textContent ?? ''),
    ) as HTMLButtonElement | undefined;
    assert.ok(close);
    assert.equal(close.disabled, true);
    await act(async () => {
      (container.firstElementChild as HTMLElement).click();
    });
    assert.equal(closed, 0);

    await act(async () => {
      finishLink();
    });

    await act(async () => {
      root.unmount();
    });
  } finally {
    apiClient.searchTmdb = original.searchTmdb;
    apiClient.reassignAnimeTmdb = original.reassignAnimeTmdb;
    uninstallDom();
  }
});

for (const staleFailure of [false, true]) {
  test(`TmdbSelector ignores superseded ${staleFailure ? 'errors' : 'results'} and loading changes`, async () => {
    const uninstallDom = installDom();
    const originalSearch = apiClient.searchTmdb;
    const requests: Array<{
      resolve: (results: StatsTmdbSearchResult[]) => void;
      reject: (error: Error) => void;
    }> = [];
    apiClient.searchTmdb = () =>
      new Promise((resolve, reject) => {
        requests.push({ resolve, reject });
      });
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    try {
      const render = (initialQuery: string) =>
        root.render(
          <TmdbSelector
            animeId={1}
            initialQuery={initialQuery}
            onClose={() => {}}
            onLinked={() => {}}
          />,
        );
      await act(async () => {
        render('First title');
      });
      await act(async () => {
        render('Second title');
      });
      assert.equal(requests.length, 2);
      await act(async () => {
        if (staleFailure) requests[0]!.reject(new Error('stale error'));
        else requests[0]!.resolve([HANZAWA]);
      });
      assert.match(container.textContent ?? '', /Searching/);
      assert.doesNotMatch(container.textContent ?? '', /Hanzawa|failed/);
      await act(async () => {
        requests[1]!.resolve([HANZAWA]);
      });
      assert.match(container.textContent ?? '', /Hanzawa/);
      assert.doesNotMatch(container.textContent ?? '', /Searching/);
      await act(async () => {
        render('Third title');
      });
      await act(async () => {
        render('');
      });
      await act(async () => {
        requests[2]!.resolve([HANZAWA]);
      });
      assert.doesNotMatch(container.textContent ?? '', /Hanzawa|Searching/);
    } finally {
      await act(async () => {
        root.unmount();
      });
      apiClient.searchTmdb = originalSearch;
      uninstallDom();
    }
  });
}

test('TmdbSelector invalidates requests as soon as the user edits or clears the query', async () => {
  const uninstallDom = installDom();
  const originalSearch = apiClient.searchTmdb;
  const requests: Array<(results: StatsTmdbSearchResult[]) => void> = [];
  apiClient.searchTmdb = () =>
    new Promise((resolve) => {
      requests.push(resolve);
    });
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  try {
    await act(async () => {
      root.render(
        <TmdbSelector animeId={1} initialQuery="First" onClose={() => {}} onLinked={() => {}} />,
      );
    });
    const input = container.querySelector('input');
    assert.ok(input);
    const setValue = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype,
      'value',
    )?.set;
    assert.ok(setValue);
    const edit = async (value: string) => {
      await act(async () => {
        setValue.call(input, value);
        input.dispatchEvent(new window.Event('input', { bubbles: true }));
        input.dispatchEvent(new window.KeyboardEvent('keyup', { bubbles: true }));
      });
    };
    await edit('Second');
    await act(async () => {
      requests[0]!([HANZAWA]);
    });
    assert.doesNotMatch(container.textContent ?? '', /Hanzawa|No results/);
    assert.match(container.textContent ?? '', /Searching/);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 450));
    });
    assert.equal(requests.length, 2);
    await edit('');
    assert.doesNotMatch(container.textContent ?? '', /Searching/);
    await act(async () => {
      requests[1]!([HANZAWA]);
    });
    assert.doesNotMatch(container.textContent ?? '', /Hanzawa|Searching/);
  } finally {
    await act(async () => {
      root.unmount();
    });
    apiClient.searchTmdb = originalSearch;
    uninstallDom();
  }
});
