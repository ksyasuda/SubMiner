import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { apiClient } from '../lib/api-client';
import { resetExcludedWordsStoreForTests, setExcludedWords } from './useExcludedWords';
import { useVocabulary } from './useVocabulary';
import type { StatsVocabularyCharts, StatsVocabularySummary } from '../types/stats';

type VocabularyState = ReturnType<typeof useVocabulary>;

function installDom(): () => void {
  const previousWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const previousDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const previousHTMLElement = Object.getOwnPropertyDescriptor(globalThis, 'HTMLElement');
  const previousIsReactActEnvironment = Object.getOwnPropertyDescriptor(
    globalThis,
    'IS_REACT_ACT_ENVIRONMENT',
  );
  const window = new Window();

  Object.defineProperty(globalThis, 'window', { value: window, configurable: true });
  Object.defineProperty(globalThis, 'document', { value: window.document, configurable: true });
  Object.defineProperty(globalThis, 'HTMLElement', {
    value: window.HTMLElement,
    configurable: true,
  });
  Object.defineProperty(globalThis, 'IS_REACT_ACT_ENVIRONMENT', {
    value: true,
    configurable: true,
    writable: true,
  });

  return () => {
    const restoreProperty = (name: string, descriptor: PropertyDescriptor | undefined) => {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor);
      else Reflect.deleteProperty(globalThis, name);
    };
    restoreProperty('window', previousWindow);
    restoreProperty('document', previousDocument);
    restoreProperty('HTMLElement', previousHTMLElement);
    restoreProperty('IS_REACT_ACT_ENVIRONMENT', previousIsReactActEnvironment);
  };
}

test('DOM harness restores the original global property descriptors', () => {
  const propertyNames = [
    'window',
    'document',
    'HTMLElement',
    'IS_REACT_ACT_ENVIRONMENT',
  ] as const;
  const before = propertyNames.map((name) => Object.getOwnPropertyDescriptor(globalThis, name));

  const restore = installDom();
  restore();

  const after = propertyNames.map((name) => Object.getOwnPropertyDescriptor(globalThis, name));
  assert.deepEqual(after, before);
});

function installLocalStorage(): () => void {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const values = new Map<string, string>();
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
    },
  });
  return () => {
    if (previous) Object.defineProperty(globalThis, 'localStorage', previous);
    else delete (globalThis as { localStorage?: unknown }).localStorage;
  };
}

interface FakeClock {
  tick: (ms: number) => void;
  restore: () => void;
}

/** Bun's `node:test` shim has no `mock.timers`, so the retry clock is faked here. */
function installFakeTimers(): FakeClock {
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  const timers = new Map<number, { at: number; fn: () => void }>();
  let now = 0;
  let nextId = 1;

  globalThis.setTimeout = ((fn: () => void, delay = 0) => {
    const id = nextId;
    nextId += 1;
    timers.set(id, { at: now + delay, fn });
    return id;
  }) as unknown as typeof globalThis.setTimeout;
  globalThis.clearTimeout = ((id: number) => {
    timers.delete(id);
  }) as unknown as typeof globalThis.clearTimeout;

  return {
    tick: (ms: number) => {
      now += ms;
      const due = [...timers.entries()]
        .filter(([, timer]) => timer.at <= now)
        .sort(([, a], [, b]) => a.at - b.at);
      for (const [id, timer] of due) {
        timers.delete(id);
        timer.fn();
      }
    },
    restore: () => {
      globalThis.setTimeout = originalSetTimeout;
      globalThis.clearTimeout = originalClearTimeout;
    },
  };
}

function summaryFixture(): StatsVocabularySummary {
  return {
    uniqueWords: 42,
    uniqueWordsWithoutNames: 40,
    uniqueKanji: 7,
    newThisWeek: 3,
    newThisWeekWithoutNames: 2,
    knownWordCount: 10,
    knownWordCountWithoutNames: 9,
  };
}

function chartsFixture(overrides: Partial<StatsVocabularyCharts> = {}): StatsVocabularyCharts {
  return {
    ready: true,
    topWords: [{ wordId: 1, headword: '猫', frequency: 5 }],
    topWordsWithoutNames: [{ wordId: 1, headword: '猫', frequency: 5 }],
    newWordsTimeline: [{ epochDay: 20_000, wordCount: 4 }],
    newWordsTimelineWithoutNames: [{ epochDay: 20_000, wordCount: 4 }],
    ...overrides,
  };
}

interface Harness {
  state: () => VocabularyState;
  flush: () => Promise<void>;
  tick: (ms: number) => Promise<void>;
  unmount: () => Promise<void>;
  teardown: () => Promise<void>;
}

async function mountHook(): Promise<Harness> {
  const uninstallDom = installDom();
  const uninstallLocalStorage = installLocalStorage();
  const clock = installFakeTimers();

  let latest: VocabularyState | null = null;
  function Probe() {
    latest = useVocabulary();
    return null;
  }

  const container = document.createElement('div');
  document.body.append(container);
  let root: Root | null = createRoot(container);
  await act(async () => {
    root!.render(<Probe />);
  });

  const flush = async (): Promise<void> => {
    // Drain promise callbacks without advancing the mocked clock.
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
  };

  return {
    state: () => {
      assert.ok(latest, 'expected the hook to have rendered');
      return latest;
    },
    flush,
    tick: async (ms: number) => {
      await act(async () => {
        clock.tick(ms);
      });
      await flush();
    },
    unmount: async () => {
      await act(async () => {
        root?.unmount();
        root = null;
      });
    },
    teardown: async () => {
      await act(async () => {
        root?.unmount();
        root = null;
      });
      clock.restore();
      // React's scheduler can still have deferred work queued; let it drain on
      // a real timer while the DOM globals it reads are still installed.
      await new Promise((resolve) => setTimeout(resolve, 0));
      uninstallLocalStorage();
      uninstallDom();
      resetExcludedWordsStoreForTests();
    },
  };
}

function stubVocabularyClient(overrides: {
  getVocabularySummary: () => Promise<StatsVocabularySummary>;
  getVocabularyCharts: () => Promise<StatsVocabularyCharts>;
}): () => void {
  const original = {
    getVocabulary: apiClient.getVocabulary,
    getKanji: apiClient.getKanji,
    getKnownWords: apiClient.getKnownWords,
    getVocabularySummary: apiClient.getVocabularySummary,
    getVocabularyCharts: apiClient.getVocabularyCharts,
    setExcludedWords: apiClient.setExcludedWords,
  };
  apiClient.getVocabulary = async () => [];
  apiClient.getKanji = async () => [];
  apiClient.getKnownWords = async () => [];
  apiClient.setExcludedWords = async () => {};
  apiClient.getVocabularySummary = overrides.getVocabularySummary;
  apiClient.getVocabularyCharts = overrides.getVocabularyCharts;
  return () => Object.assign(apiClient, original);
}

test('aggregate failures retry with backoff, then surface an error that Retry clears', async () => {
  const originalConsoleError = console.error;
  console.error = () => {};
  let summaryCalls = 0;
  let failSummary = true;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => {
      summaryCalls += 1;
      if (failSummary) throw new Error('summary unavailable');
      return summaryFixture();
    },
    getVocabularyCharts: async () => chartsFixture(),
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    assert.equal(summaryCalls, 1);
    assert.equal(harness.state().aggregatesError, null, 'no error until retries are exhausted');

    // Backoff is 1s, 2s, 4s, 8s across the remaining four attempts.
    for (const delayMs of [1_000, 2_000, 4_000, 8_000]) {
      await harness.tick(delayMs);
    }
    assert.equal(summaryCalls, 5, 'retries are bounded at the attempt limit');
    assert.match(harness.state().aggregatesError ?? '', /totals failed to load/i);

    // Nothing further is scheduled once the limit is reached.
    await harness.tick(60_000);
    assert.equal(summaryCalls, 5);

    failSummary = false;
    await act(async () => {
      harness.state().refreshAggregates();
    });
    await harness.flush();

    assert.equal(summaryCalls, 6);
    assert.equal(harness.state().aggregatesError, null);
    assert.deepEqual(harness.state().summary, summaryFixture());
  } finally {
    await harness.teardown();
    restoreClient();
    console.error = originalConsoleError;
  }
});

test('charts poll while the backfill is pending and stop once it is ready', async () => {
  let chartCalls = 0;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => summaryFixture(),
    getVocabularyCharts: async () => {
      chartCalls += 1;
      return chartsFixture({ ready: chartCalls >= 3 });
    },
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    assert.equal(chartCalls, 1);
    assert.equal(harness.state().charts?.ready, false);

    await harness.tick(1_000);
    assert.equal(chartCalls, 2);
    await harness.tick(1_000);
    assert.equal(chartCalls, 3);
    assert.equal(harness.state().charts?.ready, true);

    // A ready result ends the poll.
    await harness.tick(60_000);
    assert.equal(chartCalls, 3);
  } finally {
    await harness.teardown();
    restoreClient();
  }
});

test('chart backfill polling stops and surfaces Retry when readiness never arrives', async () => {
  let chartCalls = 0;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => summaryFixture(),
    getVocabularyCharts: async () => {
      chartCalls += 1;
      return chartsFixture({ ready: false });
    },
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    for (let poll = 0; poll < 65; poll += 1) await harness.tick(5_000);

    assert.equal(chartCalls, 60, 'a failed backfill must not poll for the lifetime of the tab');
    assert.match(harness.state().aggregatesError ?? '', /still building/i);

    await act(async () => {
      harness.state().refreshAggregates();
    });
    await harness.flush();
    assert.equal(chartCalls, 61, 'Retry starts one fresh bounded polling cycle');
    assert.equal(harness.state().aggregatesError, null);
  } finally {
    await harness.teardown();
    restoreClient();
  }
});

test('aggregates refetch after an exclusion edit is acknowledged by the server', async () => {
  let summaryCalls = 0;
  let chartCalls = 0;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => {
      summaryCalls += 1;
      return summaryFixture();
    },
    getVocabularyCharts: async () => {
      chartCalls += 1;
      return chartsFixture();
    },
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    assert.equal(summaryCalls, 1);
    assert.equal(chartCalls, 1);

    await act(async () => {
      await setExcludedWords([{ headword: '猫', word: '猫', reading: 'ねこ' }]);
    });
    await harness.flush();

    assert.equal(summaryCalls, 2, 'totals must not keep counting the excluded word');
    assert.equal(chartCalls, 2);
  } finally {
    await harness.teardown();
    restoreClient();
  }
});

test('pending retries are cancelled when the tab unmounts', async () => {
  const originalConsoleError = console.error;
  console.error = () => {};
  let summaryCalls = 0;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => {
      summaryCalls += 1;
      throw new Error('summary unavailable');
    },
    getVocabularyCharts: async () => chartsFixture(),
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    assert.equal(summaryCalls, 1);

    await harness.unmount();
    await harness.tick(60_000);

    assert.equal(summaryCalls, 1, 'no retry may run after unmount');
  } finally {
    await harness.teardown();
    restoreClient();
    console.error = originalConsoleError;
  }
});

test('a slow response from a superseded refresh cannot replace the newest aggregates', async () => {
  let summaryCalls = 0;
  let releaseSuperseded: (() => void) | null = null;
  const restoreClient = stubVocabularyClient({
    getVocabularySummary: async () => {
      summaryCalls += 1;
      const call = summaryCalls;
      // The second call is the one that gets superseded while still in flight.
      if (call === 2) {
        await new Promise<void>((resolve) => {
          releaseSuperseded = resolve;
        });
      }
      return { ...summaryFixture(), uniqueWords: call };
    },
    getVocabularyCharts: async () => chartsFixture(),
  });
  const harness = await mountHook();

  try {
    await harness.flush();
    assert.equal(harness.state().summary?.uniqueWords, 1);

    // First refresh stalls, then a second refresh supersedes it and resolves.
    await act(async () => {
      harness.state().refreshAggregates();
    });
    await harness.flush();
    await act(async () => {
      harness.state().refreshAggregates();
    });
    await harness.flush();

    assert.equal(summaryCalls, 3);
    assert.equal(harness.state().summary?.uniqueWords, 3);

    const release = releaseSuperseded as (() => void) | null;
    assert.ok(release, 'expected the superseded request to still be in flight');
    release();
    await harness.flush();

    assert.equal(
      harness.state().summary?.uniqueWords,
      3,
      'the superseded response must not overwrite the newest totals',
    );
  } finally {
    await harness.teardown();
    restoreClient();
  }
});
