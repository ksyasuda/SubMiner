import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { StatsDuplicateLineCleanupResult } from '../../types/stats';
import { DuplicateLineCleanup } from './DuplicateLineCleanup';

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
    Object.defineProperty(globalThis, 'document', { value: previousDocument, configurable: true });
    Object.defineProperty(globalThis, 'HTMLElement', {
      value: previousHTMLElement,
      configurable: true,
    });
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = previousISReactActEnvironment;
  };
}

function findButton(container: Element, label: string): HTMLButtonElement {
  const match = [...container.querySelectorAll('button')].find(
    (button) => (button.textContent ?? '').trim() === label,
  );
  assert.ok(match, `expected a "${label}" button`);
  return match as unknown as HTMLButtonElement;
}

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

function summary(
  overrides: Partial<StatsDuplicateLineCleanupResult> = {},
): StatsDuplicateLineCleanupResult {
  return {
    dryRun: false,
    lookbackDays: 30,
    scannedLines: 900,
    burstGroups: 2,
    removedLines: 180,
    removedWordOccurrences: 540,
    removedKanjiOccurrences: 120,
    samples: [],
    ...overrides,
  };
}

interface Harness {
  container: Element;
  cleanedCalls: () => number;
  closedCalls: () => number;
  teardown: () => void;
}

async function mount(cleanup: (typeof apiClient)['cleanupDuplicateLines']): Promise<Harness> {
  const uninstallDom = installDom();
  const originalCleanup = apiClient.cleanupDuplicateLines;
  apiClient.cleanupDuplicateLines = cleanup;

  let cleaned = 0;
  let closed = 0;
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);

  await act(async () => {
    root.render(
      <DuplicateLineCleanup
        onClose={() => {
          closed += 1;
        }}
        onCleaned={() => {
          cleaned += 1;
        }}
      />,
    );
  });

  return {
    container,
    cleanedCalls: () => cleaned,
    closedCalls: () => closed,
    teardown: () => {
      apiClient.cleanupDuplicateLines = originalCleanup;
      uninstallDom();
    },
  };
}

test('a reload is still owed after a later scan replaces the applied result', async () => {
  const harness = await mount(async ({ dryRun } = {}) => summary({ dryRun: dryRun === true }));

  try {
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Clean Up').click();
    });
    assert.equal(harness.cleanedCalls(), 0, 'reload must wait for the result to be read');

    // The follow-up scan clears the applied summary, but the rows are already gone.
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Close').click();
    });

    assert.equal(harness.cleanedCalls(), 1);
    assert.equal(harness.closedCalls(), 1);
  } finally {
    harness.teardown();
  }
});

test('a reload is still owed after the lookback window changes', async () => {
  const harness = await mount(async ({ dryRun } = {}) => summary({ dryRun: dryRun === true }));

  try {
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Clean Up').click();
    });
    await act(async () => {
      findButton(harness.container, '7 days').click();
    });
    await act(async () => {
      findButton(harness.container, 'Close').click();
    });

    assert.equal(harness.cleanedCalls(), 1);
  } finally {
    harness.teardown();
  }
});

test('closing is refused while an apply is in flight', async () => {
  const pending = deferred<StatsDuplicateLineCleanupResult>();
  const harness = await mount(async ({ dryRun } = {}) =>
    dryRun === true ? summary({ dryRun: true }) : pending.promise,
  );

  try {
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Clean Up').click();
    });

    await act(async () => {
      findButton(harness.container, 'Close').click();
    });
    assert.equal(harness.closedCalls(), 0, 'the modal must stay open mid-apply');
    assert.equal(harness.cleanedCalls(), 0);

    await act(async () => {
      pending.resolve(summary({ removedLines: 12 }));
      await pending.promise;
    });
    await act(async () => {
      findButton(harness.container, 'Close').click();
    });

    assert.equal(harness.closedCalls(), 1);
    assert.equal(harness.cleanedCalls(), 1);
  } finally {
    harness.teardown();
  }
});

test('a scan on its own owes no reload', async () => {
  const harness = await mount(async ({ dryRun } = {}) => summary({ dryRun: dryRun === true }));

  try {
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Close').click();
    });

    assert.equal(harness.cleanedCalls(), 0);
    assert.equal(harness.closedCalls(), 1);
  } finally {
    harness.teardown();
  }
});

test('an apply that removes nothing owes no reload', async () => {
  // The scan saw work to do, but by the time it ran another cleanup had taken it.
  const harness = await mount(async ({ dryRun } = {}) =>
    dryRun === true ? summary({ dryRun: true }) : summary({ burstGroups: 0, removedLines: 0 }),
  );

  try {
    await act(async () => {
      findButton(harness.container, 'Scan').click();
    });
    await act(async () => {
      findButton(harness.container, 'Clean Up').click();
    });
    await act(async () => {
      findButton(harness.container, 'Close').click();
    });

    assert.equal(harness.cleanedCalls(), 0);
    assert.equal(harness.closedCalls(), 1);
  } finally {
    harness.teardown();
  }
});
