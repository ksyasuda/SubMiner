import assert from 'node:assert/strict';
import test from 'node:test';

import type { ChangelogSnapshot } from '../../types/changelog';
import { createRendererState } from '../state.js';
import { createChangelogModal } from './changelog.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
    toggle: (entry: string, force?: boolean) => {
      if (force === true) tokens.add(entry);
      else if (force === false) tokens.delete(entry);
      else if (tokens.has(entry)) tokens.delete(entry);
      else tokens.add(entry);
    },
  };
}

type SummaryStub = {
  classList: ReturnType<typeof createClassList>;
  tabIndex: number;
  dataset: Record<string, string>;
  getClientRects: () => Array<{ width: number; height: number }>;
  focusCount: number;
  scrollCount: number;
  focus: () => void;
  scrollIntoView: () => void;
};

function createSummaryStub(index: number): SummaryStub {
  const summary: SummaryStub = {
    classList: createClassList(),
    tabIndex: -1,
    dataset: { changelogIndex: String(index) },
    getClientRects: () => [{ width: 10, height: 10 }],
    focusCount: 0,
    scrollCount: 0,
    focus: () => {
      summary.focusCount += 1;
    },
    scrollIntoView: () => {
      summary.scrollCount += 1;
    },
  };
  return summary;
}

function createElementStub() {
  const listeners = new Map<string, Array<(event?: unknown) => void>>();
  return {
    value: '',
    textContent: '',
    innerHTML: '',
    classList: createClassList(['hidden']),
    contains: () => false,
    setAttribute: () => {},
    addEventListener: (type: string, listener: (event?: unknown) => void) => {
      listeners.set(type, [...(listeners.get(type) ?? []), listener]);
    },
    removeEventListener: () => {},
    appendChild: () => {},
    summaries: [] as SummaryStub[],
    querySelectorAll(this: { summaries: SummaryStub[] }) {
      return this.summaries;
    },
    focus: () => {},
    dispatchEventType: (type: string, event?: unknown) => {
      for (const listener of listeners.get(type) ?? []) listener(event);
    },
  };
}

const SNAPSHOT: ChangelogSnapshot = {
  entries: [],
  installedVersion: '0.19.2',
  latestVersion: '0.19.2',
  expandedGroupKey: '0.19',
  source: 'remote',
  releaseTag: 'v0.19.2',
};

type Harness = {
  modal: ReturnType<typeof createChangelogModal>;
  dom: Record<string, ReturnType<typeof createElementStub>>;
  snapshotRequests: Array<{ refresh?: boolean } | undefined>;
  modalClosedNotifications: string[];
  restore: () => void;
};

function createHarness(
  options: {
    getChangelogSnapshot?: (request?: { refresh?: boolean }) => Promise<ChangelogSnapshot>;
  } = {},
): Harness {
  const previousWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const previousDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const previousHTMLElement = Object.getOwnPropertyDescriptor(globalThis, 'HTMLElement');
  const previousElement = Object.getOwnPropertyDescriptor(globalThis, 'Element');
  const previousDetails = Object.getOwnPropertyDescriptor(globalThis, 'HTMLDetailsElement');

  const snapshotRequests: Array<{ refresh?: boolean } | undefined> = [];
  const modalClosedNotifications: string[] = [];

  class TestElement {}
  for (const name of ['HTMLElement', 'Element'] as const) {
    Object.defineProperty(globalThis, name, {
      configurable: true,
      writable: true,
      value: TestElement,
    });
  }
  // getSelectedEntry() narrows with `instanceof HTMLDetailsElement`.
  class TestDetailsElement {
    open = false;
  }
  Object.defineProperty(globalThis, 'HTMLDetailsElement', {
    configurable: true,
    writable: true,
    value: TestDetailsElement,
  });
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      electronAPI: {
        focusMainWindow: async () => {},
        setIgnoreMouseEvents: () => {},
        notifyOverlayModalClosed: (modal: string) => {
          modalClosedNotifications.push(modal);
        },
        getChangelogSnapshot: async (request?: { refresh?: boolean }) => {
          snapshotRequests.push(request);
          return options.getChangelogSnapshot
            ? await options.getChangelogSnapshot(request)
            : SNAPSHOT;
        },
      },
      focus: () => {},
      addEventListener: () => {},
      removeEventListener: () => {},
      setTimeout: (callback: () => void) => setTimeout(callback, 0),
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      activeElement: null,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });

  const dom = {
    overlay: createElementStub(),
    changelogModal: createElementStub(),
    changelogClose: createElementStub(),
    changelogRefresh: createElementStub(),
    changelogInstalled: createElementStub(),
    changelogSource: createElementStub(),
    changelogWarning: createElementStub(),
    changelogStatus: createElementStub(),
    changelogList: createElementStub(),
  };

  const modal = createChangelogModal(
    {
      state: createRendererState(),
      platform: {
        overlayLayer: 'modal',
        isModalLayer: true,
        isLinuxPlatform: false,
        isMacOSPlatform: false,
        isWindowsPlatform: true,
        shouldToggleMouseIgnore: false,
      },
      dom,
    } as never,
    {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    },
  );

  return {
    modal,
    dom,
    snapshotRequests,
    modalClosedNotifications,
    restore: () => {
      for (const [name, descriptor] of [
        ['window', previousWindow],
        ['document', previousDocument],
        ['HTMLElement', previousHTMLElement],
        ['Element', previousElement],
        ['HTMLDetailsElement', previousDetails],
      ] as const) {
        if (descriptor) {
          Object.defineProperty(globalThis, name, descriptor);
        } else {
          delete (globalThis as Record<string, unknown>)[name];
        }
      }
    },
  };
}

test('changelog modal loads a snapshot on open and shows the installed version', async () => {
  const harness = createHarness();
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(harness.snapshotRequests, [undefined]);
    assert.equal(harness.dom.changelogInstalled?.textContent, 'Installed v0.19.2');
    assert.equal(harness.dom.changelogSource?.textContent, 'Latest release v0.19.2');
    assert.equal(harness.dom.changelogModal?.classList.contains('hidden'), false);
  } finally {
    harness.restore();
  }
});

test('changelog modal surfaces the bundled-fallback warning', async () => {
  const harness = createHarness({
    getChangelogSnapshot: async () => ({
      ...SNAPSHOT,
      source: 'bundled',
      releaseTag: undefined,
      warning: 'Showing the bundled changelog: offline',
    }),
  });
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(
      harness.dom.changelogWarning?.textContent,
      'Showing the bundled changelog: offline',
    );
    assert.equal(harness.dom.changelogSource?.textContent, 'Bundled changelog');
  } finally {
    harness.restore();
  }
});

test('changelog modal reports a load failure instead of hanging on the spinner text', async () => {
  const harness = createHarness({
    getChangelogSnapshot: async () => {
      throw new Error('ipc down');
    },
  });
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.match(
      harness.dom.changelogList?.textContent ?? '',
      /Changelog failed to load: ipc down/,
    );
    assert.equal(harness.dom.changelogStatus?.textContent, 'Press Esc to close.');
  } finally {
    harness.restore();
  }
});

test('changelog modal clears stale metadata when a later open fails to load', async () => {
  let shouldFail = false;
  const harness = createHarness({
    getChangelogSnapshot: async () => {
      if (shouldFail) throw new Error('offline');
      return SNAPSHOT;
    },
  });
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(harness.dom.changelogInstalled?.textContent, 'Installed v0.19.2');
    harness.modal.closeChangelogModal();

    shouldFail = true;
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    // The previous session's values must not linger behind the error message.
    assert.equal(harness.dom.changelogInstalled?.textContent, '');
    assert.equal(harness.dom.changelogSource?.textContent, '');
    assert.match(harness.dom.changelogList?.textContent ?? '', /Changelog failed to load: offline/);
  } finally {
    harness.restore();
  }
});

test('changelog modal clears stale metadata when a refresh fails', async () => {
  let shouldFail = false;
  const harness = createHarness({
    getChangelogSnapshot: async () => {
      if (shouldFail) throw new Error('refresh offline');
      return { ...SNAPSHOT, warning: 'Showing the bundled changelog: earlier failure' };
    },
  });
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(harness.dom.changelogInstalled?.textContent, 'Installed v0.19.2');

    shouldFail = true;
    harness.modal.handleChangelogKeydown({
      key: 'r',
      preventDefault: () => {},
    } as KeyboardEvent);
    await new Promise((resolve) => setTimeout(resolve, 0));

    // The rendered snapshot is gone, so nothing may still describe it.
    assert.equal(harness.dom.changelogInstalled?.textContent, '');
    assert.equal(harness.dom.changelogSource?.textContent, '');
    assert.equal(harness.dom.changelogWarning?.textContent, '');
    assert.match(
      harness.dom.changelogList?.textContent ?? '',
      /Changelog failed to load: refresh offline/,
    );
  } finally {
    harness.restore();
  }
});

test('changelog modal moves selection styling and focus together on J/K', async () => {
  const harness = createHarness();
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const summaries = [createSummaryStub(0), createSummaryStub(1), createSummaryStub(2)];
    harness.dom.changelogList!.summaries = summaries;

    harness.modal.handleChangelogKeydown({ key: 'j', preventDefault: () => {} } as KeyboardEvent);

    assert.deepEqual(
      summaries.map((summary) => summary.classList.contains('active')),
      [false, true, false],
    );
    assert.deepEqual(
      summaries.map((summary) => summary.tabIndex),
      [-1, 0, -1],
    );
    assert.equal(summaries[1]?.focusCount, 1, 'the keyboard path focuses the new selection');
    assert.equal(summaries[1]?.scrollCount, 1);

    // Wraps backwards past the start.
    harness.modal.handleChangelogKeydown({ key: 'k', preventDefault: () => {} } as KeyboardEvent);
    harness.modal.handleChangelogKeydown({ key: 'k', preventDefault: () => {} } as KeyboardEvent);
    assert.deepEqual(
      summaries.map((summary) => summary.classList.contains('active')),
      [false, false, true],
    );
  } finally {
    harness.restore();
  }
});

test('changelog modal click selection restyles without stealing focus back', async () => {
  const harness = createHarness();
  try {
    harness.modal.wireDomEvents();
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const summaries = [createSummaryStub(0), createSummaryStub(1)];
    harness.dom.changelogList!.summaries = summaries;

    // The handler guards on `instanceof Element`, so the target has to inherit
    // from the Element stand-in the harness installs.
    const elementCtor = (globalThis as unknown as { Element: { prototype: object } }).Element;
    const clickTarget = Object.assign(Object.create(elementCtor.prototype), {
      closest: (selector: string) =>
        selector === '.changelog-entry-summary' ? summaries[1] : null,
    });
    harness.dom.changelogList!.dispatchEventType('click', { target: clickTarget });

    assert.deepEqual(
      summaries.map((summary) => summary.classList.contains('active')),
      [false, true],
    );
    assert.deepEqual(
      summaries.map((summary) => summary.tabIndex),
      [-1, 0],
    );
    // The browser already focused the clicked summary; re-focusing would fight it.
    assert.equal(summaries[1]?.focusCount, 0);
  } finally {
    harness.restore();
  }
});

test('changelog modal folds on Enter only from the selected summary', async () => {
  const harness = createHarness();
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const detailsCtor = (
      globalThis as unknown as { HTMLDetailsElement: new () => { open: boolean } }
    ).HTMLDetailsElement;
    const entry = new detailsCtor();
    entry.open = true;

    const summaries = [createSummaryStub(0), createSummaryStub(1)];
    Object.assign(summaries[0]!, { parentElement: entry });
    harness.dom.changelogList!.summaries = summaries;

    let prevented = 0;
    const press = (target: unknown) =>
      harness.modal.handleChangelogKeydown({
        key: 'Enter',
        target,
        preventDefault: () => {
          prevented += 1;
        },
      } as unknown as KeyboardEvent);

    // Close button focused: the button must keep its own Enter activation.
    assert.equal(press(harness.dom.changelogClose), true);
    assert.equal(prevented, 0, 'Enter on a button is not swallowed');
    assert.equal(entry.open, true);

    // A non-selected summary (the nested "Internal changes" fold) is left alone.
    assert.equal(press(summaries[1]), true);
    assert.equal(prevented, 0);
    assert.equal(entry.open, true);

    // The selected summary does fold, exactly once.
    assert.equal(press(summaries[0]), true);
    assert.equal(prevented, 1);
    assert.equal(entry.open, false);
  } finally {
    harness.restore();
  }
});

test('changelog modal closes on Escape and notifies the main process', async () => {
  const harness = createHarness();
  try {
    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const handled = harness.modal.handleChangelogKeydown({
      key: 'Escape',
      preventDefault: () => {},
    } as KeyboardEvent);

    assert.equal(handled, true);
    assert.deepEqual(harness.modalClosedNotifications, ['changelog']);
    assert.equal(harness.dom.changelogModal?.classList.contains('hidden'), true);
  } finally {
    harness.restore();
  }
});

test('changelog modal refetches on R and ignores keys while closed', async () => {
  const harness = createHarness();
  try {
    assert.equal(
      harness.modal.handleChangelogKeydown({ key: 'r', preventDefault: () => {} } as KeyboardEvent),
      false,
    );

    harness.modal.openChangelogModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    harness.modal.handleChangelogKeydown({
      key: 'r',
      preventDefault: () => {},
    } as KeyboardEvent);
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(harness.snapshotRequests, [undefined, { refresh: true }]);
  } finally {
    harness.restore();
  }
});

test('changelog modal drops a late in-flight load after close', async () => {
  const pending: Array<(snapshot: ChangelogSnapshot) => void> = [];
  const harness = createHarness({
    getChangelogSnapshot: () =>
      new Promise<ChangelogSnapshot>((resolve) => {
        pending.push(resolve);
      }),
  });
  try {
    harness.modal.openChangelogModal();
    harness.modal.closeChangelogModal();
    pending[0]?.(SNAPSHOT);
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(harness.dom.changelogInstalled?.textContent, '');
  } finally {
    harness.restore();
  }
});
