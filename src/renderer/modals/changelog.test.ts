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

function createElementStub() {
  return {
    value: '',
    textContent: '',
    innerHTML: '',
    classList: createClassList(['hidden']),
    contains: () => false,
    setAttribute: () => {},
    addEventListener: () => {},
    removeEventListener: () => {},
    appendChild: () => {},
    querySelectorAll: () => [],
    focus: () => {},
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
