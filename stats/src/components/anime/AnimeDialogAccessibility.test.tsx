import assert from 'node:assert/strict';
import test from 'node:test';
import { Window } from 'happy-dom';
import { act, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import type { AnimeLibraryItem } from '../../types/stats';
import { AnimeMergeDialog } from './AnimeMergeDialog';
import { LibraryEntryPicker } from './LibraryEntryPicker';

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

test('AnimeMergeDialog focuses its close control, closes on Escape, and restores focus', async () => {
  const uninstallDom = installDom();
  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    function Harness() {
      const [open, setOpen] = useState(false);
      return (
        <>
          <button type="button" onClick={() => setOpen(true)}>
            Review merge
          </button>
          {open ? (
            <AnimeMergeDialog
              entries={[libraryItem(1, 'Show'), libraryItem(2, 'Show Season 1')]}
              onClose={() => setOpen(false)}
              onMerged={() => undefined}
            />
          ) : null}
        </>
      );
    }

    await act(async () => root.render(<Harness />));
    const trigger = container.querySelector('button') as HTMLButtonElement;
    trigger.focus();
    await act(async () => trigger.click());

    assert.equal(document.activeElement?.getAttribute('aria-label'), 'Close');
    await act(async () => {
      document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
    });
    assert.equal(container.querySelector('[role="dialog"]'), null);
    assert.equal(document.activeElement, trigger);

    await act(async () => root.unmount());
  } finally {
    uninstallDom();
  }
});

test('AnimeMergeDialog keeps keyboard focus inside the modal', async () => {
  const uninstallDom = installDom();
  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(
        <AnimeMergeDialog
          entries={[libraryItem(1, 'Show'), libraryItem(2, 'Show Season 1')]}
          onClose={() => undefined}
          onMerged={() => undefined}
        />,
      );
    });

    const dialog = container.querySelector('[role="dialog"]') as HTMLElement;
    const focusable = [...dialog.querySelectorAll('button:not([disabled])')] as HTMLButtonElement[];
    const first = focusable[0];
    const last = focusable.at(-1);
    assert.ok(first);
    assert.ok(last);

    last.focus();
    await act(async () => {
      document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Tab' }));
    });
    assert.equal(document.activeElement, first);

    first.focus();
    await act(async () => {
      document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Tab', shiftKey: true }));
    });
    assert.equal(document.activeElement, last);

    await act(async () => root.unmount());
  } finally {
    uninstallDom();
  }
});

test('LibraryEntryPicker focuses search, closes on Escape, and restores focus', async () => {
  const uninstallDom = installDom();
  const original = apiClient.getAnimeLibrary;
  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show'),
  ]) as typeof apiClient.getAnimeLibrary;
  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);

    function Harness() {
      const [open, setOpen] = useState(false);
      return (
        <>
          <button type="button" onClick={() => setOpen(true)}>
            Move
          </button>
          {open ? (
            <LibraryEntryPicker
              heading="Move episode"
              onSelect={() => undefined}
              onClose={() => setOpen(false)}
            />
          ) : null}
        </>
      );
    }

    await act(async () => root.render(<Harness />));
    const trigger = container.querySelector('button') as HTMLButtonElement;
    trigger.focus();
    await act(async () => trigger.click());

    assert.equal(document.activeElement?.getAttribute('placeholder'), 'Search library...');
    await act(async () => {
      document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
    });
    assert.equal(container.querySelector('[role="dialog"]'), null);
    assert.equal(document.activeElement, trigger);

    await act(async () => root.unmount());
  } finally {
    apiClient.getAnimeLibrary = original;
    uninstallDom();
  }
});

test('LibraryEntryPicker cannot be dismissed while a move is in flight', async () => {
  const uninstallDom = installDom();
  const original = apiClient.getAnimeLibrary;
  apiClient.getAnimeLibrary = (async () => [
    libraryItem(1, 'Show'),
  ]) as typeof apiClient.getAnimeLibrary;
  let closeCalls = 0;
  try {
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(
        <LibraryEntryPicker
          heading="Move episode"
          busyAnimeId={1}
          onSelect={() => undefined}
          onClose={() => {
            closeCalls += 1;
          }}
        />,
      );
    });

    const closeButton = container.querySelector('button[aria-label="Close"]') as HTMLButtonElement;
    assert.equal(closeButton.disabled, true);
    await act(async () => {
      closeButton.click();
      container.firstElementChild?.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
      document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
    });
    assert.equal(closeCalls, 0);

    await act(async () => root.unmount());
  } finally {
    apiClient.getAnimeLibrary = original;
    uninstallDom();
  }
});
