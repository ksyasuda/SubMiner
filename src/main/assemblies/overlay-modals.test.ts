import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayModalsAssembly, type OverlayModalsAssemblyDeps } from './overlay-modals';

type StatusCall = { message: string; title?: string; variant?: string };

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<OverlayModalsAssemblyDeps> = {}): OverlayModalsAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: { mpvClient: null, currentMediaPath: null },
    getConfig: lazy('getConfig'),
    logger: { warn: () => {}, error: () => {} },
    overlayModalRuntime: { waitForModalOpen: lazy('waitForModalOpen') },
    getOverlayNotifications: lazy('getOverlayNotifications'),
    ensureOverlayStartupPrereqs: lazy('ensureOverlayStartupPrereqs'),
    ensureOverlayWindowsReadyForVisibilityActions: lazy(
      'ensureOverlayWindowsReadyForVisibilityActions',
    ),
    sendToActiveOverlayWindow: lazy('sendToActiveOverlayWindow'),
    showMpvOsd: lazy('showMpvOsd'),
    showDesktopNotification: lazy('showDesktopNotification'),
    ...overrides,
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createOverlayModalsAssembly does not call its deps while constructing', () => {
  createOverlayModalsAssembly(makeDeps());
});

test('playlist browser warns instead of opening when mpv is not connected', () => {
  const calls: StatusCall[] = [];
  const overlayModals = createOverlayModalsAssembly(
    makeDeps({
      getOverlayNotifications: () => ({
        getConfiguredStatusNotificationType: () => 'overlay',
        showConfiguredStatusNotification: (message, options) => {
          calls.push({ message, title: options?.title, variant: options?.variant });
        },
        showOverlayNotification: () => {},
      }),
    }),
  );

  // The lazy overlay deps throw if the opener gets past the playback guard.
  overlayModals.openPlaylistBrowser();
  assert.deepEqual(calls, [
    {
      message: 'Playlist browser requires active playback.',
      title: 'Playlist browser',
      variant: 'warning',
    },
  ]);
});
