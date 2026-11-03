import assert from 'node:assert/strict';
import test from 'node:test';
import type { BrowserWindow } from 'electron';

import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from '../../core/services/overlay-window-flags';
import type { MpvIpcClient } from '../../core/services';
import type { ResolvedConfig } from '../../types';
import { createOverlayNotificationsRuntime } from './overlay-notifications-runtime';

function createRuntimeHarness(contentReady: boolean) {
  const sent: unknown[][] = [];
  const overlayWindow = {
    isDestroyed: () => false,
    [OVERLAY_WINDOW_CONTENT_READY_FLAG]: contentReady,
  } as unknown as BrowserWindow;
  const mpvClient = {
    send: ({ command }: { command: unknown[] }) => {
      sent.push(command);
    },
  } as unknown as MpvIpcClient;
  const runtime = createOverlayNotificationsRuntime({
    getResolvedConfig: () => ({}) as ResolvedConfig,
    getMainOverlayWindow: () => overlayWindow,
    getVisibleOverlayVisible: () => false,
    broadcastToOverlayWindows: () => {},
    showMpvOsd: () => {},
    getMpvClient: () => mpvClient,
    getAnkiIntegration: () => null,
    getRuntimeOptionsManager: () => null,
  });
  const scriptMessages = () =>
    sent.filter((command) => command[0] === 'script-message').map((command) => command[1]);
  return { runtime, scriptMessages };
}

test('dismissing the loading OSD before overlay content is ready only stops the plugin spinner', () => {
  const { runtime, scriptMessages } = createRuntimeHarness(false);

  runtime.dismissOverlayLoadingStatusNotification();

  assert.deepEqual(scriptMessages(), ['subminer-overlay-loading-dismissed']);
});

test('dismissing the loading OSD once overlay content is ready reports readiness to the plugin', () => {
  const { runtime, scriptMessages } = createRuntimeHarness(true);

  runtime.dismissOverlayLoadingStatusNotification();

  assert.deepEqual(scriptMessages(), ['subminer-overlay-loading-ready']);
});

test('reporting overlay readiness after an mpv connect resends ready only for loaded content', () => {
  const ready = createRuntimeHarness(true);
  ready.runtime.reportOverlayContentReadyToPlugin();
  assert.deepEqual(ready.scriptMessages(), ['subminer-overlay-loading-ready']);

  const loading = createRuntimeHarness(false);
  loading.runtime.reportOverlayContentReadyToPlugin();
  assert.deepEqual(loading.scriptMessages(), []);
});
