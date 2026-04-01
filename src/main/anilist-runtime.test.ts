import assert from 'node:assert/strict';
import test from 'node:test';

import { createAnilistRuntime } from './anilist-runtime';

function createSetupWindow() {
  const handlers = new Map<string, Array<(...args: unknown[]) => void>>();
  let destroyed = false;
  return {
    window: {
      focus: () => {},
      close: () => {
        destroyed = true;
        for (const handler of handlers.get('closed') ?? []) {
          handler();
        }
      },
      isDestroyed: () => destroyed,
      on: (event: 'closed', handler: () => void) => {
        handlers.set(event, [...(handlers.get(event) ?? []), handler]);
      },
      loadURL: async () => {},
      webContents: {
        setWindowOpenHandler: () => ({ action: 'deny' as const }),
        on: (event: string, handler: (...args: unknown[]) => void) => {
          handlers.set(event, [...(handlers.get(event) ?? []), handler]);
        },
        getURL: () => 'about:blank',
      },
    },
  };
}

function createRuntime(overrides: Partial<Parameters<typeof createAnilistRuntime>[0]> = {}) {
  const savedTokens: string[] = [];
  const queueCalls: string[] = [];
  const notifications: string[] = [];
  const state = {
    config: {
      anilist: {
        enabled: true,
        accessToken: '',
      },
    },
  };
  const setup = createSetupWindow();

  const runtime = createAnilistRuntime({
    getResolvedConfig: () => state.config,
    isTrackingEnabled: (config) => config.anilist.enabled === true,
    tokenStore: {
      saveToken: (token) => {
        savedTokens.push(token);
      },
      loadToken: () => null,
      clearToken: () => {
        savedTokens.push('cleared');
      },
    },
    updateQueue: {
      enqueue: (key, title, episode) => {
        queueCalls.push(`enqueue:${key}:${title}:${episode}`);
      },
      nextReady: () => ({
        key: 'retry-1',
        title: 'Demo',
        episode: 2,
        createdAt: 1,
        attemptCount: 0,
        nextAttemptAt: 0,
        lastError: null,
      }),
      markSuccess: (key) => {
        queueCalls.push(`success:${key}`);
      },
      markFailure: (key, message) => {
        queueCalls.push(`failure:${key}:${message}`);
      },
      getSnapshot: () => ({
        pending: 3,
        ready: 1,
        deadLetter: 2,
      }),
    },
    getCurrentMediaPath: () => '/tmp/demo.mkv',
    getCurrentMediaTitle: () => 'Demo',
    getWatchedSeconds: () => 0,
    hasMpvClient: () => false,
    requestMpvDuration: async () => 120,
    resolveMediaPathForJimaku: (value) => value,
    guessAnilistMediaInfo: async () => null,
    updateAnilistPostWatchProgress: async () => ({
      status: 'updated',
      message: 'updated ok',
    }),
    createBrowserWindow: () => setup.window,
    authorizeUrl: 'https://anilist.co/api/v2/oauth/authorize',
    clientId: '36084',
    responseType: 'token',
    redirectUri: 'https://anilist.subminer.moe/',
    developerSettingsUrl: 'https://anilist.co/settings/developer',
    isAllowedExternalUrl: () => true,
    isAllowedNavigationUrl: () => true,
    openExternal: async () => {},
    showMpvOsd: (message) => {
      notifications.push(`osd:${message}`);
    },
    showDesktopNotification: (_title, options) => {
      notifications.push(`notify:${options.body}`);
    },
    logInfo: (message) => {
      notifications.push(`info:${message}`);
    },
    logWarn: () => {},
    logError: () => {},
    logDebug: () => {},
    isDefaultApp: () => false,
    getArgv: () => [],
    execPath: process.execPath,
    resolvePath: (value) => value,
    setAsDefaultProtocolClient: () => true,
    now: () => 1234,
    ...overrides,
  });

  return {
    runtime,
    state,
    savedTokens,
    queueCalls,
    notifications,
  };
}

test('anilist runtime saves setup token and updates resolved state', () => {
  const harness = createRuntime();

  const consumed = harness.runtime.consumeAnilistSetupTokenFromUrl(
    'subminer://anilist-setup?access_token=token-123',
  );

  assert.equal(consumed, true);
  assert.deepEqual(harness.savedTokens, ['token-123']);
  assert.equal(harness.runtime.getStatusSnapshot().tokenStatus, 'resolved');
  assert.equal(harness.runtime.getStatusSnapshot().tokenSource, 'stored');
  assert.equal(harness.runtime.getStatusSnapshot().tokenMessage, 'saved token from AniList login');
  assert.ok(harness.notifications.includes('notify:AniList login success'));
});

test('anilist runtime bypasses refresh when tracking disabled', async () => {
  const harness = createRuntime();
  harness.state.config = {
    anilist: {
      enabled: false,
      accessToken: '',
    },
  };

  const token = await harness.runtime.refreshAnilistClientSecretStateIfEnabled({
    force: true,
  });

  assert.equal(token, null);
  assert.equal(harness.runtime.getStatusSnapshot().tokenStatus, 'not_checked');
  assert.equal(harness.runtime.getStatusSnapshot().tokenSource, 'none');
});

test('anilist runtime refreshes queue snapshot and retry state after processing', async () => {
  const harness = createRuntime({
    tokenStore: {
      saveToken: () => {},
      loadToken: () => 'stored-token',
      clearToken: () => {},
    },
  });

  harness.runtime.refreshRetryQueueState();
  assert.deepEqual(harness.runtime.getQueueStatusSnapshot(), {
    pending: 3,
    ready: 1,
    deadLetter: 2,
    lastAttemptAt: null,
    lastError: null,
  });

  const result = await harness.runtime.processNextAnilistRetryUpdate();

  assert.deepEqual(result, { ok: true, message: 'updated ok' });
  assert.ok(harness.queueCalls.includes('success:retry-1'));
  assert.equal(harness.runtime.getQueueStatusSnapshot().lastAttemptAt, 1234);
  assert.equal(harness.runtime.getQueueStatusSnapshot().lastError, null);
});
