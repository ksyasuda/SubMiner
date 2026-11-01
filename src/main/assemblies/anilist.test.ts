import assert from 'node:assert/strict';
import test from 'node:test';
import { createAnilistRateLimiter } from '../../core/services/anilist/rate-limiter';
import { createAnilistAssembly, type AnilistAssemblyDeps } from './anilist';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): AnilistAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      mpvClient: null,
      currentMediaPath: null,
      currentMediaTitle: null,
      anilistClientSecretState: {
        status: 'not_checked',
        source: 'none',
        message: null,
        resolvedAt: null,
        errorAt: null,
      },
      anilistRetryQueueState: {
        pending: 0,
        ready: 0,
        deadLetter: 0,
        lastAttemptAt: null,
        lastError: null,
      },
      anilistSetupPageOpened: false,
      anilistSetupWindow: null,
    },
    configService: { getConfig: lazy('getConfig') },
    logger: { info: lazy('info'), warn: lazy('warn'), debug: lazy('debug'), error: lazy('error') },
    anilistTokenStore: {
      saveToken: lazy('saveToken'),
      loadToken: lazy('loadToken'),
      clearToken: lazy('clearToken'),
    },
    anilistUpdateQueue: {
      enqueue: lazy('enqueue'),
      nextReady: lazy('nextReady'),
      markSuccess: lazy('markSuccess'),
      markFailure: lazy('markFailure'),
      getSnapshot: lazy('getSnapshot'),
    },
    anilistRateLimiter: createAnilistRateLimiter(),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    showDesktopNotification: lazy('showDesktopNotification'),
    resolveMediaPathForJimaku: lazy('resolveMediaPathForJimaku'),
    resolvePinnedMediaId: lazy('resolvePinnedMediaId'),
    createBrowserWindow: lazy('createBrowserWindow'),
    shell: { openExternal: lazy('openExternal') },
    app: { setAsDefaultProtocolClient: lazy('setAsDefaultProtocolClient') },
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createAnilistAssembly does not call its deps while constructing', () => {
  const anilist = createAnilistAssembly(makeDeps());
  assert.equal(anilist.getMediaDurationSec(), null);
});
