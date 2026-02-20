import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createHandleAnilistSetupWindowClosedHandler,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createHandleAnilistSetupWindowOpenedHandler,
  createAnilistSetupDidFailLoadHandler,
  createAnilistSetupDidFinishLoadHandler,
  createAnilistSetupDidNavigateHandler,
  createAnilistSetupFallbackHandler,
  createAnilistSetupWillNavigateHandler,
  createAnilistSetupWillRedirectHandler,
  createAnilistSetupWindowOpenHandler,
  createHandleManualAnilistSetupSubmissionHandler,
} from './anilist-setup-window';

test('manual anilist setup submission forwards access token to callback consumer', () => {
  const consumed: string[] = [];
  const handleSubmission = createHandleManualAnilistSetupSubmissionHandler({
    consumeCallbackUrl: (rawUrl) => {
      consumed.push(rawUrl);
      return true;
    },
    redirectUri: 'https://anilist.subminer.moe/',
    logWarn: () => {},
  });

  const handled = handleSubmission('subminer://anilist-setup?access_token=abc123');
  assert.equal(handled, true);
  assert.equal(consumed.length, 1);
  assert.ok(consumed[0].includes('https://anilist.subminer.moe/#access_token=abc123'));
});

test('maybe focus anilist setup window focuses existing window', () => {
  let focused = false;
  const handler = createMaybeFocusExistingAnilistSetupWindowHandler({
    getSetupWindow: () => ({
      focus: () => {
        focused = true;
      },
    }),
  });
  const handled = handler();
  assert.equal(handled, true);
  assert.equal(focused, true);
});

test('manual anilist setup submission warns on missing token', () => {
  const warnings: string[] = [];
  const handleSubmission = createHandleManualAnilistSetupSubmissionHandler({
    consumeCallbackUrl: () => false,
    redirectUri: 'https://anilist.subminer.moe/',
    logWarn: (message) => warnings.push(message),
  });

  const handled = handleSubmission('subminer://anilist-setup');
  assert.equal(handled, true);
  assert.deepEqual(warnings, ['AniList setup submission missing access token']);
});

test('anilist setup fallback handler triggers browser + manual entry on load fail', () => {
  const calls: string[] = [];
  const fallback = createAnilistSetupFallbackHandler({
    authorizeUrl: 'https://anilist.co',
    developerSettingsUrl: 'https://anilist.co/settings/developer',
    setupWindow: {
      isDestroyed: () => false,
    },
    openSetupInBrowser: () => calls.push('open-browser'),
    loadManualTokenEntry: () => calls.push('load-manual'),
    logError: () => calls.push('error'),
    logWarn: () => calls.push('warn'),
  });

  fallback.onLoadFailure({
    errorCode: -1,
    errorDescription: 'failed',
    validatedURL: 'about:blank',
  });

  assert.deepEqual(calls, ['error', 'open-browser', 'load-manual']);
});

test('anilist setup window open handler denies unsafe url', () => {
  const calls: string[] = [];
  const handler = createAnilistSetupWindowOpenHandler({
    isAllowedExternalUrl: () => false,
    openExternal: () => calls.push('open'),
    logWarn: () => calls.push('warn'),
  });

  const result = handler({ url: 'https://malicious.example' });
  assert.deepEqual(result, { action: 'deny' });
  assert.deepEqual(calls, ['warn']);
});

test('anilist setup will-navigate handler blocks callback redirect uri', () => {
  let prevented = false;
  const handler = createAnilistSetupWillNavigateHandler({
    handleManualSubmission: () => false,
    consumeCallbackUrl: () => false,
    redirectUri: 'https://anilist.subminer.moe/',
    isAllowedNavigationUrl: () => true,
    logWarn: () => {},
  });

  handler({
    url: 'https://anilist.subminer.moe/#access_token=abc',
    preventDefault: () => {
      prevented = true;
    },
  });

  assert.equal(prevented, true);
});

test('anilist setup will-navigate handler blocks unsafe urls', () => {
  const calls: string[] = [];
  let prevented = false;
  const handler = createAnilistSetupWillNavigateHandler({
    handleManualSubmission: () => false,
    consumeCallbackUrl: () => false,
    redirectUri: 'https://anilist.subminer.moe/',
    isAllowedNavigationUrl: () => false,
    logWarn: () => calls.push('warn'),
  });

  handler({
    url: 'https://unsafe.example',
    preventDefault: () => {
      prevented = true;
    },
  });

  assert.equal(prevented, true);
  assert.deepEqual(calls, ['warn']);
});

test('anilist setup will-redirect handler prevents callback redirects', () => {
  let prevented = false;
  const handler = createAnilistSetupWillRedirectHandler({
    consumeCallbackUrl: () => true,
  });

  handler({
    url: 'https://anilist.subminer.moe/#access_token=abc',
    preventDefault: () => {
      prevented = true;
    },
  });

  assert.equal(prevented, true);
});

test('anilist setup did-navigate handler consumes callback url', () => {
  const seen: string[] = [];
  const handler = createAnilistSetupDidNavigateHandler({
    consumeCallbackUrl: (url) => {
      seen.push(url);
      return true;
    },
  });

  handler('https://anilist.subminer.moe/#access_token=abc');
  assert.deepEqual(seen, ['https://anilist.subminer.moe/#access_token=abc']);
});

test('anilist setup did-fail-load handler forwards details', () => {
  const seen: Array<{ errorCode: number; errorDescription: string; validatedURL: string }> = [];
  const handler = createAnilistSetupDidFailLoadHandler({
    onLoadFailure: (details) => seen.push(details),
  });

  handler({
    errorCode: -3,
    errorDescription: 'timeout',
    validatedURL: 'https://anilist.co/api/v2/oauth/authorize',
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0].errorCode, -3);
});

test('anilist setup did-finish-load handler triggers fallback on blank page', () => {
  const calls: string[] = [];
  const handler = createAnilistSetupDidFinishLoadHandler({
    getLoadedUrl: () => 'about:blank',
    onBlankPageLoaded: () => calls.push('fallback'),
  });

  handler();
  assert.deepEqual(calls, ['fallback']);
});

test('anilist setup did-finish-load handler no-ops on non-blank page', () => {
  const calls: string[] = [];
  const handler = createAnilistSetupDidFinishLoadHandler({
    getLoadedUrl: () => 'https://anilist.co/api/v2/oauth/authorize',
    onBlankPageLoaded: () => calls.push('fallback'),
  });

  handler();
  assert.equal(calls.length, 0);
});

test('anilist setup window closed handler clears references', () => {
  const calls: string[] = [];
  const handler = createHandleAnilistSetupWindowClosedHandler({
    clearSetupWindow: () => calls.push('clear-window'),
    setSetupPageOpened: (opened) => calls.push(`opened:${opened ? 'yes' : 'no'}`),
  });

  handler();
  assert.deepEqual(calls, ['clear-window', 'opened:no']);
});

test('anilist setup window opened handler sets references', () => {
  const calls: string[] = [];
  const handler = createHandleAnilistSetupWindowOpenedHandler({
    setSetupWindow: () => calls.push('set-window'),
    setSetupPageOpened: (opened) => calls.push(`opened:${opened ? 'yes' : 'no'}`),
  });

  handler();
  assert.deepEqual(calls, ['set-window', 'opened:yes']);
});
