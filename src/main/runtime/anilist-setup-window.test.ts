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
  createOpenAnilistSetupWindowHandler,
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

test('open anilist setup handler no-ops when existing setup window focused', () => {
  const calls: string[] = [];
  const handler = createOpenAnilistSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => {
      calls.push('focus-existing');
      return true;
    },
    createSetupWindow: () => {
      calls.push('create-window');
      throw new Error('should not create');
    },
    buildAuthorizeUrl: () => 'https://anilist.co/api/v2/oauth/authorize?client_id=36084',
    consumeCallbackUrl: () => false,
    openSetupInBrowser: () => {},
    loadManualTokenEntry: () => {},
    redirectUri: 'https://anilist.subminer.moe/',
    developerSettingsUrl: 'https://anilist.co/settings/developer',
    isAllowedExternalUrl: () => true,
    isAllowedNavigationUrl: () => true,
    logWarn: () => {},
    logError: () => {},
    clearSetupWindow: () => {},
    setSetupPageOpened: () => {},
    setSetupWindow: () => {},
    openExternal: () => {},
  });

  handler();
  assert.deepEqual(calls, ['focus-existing']);
});

test('open anilist setup handler wires navigation, fallback, and lifecycle', () => {
  let openHandler: ((params: { url: string }) => { action: 'deny' }) | null = null;
  let willNavigateHandler: ((event: { preventDefault: () => void }, url: string) => void) | null = null;
  let didNavigateHandler: ((event: unknown, url: string) => void) | null = null;
  let didFinishLoadHandler: (() => void) | null = null;
  let didFailLoadHandler:
    | ((event: unknown, errorCode: number, errorDescription: string, validatedURL: string) => void)
    | null = null;
  let closedHandler: (() => void) | null = null;
  let prevented = false;
  const calls: string[] = [];

  const fakeWindow = {
    focus: () => {},
    webContents: {
      setWindowOpenHandler: (handler: (params: { url: string }) => { action: 'deny' }) => {
        openHandler = handler;
      },
      on: (
        event: 'will-navigate' | 'will-redirect' | 'did-navigate' | 'did-fail-load' | 'did-finish-load',
        handler: (...args: any[]) => void,
      ) => {
        if (event === 'will-navigate') willNavigateHandler = handler as never;
        if (event === 'did-navigate') didNavigateHandler = handler as never;
        if (event === 'did-finish-load') didFinishLoadHandler = handler as never;
        if (event === 'did-fail-load') didFailLoadHandler = handler as never;
      },
      getURL: () => 'about:blank',
    },
    on: (event: 'closed', handler: () => void) => {
      if (event === 'closed') closedHandler = handler;
    },
    isDestroyed: () => false,
  };

  const handler = createOpenAnilistSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => fakeWindow,
    buildAuthorizeUrl: () => 'https://anilist.co/api/v2/oauth/authorize?client_id=36084',
    consumeCallbackUrl: (rawUrl) => {
      calls.push(`consume:${rawUrl}`);
      return rawUrl.includes('access_token=');
    },
    openSetupInBrowser: () => calls.push('open-browser'),
    loadManualTokenEntry: () => calls.push('load-manual'),
    redirectUri: 'https://anilist.subminer.moe/',
    developerSettingsUrl: 'https://anilist.co/settings/developer',
    isAllowedExternalUrl: () => true,
    isAllowedNavigationUrl: () => true,
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
    clearSetupWindow: () => calls.push('clear-window'),
    setSetupPageOpened: (opened) => calls.push(`opened:${opened ? 'yes' : 'no'}`),
    setSetupWindow: () => calls.push('set-window'),
    openExternal: (url) => calls.push(`open:${url}`),
  });

  handler();
  assert.ok(openHandler);
  assert.ok(willNavigateHandler);
  assert.ok(didNavigateHandler);
  assert.ok(didFinishLoadHandler);
  assert.ok(didFailLoadHandler);
  assert.ok(closedHandler);
  assert.deepEqual(calls.slice(0, 3), ['load-manual', 'set-window', 'opened:yes']);

  const onOpen = openHandler as ((params: { url: string }) => { action: 'deny' }) | null;
  if (!onOpen) throw new Error('missing window open handler');
  assert.deepEqual(onOpen({ url: 'https://anilist.co/settings/developer' }), { action: 'deny' });
  assert.ok(calls.includes('open:https://anilist.co/settings/developer'));

  const onWillNavigate = willNavigateHandler as
    | ((event: { preventDefault: () => void }, url: string) => void)
    | null;
  if (!onWillNavigate) throw new Error('missing will navigate handler');
  onWillNavigate(
    {
      preventDefault: () => {
        prevented = true;
      },
    },
    'https://anilist.subminer.moe/#access_token=abc',
  );
  assert.equal(prevented, true);

  const onDidNavigate = didNavigateHandler as ((event: unknown, url: string) => void) | null;
  if (!onDidNavigate) throw new Error('missing did navigate handler');
  onDidNavigate({}, 'https://anilist.subminer.moe/#access_token=abc');

  const onDidFinishLoad = didFinishLoadHandler as (() => void) | null;
  if (!onDidFinishLoad) throw new Error('missing did finish load handler');
  onDidFinishLoad();
  assert.ok(calls.includes('warn:AniList setup loaded a blank page; using fallback'));
  assert.ok(calls.includes('open-browser'));

  const onDidFailLoad = didFailLoadHandler as
    | ((event: unknown, errorCode: number, errorDescription: string, validatedURL: string) => void)
    | null;
  if (!onDidFailLoad) throw new Error('missing did fail load handler');
  onDidFailLoad({}, -1, 'load failed', 'about:blank');
  assert.ok(calls.includes('error:AniList setup window failed to load'));

  const onClosed = closedHandler as (() => void) | null;
  if (!onClosed) throw new Error('missing closed handler');
  onClosed();
  assert.ok(calls.includes('clear-window'));
  assert.ok(calls.includes('opened:no'));
});
