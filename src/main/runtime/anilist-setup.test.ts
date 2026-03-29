import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildAnilistSetupFallbackHtml,
  buildAnilistManualTokenEntryHtml,
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  extractAnilistAccessTokenFromUrl,
  findAnilistSetupDeepLinkArgvUrl,
} from './anilist-setup';

test('buildAnilistSetupUrl includes required query params', () => {
  const url = buildAnilistSetupUrl({
    authorizeUrl: 'https://anilist.co/api/v2/oauth/authorize',
    clientId: '36084',
    responseType: 'token',
    redirectUri: 'https://anilist.subminer.moe/',
  });
  assert.match(url, /client_id=36084/);
  assert.match(url, /response_type=token/);
  assert.match(url, /redirect_uri=https%3A%2F%2Fanilist\.subminer\.moe%2F/);
});

test('buildAnilistSetupUrl omits redirect_uri when unset', () => {
  const url = buildAnilistSetupUrl({
    authorizeUrl: 'https://anilist.co/api/v2/oauth/authorize',
    clientId: '36084',
    responseType: 'token',
  });
  assert.match(url, /client_id=36084/);
  assert.match(url, /response_type=token/);
  assert.equal(url.includes('redirect_uri='), false);
});

test('buildAnilistSetupFallbackHtml escapes reason content', () => {
  const html = buildAnilistSetupFallbackHtml({
    reason: '<script>alert(1)</script>',
    authorizeUrl: 'https://anilist.example/auth',
    developerSettingsUrl: 'https://anilist.example/dev',
  });
  assert.equal(html.includes('<script>alert(1)</script>'), false);
  assert.match(html, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
});

test('buildAnilistManualTokenEntryHtml includes access-token submit route only', () => {
  const html = buildAnilistManualTokenEntryHtml({
    authorizeUrl: 'https://anilist.example/auth',
    developerSettingsUrl: 'https://anilist.example/dev',
  });
  assert.match(html, /subminer:\/\/anilist-setup\?access_token=/);
  assert.equal(html.includes('callback_url='), false);
  assert.equal(html.includes('subminer://anilist-setup?code='), false);
});

test('extractAnilistAccessTokenFromUrl returns access token from hash fragment', () => {
  const token = extractAnilistAccessTokenFromUrl(
    'https://anilist.subminer.moe/#access_token=token-from-hash&token_type=Bearer',
  );
  assert.equal(token, 'token-from-hash');
});

test('extractAnilistAccessTokenFromUrl returns access token from query', () => {
  const token = extractAnilistAccessTokenFromUrl(
    'https://anilist.subminer.moe/?access_token=token-from-query&token_type=Bearer',
  );
  assert.equal(token, 'token-from-query');
});

test('findAnilistSetupDeepLinkArgvUrl finds subminer deep link from argv', () => {
  const rawUrl = findAnilistSetupDeepLinkArgvUrl([
    '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    '--start',
    'subminer://anilist-setup?access_token=argv-token',
  ]);
  assert.equal(rawUrl, 'subminer://anilist-setup?access_token=argv-token');
});

test('findAnilistSetupDeepLinkArgvUrl returns null when missing', () => {
  const rawUrl = findAnilistSetupDeepLinkArgvUrl([
    '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    '--start',
  ]);
  assert.equal(rawUrl, null);
});

test('consumeAnilistSetupCallbackUrl persists token and closes window for callback URL', () => {
  const originalDateNow = Date.now;
  const events: string[] = [];
  try {
    Date.now = () => 120_000;
    const handled = consumeAnilistSetupCallbackUrl({
      rawUrl: 'https://anilist.subminer.moe/#access_token=saved-token',
      saveToken: (value: string) => events.push(`save:${value}`),
      setCachedToken: (value: string) => events.push(`cache:${value}`),
      setResolvedState: (timestampMs: number) =>
        events.push(`state:${timestampMs > 0 ? 'ok' : 'bad'}`),
      setSetupPageOpened: (opened: boolean) => events.push(`opened:${opened}`),
      onSuccess: () => events.push('success'),
      closeWindow: () => events.push('close'),
    });

    assert.equal(handled, true);
    assert.deepEqual(events, [
      'save:saved-token',
      'cache:saved-token',
      'state:ok',
      'opened:false',
      'success',
      'close',
    ]);
  } finally {
    Date.now = originalDateNow;
  }
});

test('consumeAnilistSetupCallbackUrl persists token for subminer deep link URL', () => {
  const originalDateNow = Date.now;
  const events: string[] = [];
  try {
    Date.now = () => 120_000;
    const handled = consumeAnilistSetupCallbackUrl({
      rawUrl: 'subminer://anilist-setup?access_token=saved-token',
      saveToken: (value: string) => events.push(`save:${value}`),
      setCachedToken: (value: string) => events.push(`cache:${value}`),
      setResolvedState: (timestampMs: number) =>
        events.push(`state:${timestampMs > 0 ? 'ok' : 'bad'}`),
      setSetupPageOpened: (opened: boolean) => events.push(`opened:${opened}`),
      onSuccess: () => events.push('success'),
      closeWindow: () => events.push('close'),
    });

    assert.equal(handled, true);
    assert.deepEqual(events, [
      'save:saved-token',
      'cache:saved-token',
      'state:ok',
      'opened:false',
      'success',
      'close',
    ]);
  } finally {
    Date.now = originalDateNow;
  }
});

test('consumeAnilistSetupCallbackUrl ignores non-callback URLs', () => {
  const events: string[] = [];
  const handled = consumeAnilistSetupCallbackUrl({
    rawUrl: 'https://anilist.co/settings/developer',
    saveToken: () => events.push('save'),
    setCachedToken: () => events.push('cache'),
    setResolvedState: () => events.push('state'),
    setSetupPageOpened: () => events.push('opened'),
    onSuccess: () => events.push('success'),
    closeWindow: () => events.push('close'),
  });

  assert.equal(handled, false);
  assert.deepEqual(events, []);
});
