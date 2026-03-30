import test from 'node:test';
import assert from 'node:assert/strict';
import { createRefreshAnilistClientSecretStateHandler } from './anilist-token-refresh';

test('refresh handler marks state not_checked when tracking disabled', async () => {
  let cached: string | null = 'abc';
  let opened = true;
  const states: Array<{ status: string; source: string }> = [];
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '' } }),
    isAnilistTrackingEnabled: () => false,
    getCachedAccessToken: () => cached,
    setCachedAccessToken: (token) => {
      cached = token;
    },
    saveStoredToken: () => {},
    loadStoredToken: () => '',
    setClientSecretState: (state) => {
      states.push({ status: state.status, source: state.source });
    },
    getAnilistSetupPageOpened: () => opened,
    setAnilistSetupPageOpened: (next) => {
      opened = next;
    },
    openAnilistSetupWindow: () => {},
    now: () => 100,
  });

  const token = await refresh();
  assert.equal(token, null);
  assert.equal(cached, null);
  assert.equal(opened, false);
  assert.deepEqual(states, [{ status: 'not_checked', source: 'none' }]);
});

test('refresh handler uses literal config token and stores it', async () => {
  let cached: string | null = null;
  const saves: string[] = [];
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '  token-1  ' } }),
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => cached,
    setCachedAccessToken: (token) => {
      cached = token;
    },
    saveStoredToken: (token) => saves.push(token),
    loadStoredToken: () => '',
    setClientSecretState: () => {},
    getAnilistSetupPageOpened: () => false,
    setAnilistSetupPageOpened: () => {},
    openAnilistSetupWindow: () => {},
    now: () => 200,
  });

  const token = await refresh({ force: true });
  assert.equal(token, 'token-1');
  assert.equal(cached, 'token-1');
  assert.deepEqual(saves, ['token-1']);
});

test('refresh handler prefers cached token when not forced', async () => {
  let loadCalls = 0;
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '' } }),
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => 'cached-token',
    setCachedAccessToken: () => {},
    saveStoredToken: () => {},
    loadStoredToken: () => {
      loadCalls += 1;
      return 'stored-token';
    },
    setClientSecretState: () => {},
    getAnilistSetupPageOpened: () => false,
    setAnilistSetupPageOpened: () => {},
    openAnilistSetupWindow: () => {},
    now: () => 300,
  });

  const token = await refresh();
  assert.equal(token, 'cached-token');
  assert.equal(loadCalls, 0);
});

test('refresh handler falls back to stored token without opening setup', async () => {
  let cached: string | null = null;
  let opened = false;
  let openCalls = 0;
  const states: Array<{ status: string; source: string }> = [];
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '' } }),
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => cached,
    setCachedAccessToken: (token) => {
      cached = token;
    },
    saveStoredToken: () => {},
    loadStoredToken: () => '  stored-token  ',
    setClientSecretState: (state) => {
      states.push({ status: state.status, source: state.source });
    },
    getAnilistSetupPageOpened: () => opened,
    setAnilistSetupPageOpened: (next) => {
      opened = next;
    },
    openAnilistSetupWindow: () => {
      openCalls += 1;
    },
    now: () => 400,
  });

  const token = await refresh({ force: true });
  assert.equal(token, 'stored-token');
  assert.equal(cached, 'stored-token');
  assert.equal(opened, false);
  assert.equal(openCalls, 0);
  assert.deepEqual(states, [{ status: 'resolved', source: 'stored' }]);
});

test('refresh handler opens setup when missing token and prompting allowed', async () => {
  let cached: string | null = null;
  let opened = false;
  let openCalls = 0;
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '' } }),
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => cached,
    setCachedAccessToken: (token) => {
      cached = token;
    },
    saveStoredToken: () => {},
    loadStoredToken: () => '',
    setClientSecretState: () => {},
    getAnilistSetupPageOpened: () => opened,
    setAnilistSetupPageOpened: (next) => {
      opened = next;
    },
    openAnilistSetupWindow: () => {
      openCalls += 1;
    },
    now: () => 400,
  });

  const token = await refresh({ force: true });
  assert.equal(token, null);
  assert.equal(cached, null);
  assert.equal(openCalls, 1);
});

test('refresh handler skips setup open when missing token and prompting disabled', async () => {
  let cached: string | null = null;
  let opened = false;
  let openCalls = 0;
  const states: Array<{ status: string; source: string; message: string | null }> = [];
  const refresh = createRefreshAnilistClientSecretStateHandler({
    getResolvedConfig: () => ({ anilist: { accessToken: '' } }),
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => cached,
    setCachedAccessToken: (token) => {
      cached = token;
    },
    saveStoredToken: () => {},
    loadStoredToken: () => '',
    setClientSecretState: (state) => {
      states.push({ status: state.status, source: state.source, message: state.message });
    },
    getAnilistSetupPageOpened: () => opened,
    setAnilistSetupPageOpened: (next) => {
      opened = next;
    },
    openAnilistSetupWindow: () => {
      openCalls += 1;
    },
    now: () => 500,
  });

  const token = await refresh({ force: true, allowSetupPrompt: false });
  assert.equal(token, null);
  assert.equal(cached, null);
  assert.equal(opened, false);
  assert.equal(openCalls, 0);
  assert.deepEqual(states, [
    {
      status: 'error',
      source: 'none',
      message: 'cannot authenticate without anilist.accessToken',
    },
  ]);
});
