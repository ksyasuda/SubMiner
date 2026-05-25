import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildGetJellyfinClientInfoMainDepsHandler,
  createBuildGetResolvedJellyfinConfigMainDepsHandler,
} from './jellyfin-client-info-main-deps';

test('get resolved jellyfin config main deps builder maps callbacks', () => {
  const resolved = { jellyfin: { url: 'https://example.com' } };
  const deps = createBuildGetResolvedJellyfinConfigMainDepsHandler({
    getResolvedConfig: () => resolved as never,
    loadStoredSession: () => ({ accessToken: 'stored-token', userId: 'uid' }),
    getEnv: (key: string) => (key === 'TEST' ? 'x' : undefined),
  })();
  assert.equal(deps.getResolvedConfig(), resolved);
  assert.deepEqual(deps.loadStoredSession(), { accessToken: 'stored-token', userId: 'uid' });
  assert.equal(deps.getEnv('TEST'), 'x');
});

test('get jellyfin client info main deps builder maps callbacks', () => {
  const configured = { clientName: 'Configured' };
  const deps = createBuildGetJellyfinClientInfoMainDepsHandler({
    getResolvedJellyfinConfig: () => configured as never,
    getHostName: () => 'workstation',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0.0',
  })();

  assert.equal(deps.getResolvedJellyfinConfig(), configured);
  assert.equal(deps.getHostName?.(), 'workstation');
  assert.equal(deps.defaultClientName, 'SubMiner');
  assert.equal(deps.defaultClientVersion, '1.0.0');
});
