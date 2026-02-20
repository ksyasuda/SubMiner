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
    loadStoredToken: () => 'stored-token',
  })();
  assert.equal(deps.getResolvedConfig(), resolved);
  assert.equal(deps.loadStoredToken(), 'stored-token');
});

test('get jellyfin client info main deps builder maps callbacks', () => {
  const configured = { clientName: 'Configured' };
  const defaults = { clientName: 'Default' };
  const deps = createBuildGetJellyfinClientInfoMainDepsHandler({
    getResolvedJellyfinConfig: () => configured as never,
    getDefaultJellyfinConfig: () => defaults as never,
  })();

  assert.equal(deps.getResolvedJellyfinConfig(), configured);
  assert.equal(deps.getDefaultJellyfinConfig(), defaults);
});
