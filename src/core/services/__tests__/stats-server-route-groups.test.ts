import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import {
  registerStatsAnalyticsRoutes,
  registerStatsIntegrationRoutes,
  registerStatsLibraryRoutes,
  registerStatsMiningRoutes,
  registerStatsStaticRoutes,
} from '../stats-server/routes.js';

describe('stats server route groups', () => {
  test('exposes focused registrars for createStatsApp composition', () => {
    for (const registrar of [
      registerStatsAnalyticsRoutes,
      registerStatsLibraryRoutes,
      registerStatsIntegrationRoutes,
      registerStatsMiningRoutes,
      registerStatsStaticRoutes,
    ]) {
      assert.equal(typeof registrar, 'function');
    }
  });
});
