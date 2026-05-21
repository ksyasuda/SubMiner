import assert from 'node:assert/strict';
import test from 'node:test';
import { parseArgs } from '../../cli/args';
import {
  getStartupModeFlags,
  shouldRefreshAnilistOnConfigReload,
  shouldStartAutomaticUpdateChecks,
} from './startup-mode-flags';

test('settings window startup uses minimal startup and skips background integrations', () => {
  const args = parseArgs(['--settings']);
  const flags = getStartupModeFlags(args);

  assert.equal(flags.shouldUseMinimalStartup, true);
  assert.equal(flags.shouldSkipHeavyStartup, true);
  assert.equal(shouldRefreshAnilistOnConfigReload(args), false);
  assert.equal(shouldStartAutomaticUpdateChecks(args), false);
});

test('normal startup still allows background integrations', () => {
  const flags = getStartupModeFlags(null);

  assert.equal(flags.shouldUseMinimalStartup, false);
  assert.equal(flags.shouldSkipHeavyStartup, false);
  assert.equal(shouldRefreshAnilistOnConfigReload(null), true);
  assert.equal(shouldStartAutomaticUpdateChecks(null), true);
});
