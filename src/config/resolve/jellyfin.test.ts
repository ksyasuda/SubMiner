import test from 'node:test';
import assert from 'node:assert/strict';
import { createResolveContext } from './context';
import { applyIntegrationConfig } from './integrations';

test('jellyfin directPlayContainers are normalized', () => {
  const { context } = createResolveContext({
    jellyfin: {
      directPlayContainers: [' MKV ', 'mp4', '', ' WebM  ', 42 as unknown as string],
    },
  });

  applyIntegrationConfig(context);

  assert.deepEqual(context.resolved.jellyfin.directPlayContainers, ['mkv', 'mp4', 'webm']);
});

test('jellyfin legacy auth keys are ignored by resolver', () => {
  const { context } = createResolveContext({
    jellyfin: { accessToken: 'legacy-token', userId: 'legacy-user' } as unknown as never,
  });

  applyIntegrationConfig(context);

  assert.equal('accessToken' in (context.resolved.jellyfin as Record<string, unknown>), false);
  assert.equal('userId' in (context.resolved.jellyfin as Record<string, unknown>), false);
});

test('discordPresence fields are parsed and clamped', () => {
  const { context } = createResolveContext({
    discordPresence: {
      enabled: true,
      updateIntervalMs: 500,
      debounceMs: -100,
    },
  });

  applyIntegrationConfig(context);

  assert.equal(context.resolved.discordPresence.enabled, true);
  assert.equal(context.resolved.discordPresence.updateIntervalMs, 1000);
  assert.equal(context.resolved.discordPresence.debounceMs, 0);
});

test('discordPresence invalid values warn and keep defaults', () => {
  const { context, warnings } = createResolveContext({
    discordPresence: {
      enabled: 'true' as never,
      updateIntervalMs: 'fast' as never,
      debounceMs: null as never,
    },
  });

  applyIntegrationConfig(context);

  assert.equal(context.resolved.discordPresence.enabled, false);
  assert.equal(context.resolved.discordPresence.updateIntervalMs, 3_000);
  assert.equal(context.resolved.discordPresence.debounceMs, 750);

  const warnedPaths = warnings.map((warning) => warning.path);
  assert.ok(warnedPaths.includes('discordPresence.enabled'));
  assert.ok(warnedPaths.includes('discordPresence.updateIntervalMs'));
  assert.ok(warnedPaths.includes('discordPresence.debounceMs'));
});
