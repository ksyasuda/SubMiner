import test from 'node:test';
import assert from 'node:assert/strict';
import * as os from 'node:os';
import * as path from 'node:path';
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

test('jellyfin recentServers are normalized, deduped, and capped', () => {
  const { context } = createResolveContext({
    jellyfin: {
      recentServers: [
        ' http://one.local:8096/ ',
        '',
        'http://two.local:8096',
        'http://one.local:8096',
        42 as unknown as string,
        'http://three.local:8096',
        'http://four.local:8096',
        'http://five.local:8096',
        'http://six.local:8096',
      ],
    },
  });

  applyIntegrationConfig(context);

  assert.deepEqual(context.resolved.jellyfin.recentServers, [
    'http://one.local:8096',
    'http://two.local:8096',
    'http://three.local:8096',
    'http://four.local:8096',
    'http://five.local:8096',
  ]);
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

  assert.equal(context.resolved.discordPresence.enabled, true);
  assert.equal(context.resolved.discordPresence.updateIntervalMs, 3_000);
  assert.equal(context.resolved.discordPresence.debounceMs, 750);

  const warnedPaths = warnings.map((warning) => warning.path);
  assert.ok(warnedPaths.includes('discordPresence.enabled'));
  assert.ok(warnedPaths.includes('discordPresence.updateIntervalMs'));
  assert.ok(warnedPaths.includes('discordPresence.debounceMs'));
});

test('anilist character dictionary fields are parsed, clamped, and enum-validated', () => {
  const { context, warnings } = createResolveContext({
    anilist: {
      characterDictionary: {
        enabled: true,
        refreshTtlHours: 0,
        maxLoaded: 99,
        evictionPolicy: 'purge' as never,
        profileScope: 'global' as never,
        collapsibleSections: {
          description: true,
          characterInformation: 'invalid' as never,
          voicedBy: true,
        } as never,
      },
    },
  });

  applyIntegrationConfig(context);

  assert.equal(context.resolved.anilist.characterDictionary.enabled, true);
  assert.equal(context.resolved.anilist.characterDictionary.refreshTtlHours, 1);
  assert.equal(context.resolved.anilist.characterDictionary.maxLoaded, 20);
  assert.equal(context.resolved.anilist.characterDictionary.evictionPolicy, 'delete');
  assert.equal(context.resolved.anilist.characterDictionary.profileScope, 'all');
  assert.equal(context.resolved.anilist.characterDictionary.collapsibleSections.description, true);
  assert.equal(
    context.resolved.anilist.characterDictionary.collapsibleSections.characterInformation,
    false,
  );
  assert.equal(context.resolved.anilist.characterDictionary.collapsibleSections.voicedBy, true);

  const warnedPaths = warnings.map((warning) => warning.path);
  assert.ok(warnedPaths.includes('anilist.characterDictionary.refreshTtlHours'));
  assert.ok(warnedPaths.includes('anilist.characterDictionary.maxLoaded'));
  assert.ok(warnedPaths.includes('anilist.characterDictionary.evictionPolicy'));
  assert.ok(warnedPaths.includes('anilist.characterDictionary.profileScope'));
  assert.ok(
    warnedPaths.includes('anilist.characterDictionary.collapsibleSections.characterInformation'),
  );
});

test('yomitan externalProfilePath is trimmed and invalid values warn', () => {
  const { context, warnings } = createResolveContext({
    yomitan: {
      externalProfilePath: '  /tmp/gsm-profile  ',
    },
  });

  applyIntegrationConfig(context);

  assert.equal(context.resolved.yomitan.externalProfilePath, '/tmp/gsm-profile');

  const invalid = createResolveContext({
    yomitan: {
      externalProfilePath: 42 as never,
    },
  });

  applyIntegrationConfig(invalid.context);

  assert.equal(invalid.context.resolved.yomitan.externalProfilePath, '');
  assert.ok(invalid.warnings.some((warning) => warning.path === 'yomitan.externalProfilePath'));
});

test('yomitan externalProfilePath expands leading tilde to the current home directory', () => {
  const homeDir = os.homedir();
  const { context } = createResolveContext({
    yomitan: {
      externalProfilePath: '~/.config/gsm_overlay',
    },
  });

  applyIntegrationConfig(context);

  assert.equal(
    context.resolved.yomitan.externalProfilePath,
    path.join(homeDir, '.config', 'gsm_overlay'),
  );
});
