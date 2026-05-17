import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG, type ReloadConfigStrictResult } from '../../config';
import type { ResolvedConfig } from '../../types/config';
import type { ConfigSettingsSnapshot } from '../../types/settings';
import { createSaveConfigSettingsPatchHandler } from './config-settings-save';

function snapshot(): ConfigSettingsSnapshot {
  return {
    configPath: '/tmp/config.jsonc',
    fields: [],
    values: {},
    warnings: [],
  };
}

test('config settings save applies hot-reloadable diff live', () => {
  const calls: string[] = [];
  const previous = DEFAULT_CONFIG;
  const next: ResolvedConfig = {
    ...DEFAULT_CONFIG,
    subtitleStyle: {
      ...DEFAULT_CONFIG.subtitleStyle,
      autoPauseVideoOnHover: false,
    },
  };
  let written = '';
  const save = createSaveConfigSettingsPatchHandler({
    getConfigPath: () => '/tmp/config.jsonc',
    getCurrentConfig: () => previous,
    getWarnings: () => [],
    getSnapshot: () => snapshot(),
    fileExists: () => true,
    readText: () => '{}',
    writeTextAtomically: (_path, content) => {
      written = content;
      calls.push('write');
    },
    reloadConfigStrict: (): ReloadConfigStrictResult => ({
      ok: true,
      config: next,
      warnings: [],
      path: '/tmp/config.jsonc',
    }),
    classifyDiff: () => ({
      hotReloadFields: ['subtitleStyle'],
      restartRequiredFields: [],
    }),
    applyHotReload: (diff) => calls.push(`hot:${diff.hotReloadFields.join(',')}`),
    getRestartRequiredSections: () => [],
  });

  const result = save({
    operations: [
      {
        op: 'set',
        path: 'subtitleStyle.autoPauseVideoOnHover',
        value: false,
      },
    ],
  });

  assert.equal(result.ok, true);
  assert.match(written, /autoPauseVideoOnHover/);
  assert.deepEqual(calls, ['write', 'hot:subtitleStyle']);
  assert.deepEqual(result.hotReloadFields, ['subtitleStyle']);
  assert.deepEqual(result.restartRequiredFields, []);
});

test('config settings save returns restart-required sections without applying hot reload', () => {
  const calls: string[] = [];
  const previous = DEFAULT_CONFIG;
  const next: ResolvedConfig = {
    ...DEFAULT_CONFIG,
    mpv: {
      ...DEFAULT_CONFIG.mpv,
      launchMode: 'fullscreen',
    },
  };
  const save = createSaveConfigSettingsPatchHandler({
    getConfigPath: () => '/tmp/config.jsonc',
    getCurrentConfig: () => previous,
    getWarnings: () => [],
    getSnapshot: () => snapshot(),
    fileExists: () => true,
    readText: () => '{}',
    writeTextAtomically: () => calls.push('write'),
    reloadConfigStrict: (): ReloadConfigStrictResult => ({
      ok: true,
      config: next,
      warnings: [],
      path: '/tmp/config.jsonc',
    }),
    classifyDiff: () => ({
      hotReloadFields: [],
      restartRequiredFields: ['mpv'],
    }),
    applyHotReload: () => calls.push('hot'),
    getRestartRequiredSections: () => ['mpv launcher'],
  });

  const result = save({
    operations: [{ op: 'set', path: 'mpv.launchMode', value: 'fullscreen' }],
  });

  assert.equal(result.ok, true);
  assert.deepEqual(calls, ['write']);
  assert.deepEqual(result.hotReloadFields, []);
  assert.deepEqual(result.restartRequiredFields, ['mpv']);
  assert.deepEqual(result.restartRequiredSections, ['mpv launcher']);
});

test('config settings save restores previous file content when strict reload fails', () => {
  const writes: string[] = [];
  const save = createSaveConfigSettingsPatchHandler({
    getConfigPath: () => '/tmp/config.jsonc',
    getCurrentConfig: () => DEFAULT_CONFIG,
    getWarnings: () => [],
    getSnapshot: () => snapshot(),
    fileExists: () => true,
    readText: () => '{"mpv":{"launchMode":"normal"}}\n',
    writeTextAtomically: (_path, content) => {
      writes.push(content);
    },
    reloadConfigStrict: (): ReloadConfigStrictResult => ({
      ok: false,
      error: 'invalid config',
      path: '/tmp/config.jsonc',
    }),
    classifyDiff: () => {
      throw new Error('Should not classify invalid config.');
    },
    applyHotReload: () => {
      throw new Error('Should not hot reload invalid config.');
    },
    getRestartRequiredSections: () => [],
  });

  const result = save({
    operations: [{ op: 'set', path: 'mpv.launchMode', value: 'fullscreen' }],
  });

  assert.equal(result.ok, false);
  assert.equal(result.error, 'invalid config');
  assert.equal(writes.length, 2);
  assert.match(writes[0] ?? '', /fullscreen/);
  assert.equal(writes[1], '{"mpv":{"launchMode":"normal"}}\n');
});
