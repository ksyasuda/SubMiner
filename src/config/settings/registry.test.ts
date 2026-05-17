import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG } from '../definitions';
import {
  buildConfigSettingsRegistry,
  getConfigSettingsCoverage,
  LEGACY_HIDDEN_CONFIG_PATHS,
} from './registry';

test('config settings registry places hover pause under viewing playback behavior', () => {
  const fields = buildConfigSettingsRegistry(DEFAULT_CONFIG);
  const hoverPause = fields.find(
    (field) => field.configPath === 'subtitleStyle.autoPauseVideoOnHover',
  );

  assert.ok(hoverPause);
  assert.equal(hoverPause.category, 'viewing');
  assert.equal(hoverPause.section, 'Playback pause behavior');
  assert.equal(hoverPause.control, 'boolean');
});

test('config settings registry hides legacy and ignored paths from normal fields', () => {
  const fields = buildConfigSettingsRegistry(DEFAULT_CONFIG);
  const visiblePaths = new Set(
    fields.filter((field) => !field.legacyHidden).map((field) => field.configPath),
  );

  for (const path of LEGACY_HIDDEN_CONFIG_PATHS) {
    assert.equal(visiblePaths.has(path), false, path);
  }
  assert.equal(visiblePaths.has('controller.buttonIndices'), false);
});

test('config settings registry covers canonical defaults or marks explicit raw-only gaps', () => {
  const fields = buildConfigSettingsRegistry(DEFAULT_CONFIG);
  const coverage = getConfigSettingsCoverage(DEFAULT_CONFIG, fields);

  assert.deepEqual(coverage.uncoveredDefaultPaths, []);
});
