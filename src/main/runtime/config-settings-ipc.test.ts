import test from 'node:test';
import assert from 'node:assert/strict';
import { isConfigSettingsPatch } from './config-settings-ipc';
import type { ConfigSettingsField } from '../../types/settings';

const fields: ConfigSettingsField[] = [
  {
    id: 'mpv.launchMode',
    label: 'Launch mode',
    description: 'Launch mode setting.',
    configPath: 'mpv.launchMode',
    category: 'behavior',
    section: 'mpv Playback',
    control: 'select',
    defaultValue: 'windowed',
    restartBehavior: 'restart',
  },
];

test('isConfigSettingsPatch rejects set operations without a value property', () => {
  assert.equal(
    isConfigSettingsPatch(
      {
        operations: [{ op: 'set', path: 'mpv.launchMode' }],
      },
      fields,
    ),
    false,
  );
});

test('isConfigSettingsPatch accepts set operations with an explicit value', () => {
  assert.equal(
    isConfigSettingsPatch(
      {
        operations: [{ op: 'set', path: 'mpv.launchMode', value: 'fullscreen' }],
      },
      fields,
    ),
    true,
  );
});

test('isConfigSettingsPatch accepts reset operations without a value', () => {
  assert.equal(
    isConfigSettingsPatch(
      {
        operations: [{ op: 'reset', path: 'mpv.launchMode' }],
      },
      fields,
    ),
    true,
  );
});

test('isConfigSettingsPatch rejects unknown config paths', () => {
  assert.equal(
    isConfigSettingsPatch(
      {
        operations: [{ op: 'reset', path: 'unknown.path' }],
      },
      fields,
    ),
    false,
  );
});
