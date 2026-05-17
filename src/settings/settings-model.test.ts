import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createSettingsDraft,
  filterSettingsFields,
  setDraftValue,
  resetDraftPath,
  getDirtyOperations,
} from './settings-model';
import type { ConfigSettingsField } from '../types/settings';

const fields: ConfigSettingsField[] = [
  {
    id: 'subtitleStyle.autoPauseVideoOnHover',
    label: 'Pause on subtitle hover',
    description: 'Pause while hovering subtitles.',
    configPath: 'subtitleStyle.autoPauseVideoOnHover',
    category: 'behavior',
    section: 'Playback Pause Behavior',
    control: 'boolean',
    defaultValue: true,
    restartBehavior: 'hot-reload',
  },
  {
    id: 'ankiConnect.enabled',
    label: 'Enable AnkiConnect',
    description: 'Enable Anki integration.',
    configPath: 'ankiConnect.enabled',
    category: 'mining-anki',
    section: 'Connection',
    control: 'boolean',
    defaultValue: true,
    restartBehavior: 'restart',
  },
  {
    id: 'subtitleStyle.fontSize',
    label: 'Font Size',
    description: 'Hidden behind CSS editor.',
    configPath: 'subtitleStyle.fontSize',
    category: 'appearance',
    section: 'Primary Subtitle Appearance',
    control: 'number',
    defaultValue: 35,
    restartBehavior: 'hot-reload',
    settingsHidden: true,
  },
];

test('filterSettingsFields searches label, section, and config path', () => {
  assert.deepEqual(
    filterSettingsFields(fields, { category: 'behavior', query: 'hover' }).map(
      (field) => field.configPath,
    ),
    ['subtitleStyle.autoPauseVideoOnHover'],
  );
  assert.deepEqual(filterSettingsFields(fields, { category: 'behavior', query: 'anki' }), []);
  assert.deepEqual(
    filterSettingsFields(fields, { query: 'anki' }).map((field) => field.configPath),
    ['ankiConnect.enabled'],
  );
  assert.deepEqual(
    filterSettingsFields(fields, { query: 'pause hover' }).map((field) => field.configPath),
    ['subtitleStyle.autoPauseVideoOnHover'],
  );
  assert.deepEqual(filterSettingsFields(fields, { query: 'pause anki' }), []);
  assert.deepEqual(filterSettingsFields(fields, { category: 'appearance', query: '' }), []);
});

test('settings draft tracks dirty set and emits save operations', () => {
  const draft = createSettingsDraft({
    'subtitleStyle.autoPauseVideoOnHover': true,
  });

  setDraftValue(draft, 'subtitleStyle.autoPauseVideoOnHover', false);
  assert.deepEqual(getDirtyOperations(draft), [
    {
      op: 'set',
      path: 'subtitleStyle.autoPauseVideoOnHover',
      value: false,
    },
  ]);

  setDraftValue(draft, 'subtitleStyle.autoPauseVideoOnHover', true);
  assert.deepEqual(getDirtyOperations(draft), []);
});

test('settings draft emits reset operations for css-editor-owned legacy style paths', () => {
  const draft = createSettingsDraft({
    'subtitleStyle.css': {},
    'subtitleStyle.fontSize': 35,
  });

  setDraftValue(draft, 'subtitleStyle.css', { 'font-size': '42px' });
  resetDraftPath(draft, 'subtitleStyle.fontSize', undefined);

  assert.deepEqual(getDirtyOperations(draft), [
    {
      op: 'set',
      path: 'subtitleStyle.css',
      value: { 'font-size': '42px' },
    },
    {
      op: 'reset',
      path: 'subtitleStyle.fontSize',
    },
  ]);
});
