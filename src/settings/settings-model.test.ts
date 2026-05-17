import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createSettingsDraft,
  filterSettingsFields,
  setDraftValue,
  getDirtyOperations,
} from './settings-model';
import type { ConfigSettingsField } from '../types/settings';

const fields: ConfigSettingsField[] = [
  {
    id: 'subtitleStyle.autoPauseVideoOnHover',
    label: 'Pause on subtitle hover',
    description: 'Pause while hovering subtitles.',
    configPath: 'subtitleStyle.autoPauseVideoOnHover',
    category: 'viewing',
    section: 'Playback pause behavior',
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
];

test('filterSettingsFields searches label, section, and config path', () => {
  assert.deepEqual(
    filterSettingsFields(fields, { category: 'viewing', query: 'hover' }).map(
      (field) => field.configPath,
    ),
    ['subtitleStyle.autoPauseVideoOnHover'],
  );
  assert.deepEqual(filterSettingsFields(fields, { category: 'viewing', query: 'anki' }), []);
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
