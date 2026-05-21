import assert from 'node:assert/strict';
import test from 'node:test';
import type { ConfigSettingsField } from '../types/settings';
import { getFieldTitleBadges } from './settings-field-layout';

const advancedRestartField: ConfigSettingsField = {
  id: 'ankiConnect.knownWords.highlightEnabled',
  label: 'Enabled',
  description: 'Enable fast local highlighting for words already known in Anki.',
  configPath: 'ankiConnect.knownWords.highlightEnabled',
  category: 'mining-anki',
  section: 'Known Words',
  control: 'boolean',
  defaultValue: false,
  restartBehavior: 'restart',
  advanced: true,
};

test('field title badges show restart status without config paths or advanced labels', () => {
  const badges = getFieldTitleBadges(advancedRestartField);

  assert.deepEqual(badges, [
    {
      className: 'restart-chip restart',
      text: 'Restart',
    },
  ]);
  assert.equal(JSON.stringify(badges).includes(advancedRestartField.configPath), false);
  assert.equal(JSON.stringify(badges).includes('Advanced'), false);
});
