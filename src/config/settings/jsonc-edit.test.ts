import test from 'node:test';
import assert from 'node:assert/strict';
import { parse } from 'jsonc-parser';
import { DEFAULT_CONFIG } from '../definitions';
import { applyConfigSettingsPatchToContent, buildConfigSettingsSnapshot } from './jsonc-edit';
import { buildConfigSettingsRegistry } from './registry';

test('applyConfigSettingsPatchToContent preserves JSONC comments while setting nested values', () => {
  const input = `{
  // keep this comment
  "subtitleStyle": {
    "fontSize": 35,
  },
}`;

  const result = applyConfigSettingsPatchToContent({
    content: input,
    operations: [
      {
        op: 'set',
        path: 'subtitleStyle.autoPauseVideoOnHover',
        value: false,
      },
    ],
    previousWarnings: [],
  });

  assert.equal(result.ok, true);
  assert.match(result.content, /keep this comment/);
  const parsed = parse(result.content);
  assert.equal(parsed.subtitleStyle.autoPauseVideoOnHover, false);
  assert.equal(parsed.subtitleStyle.fontSize, 35);
});

test('applyConfigSettingsPatchToContent updates effective duplicate object path', () => {
  const input = `{
  "ankiConnect": {
    "nPlusOne": {
      "enabled": true
    },
    "knownWords": {
      "highlightEnabled": true
    },
    "nPlusOne": {
      "minSentenceWords": 3
    }
  }
}`;

  const result = applyConfigSettingsPatchToContent({
    content: input,
    operations: [
      {
        op: 'set',
        path: 'ankiConnect.nPlusOne.enabled',
        value: true,
      },
    ],
    previousWarnings: [],
  });

  assert.equal(result.ok, true);
  const parsed = parse(result.content);
  assert.equal(parsed.ankiConnect.nPlusOne.enabled, true);
  assert.equal(parsed.ankiConnect.nPlusOne.minSentenceWords, 3);
});

test('applyConfigSettingsPatchToContent removes duplicate properties across JSONC trivia', () => {
  const input = `{
  "ankiConnect": {
    "nPlusOne": {
      "enabled": false
    } /* old value */ ,
    // effective value follows
    "nPlusOne": {
      "minSentenceWords": 3
    }
  }
}`;

  const result = applyConfigSettingsPatchToContent({
    content: input,
    operations: [
      {
        op: 'set',
        path: 'ankiConnect.nPlusOne.enabled',
        value: true,
      },
    ],
    previousWarnings: [],
  });

  assert.equal(result.ok, true);
  const parsed = parse(result.content);
  assert.equal(parsed.ankiConnect.nPlusOne.enabled, true);
  assert.equal(parsed.ankiConnect.nPlusOne.minSentenceWords, 3);
});

test('applyConfigSettingsPatchToContent reset removes explicit path', () => {
  const input = `{
  "subtitleStyle": {
    "fontSize": 41,
    "autoPauseVideoOnHover": false
  }
}`;

  const result = applyConfigSettingsPatchToContent({
    content: input,
    operations: [{ op: 'reset', path: 'subtitleStyle.autoPauseVideoOnHover' }],
    previousWarnings: [],
  });

  assert.equal(result.ok, true);
  const parsed = parse(result.content);
  assert.equal(Object.hasOwn(parsed.subtitleStyle, 'autoPauseVideoOnHover'), false);
  assert.equal(parsed.subtitleStyle.fontSize, 41);
});

test('applyConfigSettingsPatchToContent rejects warnings caused by modified fields', () => {
  const result = applyConfigSettingsPatchToContent({
    content: '{}',
    operations: [
      {
        op: 'set',
        path: 'subtitleStyle.autoPauseVideoOnHover',
        value: 'bad',
      },
    ],
    previousWarnings: [],
  });

  assert.equal(result.ok, false);
  assert.equal(result.warnings[0]?.path, 'subtitleStyle.autoPauseVideoOnHover');
});

test('buildConfigSettingsSnapshot masks configured secret values', () => {
  const fields = buildConfigSettingsRegistry(DEFAULT_CONFIG);
  const snapshot = buildConfigSettingsSnapshot({
    configPath: '/tmp/config.jsonc',
    rawConfig: {
      ai: {
        apiKey: 'secret-key',
      },
    },
    resolvedConfig: {
      ...DEFAULT_CONFIG,
      ai: {
        ...DEFAULT_CONFIG.ai,
        apiKey: 'secret-key',
      },
    },
    warnings: [],
    fields,
  });

  const apiKey = snapshot.values['ai.apiKey'];
  assert.deepEqual(apiKey, { configured: true });
});
