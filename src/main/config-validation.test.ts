import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildConfigWarningNotificationBody,
  buildConfigWarningSummary,
  failStartupFromConfig,
  formatConfigValue,
} from './config-validation';

test('formatConfigValue handles undefined and JSON values', () => {
  assert.equal(formatConfigValue(undefined), 'undefined');
  assert.equal(formatConfigValue({ x: 1 }), '{"x":1}');
  assert.equal(formatConfigValue(['a', 2]), '["a",2]');
});

test('buildConfigWarningSummary includes warnings with formatted values', () => {
  const summary = buildConfigWarningSummary('/tmp/config.jsonc', [
    {
      path: 'ankiConnect.pollingRate',
      message: 'must be >= 50',
      value: 20,
      fallback: 250,
    },
  ]);

  assert.match(summary, /Validation found 1 issue\(s\)\. File: \/tmp\/config\.jsonc/);
  assert.match(summary, /ankiConnect\.pollingRate: must be >= 50 actual=20 fallback=250/);
});

test('buildConfigWarningNotificationBody includes concise warning details', () => {
  const body = buildConfigWarningNotificationBody('/tmp/config.jsonc', [
    {
      path: 'ankiConnect.ai',
      message: 'Expected boolean.',
      value: { enabled: true },
      fallback: false,
    },
    {
      path: 'ankiConnect.isLapis.sentenceCardSentenceField',
      message: 'Deprecated key; sentence-card sentence field is fixed to Sentence.',
      value: 'Sentence',
      fallback: 'Sentence',
    },
  ]);

  assert.match(body, /2 config validation issue\(s\) detected\./);
  assert.match(body, /File: \/tmp\/config\.jsonc/);
  assert.match(body, /1\. ankiConnect\.ai: Expected boolean\./);
  assert.match(
    body,
    /2\. ankiConnect\.isLapis\.sentenceCardSentenceField: Deprecated key; sentence-card sentence field is fixed to Sentence\./,
  );
});

test('buildConfigWarningNotificationBody caps listed warnings and clips long config paths', () => {
  const longPath = `/home/user/${'nested/'.repeat(10)}config.jsonc`;
  const warnings = [1, 2, 3, 4, 5].map((n) => ({
    path: `key${n}`,
    message: 'invalid',
    value: n,
    fallback: 0,
  }));

  const lines = buildConfigWarningNotificationBody(longPath, warnings).split('\n');

  assert.equal(lines[0], '5 config validation issue(s) detected.');
  const fileLine = lines.find((line) => line.startsWith('File: '));
  assert.equal(fileLine, `File: ...${longPath.slice(-45)}`);
  assert.equal(fileLine?.length, 'File: '.length + 48);
  assert.deepEqual(lines.slice(-4), [
    '1. key1: invalid',
    '2. key2: invalid',
    '3. key3: invalid',
    '+2 more issue(s)',
  ]);
});

test('failStartupFromConfig invokes handlers and throws', () => {
  const calls: string[] = [];
  const exitCodes: number[] = [];

  assert.throws(
    () =>
      failStartupFromConfig('Config Error', 'bad value', {
        logError: (details) => {
          calls.push(`log:${details}`);
        },
        showErrorBox: (title, details) => {
          calls.push(`dialog:${title}:${details}`);
        },
        setExitCode: (code) => {
          exitCodes.push(code);
        },
        quit: () => {
          calls.push('quit');
        },
      }),
    /bad value/,
  );

  assert.deepEqual(exitCodes, [1]);
  assert.deepEqual(calls, ['log:bad value', 'dialog:Config Error:bad value', 'quit']);
});
