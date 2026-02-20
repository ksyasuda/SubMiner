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
      path: 'ankiConnect.openRouter',
      message: 'Deprecated key; use ankiConnect.ai instead.',
      value: { enabled: true },
      fallback: {},
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
  assert.match(body, /1\. ankiConnect\.openRouter: Deprecated key; use ankiConnect\.ai instead\./);
  assert.match(
    body,
    /2\. ankiConnect\.isLapis\.sentenceCardSentenceField: Deprecated key; sentence-card sentence field is fixed to Sentence\./,
  );
});

test('failStartupFromConfig invokes handlers and throws', () => {
  const calls: string[] = [];
  const previousExitCode = process.exitCode;
  process.exitCode = 0;

  assert.throws(
    () =>
      failStartupFromConfig('Config Error', 'bad value', {
        logError: (details) => {
          calls.push(`log:${details}`);
        },
        showErrorBox: (title, details) => {
          calls.push(`dialog:${title}:${details}`);
        },
        quit: () => {
          calls.push('quit');
        },
      }),
    /bad value/,
  );

  assert.equal(process.exitCode, 1);
  assert.deepEqual(calls, ['log:bad value', 'dialog:Config Error:bad value', 'quit']);

  process.exitCode = previousExitCode;
});
