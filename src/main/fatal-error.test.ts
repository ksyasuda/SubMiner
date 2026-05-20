import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildFatalErrorDetails,
  createFatalErrorReporter,
  resetFatalErrorReporterForTests,
} from './fatal-error';

test('buildFatalErrorDetails includes context, error, and log path', () => {
  const details = buildFatalErrorDetails({
    context: 'Startup failed.',
    error: new Error('boom'),
    logFilePath: 'C:\\Users\\tester\\AppData\\Roaming\\SubMiner\\logs\\app.log',
  });

  assert.match(details, /Startup failed\./);
  assert.match(details, /Error: boom/);
  assert.match(details, /Log file: C:\\Users\\tester\\AppData\\Roaming\\SubMiner\\logs\\app\.log/);
});

test('fatal error reporter writes one log entry and shows one dialog', () => {
  resetFatalErrorReporterForTests();
  const logLines: string[] = [];
  const dialogs: string[] = [];

  const reporter = createFatalErrorReporter({
    appendLogLine: (line) => logLines.push(line),
    showErrorBox: (title, details) => dialogs.push(`${title}:${details}`),
    resolveLogFilePath: () => 'C:\\SubMiner\\logs\\app.log',
    now: () => new Date('2026-05-20T01:02:03.000Z'),
  });

  reporter('first failure', {
    title: 'SubMiner startup failed',
    context: 'SubMiner could not start.',
  });
  reporter('second failure');

  assert.equal(logLines.length, 1);
  assert.match(logLines[0]!, /\[main:fatal\] SubMiner could not start\./);
  assert.match(logLines[0]!, /first failure/);
  assert.equal(dialogs.length, 1);
  assert.match(dialogs[0]!, /^SubMiner startup failed:SubMiner could not start\./);
});
