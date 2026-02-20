import test from 'node:test';
import assert from 'node:assert/strict';
import { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './mpv-osd-log';

test('append mpv log writes timestamped message', () => {
  const calls: string[] = [];
  const appendToMpvLog = createAppendToMpvLogHandler({
    logPath: '/tmp/subminer/mpv.log',
    dirname: (targetPath) => {
      calls.push(`dirname:${targetPath}`);
      return '/tmp/subminer';
    },
    mkdirSync: (targetPath) => {
      calls.push(`mkdir:${targetPath}`);
    },
    appendFileSync: (_targetPath, data) => {
      calls.push(`append:${data.trimEnd()}`);
    },
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  });

  appendToMpvLog('hello');
  assert.deepEqual(calls, [
    'dirname:/tmp/subminer/mpv.log',
    'mkdir:/tmp/subminer',
    'append:[2026-02-20T00:00:00.000Z] hello',
  ]);
});

test('append mpv log swallows filesystem errors', () => {
  const appendToMpvLog = createAppendToMpvLogHandler({
    logPath: '/tmp/subminer/mpv.log',
    dirname: () => '/tmp/subminer',
    mkdirSync: () => {
      throw new Error('disk error');
    },
    appendFileSync: () => {
      throw new Error('should not reach');
    },
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  });

  assert.doesNotThrow(() => appendToMpvLog('hello'));
});

test('show mpv osd logs marker and forwards fallback logging', () => {
  const calls: string[] = [];
  const client = { connected: false, send: () => {} } as never;
  const showMpvOsd = createShowMpvOsdHandler({
    appendToMpvLog: (message) => calls.push(`append:${message}`),
    showMpvOsdRuntime: (_client, text, fallbackLog) => {
      calls.push(`show:${text}`);
      fallbackLog('fallback-line');
    },
    getMpvClient: () => client,
    logInfo: (line) => calls.push(`info:${line}`),
  });

  showMpvOsd('subtitle copied');
  assert.deepEqual(calls, [
    'append:[OSD] subtitle copied',
    'show:subtitle copied',
    'info:fallback-line',
  ]);
});
