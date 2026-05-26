import test from 'node:test';
import assert from 'node:assert/strict';
import { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './mpv-osd-log';

test('append mpv log writes timestamped message', () => {
  const writes: string[] = [];
  const { appendToMpvLog, flushMpvLog } = createAppendToMpvLogHandler({
    getLogPath: () => '/tmp/subminer/mpv.log',
    dirname: (targetPath: string) => {
      writes.push(`dirname:${targetPath}`);
      return '/tmp/subminer';
    },
    mkdir: async (targetPath: string) => {
      writes.push(`mkdir:${targetPath}`);
    },
    appendFile: async (_targetPath: string, data: string) => {
      writes.push(`append:${data.trimEnd()}`);
    },
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  });

  appendToMpvLog('hello');
  return flushMpvLog().then(() => {
    assert.deepEqual(writes, [
      'dirname:/tmp/subminer/mpv.log',
      'mkdir:/tmp/subminer',
      'append:[2026-02-20T00:00:00.000Z] hello',
    ]);
  });
});

test('append mpv log observes path changes', async () => {
  const writes: string[] = [];
  let logPath = '';
  const { appendToMpvLog, flushMpvLog } = createAppendToMpvLogHandler({
    getLogPath: () => logPath,
    dirname: () => '/tmp/subminer',
    mkdir: async () => {},
    appendFile: async (targetPath: string, data: string) => {
      writes.push(`${targetPath}:${data.trimEnd()}`);
    },
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  });

  appendToMpvLog('disabled');
  logPath = '/tmp/subminer/mpv.log';
  appendToMpvLog('enabled');
  await flushMpvLog();

  assert.deepEqual(writes, ['/tmp/subminer/mpv.log:[2026-02-20T00:00:00.000Z] enabled']);
});

test('append mpv log queues multiple messages and flush waits for pending write', async () => {
  const writes: string[] = [];
  const { appendToMpvLog, flushMpvLog } = createAppendToMpvLogHandler({
    getLogPath: () => '/tmp/subminer/mpv.log',
    dirname: () => '/tmp/subminer',
    mkdir: async () => {
      writes.push('mkdir');
    },
    appendFile: async (_targetPath: string, data: string) => {
      writes.push(`append:${data.trimEnd()}`);
      await new Promise<void>((resolve) => {
        setTimeout(resolve, 25);
      });
    },
    now: (() => {
      const values = ['2026-02-20T00:00:00.000Z', '2026-02-20T00:00:01.000Z'];
      let index = 0;
      return () => {
        const nextValue = values[index] ?? values[values.length - 1] ?? values[0] ?? '';
        index += 1;
        return new Date(nextValue);
      };
    })(),
  });

  appendToMpvLog('first');
  appendToMpvLog('second');

  let flushed = false;
  const flushPromise = flushMpvLog().then(() => {
    flushed = true;
  });

  await Promise.resolve();
  assert.equal(flushed, false);
  await flushPromise;

  const appendedPayload = writes
    .filter((entry) => entry.startsWith('append:'))
    .map((entry) => entry.replace('append:', ''))
    .join('\n');
  assert.ok(appendedPayload.includes('[2026-02-20T00:00:00.000Z] first'));
  assert.ok(appendedPayload.includes('[2026-02-20T00:00:01.000Z] second'));
});

test('append mpv log swallows async filesystem errors', async () => {
  const { appendToMpvLog, flushMpvLog } = createAppendToMpvLogHandler({
    getLogPath: () => '/tmp/subminer/mpv.log',
    dirname: () => '/tmp/subminer',
    mkdir: async () => {
      throw new Error('disk error');
    },
    appendFile: async () => {
      throw new Error('should not reach');
    },
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  });

  assert.doesNotThrow(() => appendToMpvLog('hello'));
  await assert.doesNotReject(async () => flushMpvLog());
});

test('show mpv osd logs marker and forwards fallback logging', () => {
  const calls: string[] = [];
  const client = { connected: false, send: () => {} } as never;
  const showMpvOsd = createShowMpvOsdHandler({
    appendToMpvLog: (message) => calls.push(`append:${message}`),
    showMpvOsdRuntime: (_client, text, fallbackLog) => {
      calls.push(`show:${text}`);
      fallbackLog('fallback-line');
      return false;
    },
    getMpvClient: () => client,
    logInfo: (line) => calls.push(`info:${line}`),
  });

  const shown = showMpvOsd('subtitle copied');
  assert.equal(shown, false);
  assert.deepEqual(calls, [
    'append:[OSD] subtitle copied',
    'show:subtitle copied',
    'info:fallback-line',
  ]);
});
