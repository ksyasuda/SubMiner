import assert from 'node:assert/strict';
import test from 'node:test';
import { createMpvOsdRuntimeHandlers } from './mpv-osd-runtime-handlers';

test('mpv osd runtime handlers compose append and osd logging flow', () => {
  const calls: string[] = [];
  const runtime = createMpvOsdRuntimeHandlers({
    appendToMpvLogMainDeps: {
      logPath: '/tmp/subminer/mpv.log',
      dirname: () => '/tmp/subminer',
      mkdirSync: () => {},
      appendFileSync: (_targetPath, data) => calls.push(`append:${data.trimEnd()}`),
      now: () => new Date('2026-02-20T00:00:00.000Z'),
    },
    buildShowMpvOsdMainDeps: (appendToMpvLog) => ({
      appendToMpvLog,
      showMpvOsdRuntime: (_client, text, fallbackLog) => {
        calls.push(`show:${text}`);
        fallbackLog('fallback');
      },
      getMpvClient: () => null,
      logInfo: (line) => calls.push(`info:${line}`),
    }),
  });

  runtime.showMpvOsd('hello');

  assert.deepEqual(calls, [
    'append:[2026-02-20T00:00:00.000Z] [OSD] hello',
    'show:hello',
    'info:fallback',
  ]);
});
