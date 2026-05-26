import assert from 'node:assert/strict';
import test from 'node:test';
import { createMpvOsdRuntimeHandlers } from './mpv-osd-runtime-handlers';

test('mpv osd runtime handlers compose append and osd logging flow', async () => {
  const calls: string[] = [];
  const runtime = createMpvOsdRuntimeHandlers({
    appendToMpvLogMainDeps: {
      getLogPath: () => '/tmp/subminer/mpv.log',
      dirname: () => '/tmp/subminer',
      mkdir: async () => {},
      appendFile: async (_targetPath: string, data: string) => {
        calls.push(`append:${data.trimEnd()}`);
      },
      now: () => new Date('2026-02-20T00:00:00.000Z'),
    },
    buildShowMpvOsdMainDeps: (appendToMpvLog) => ({
      appendToMpvLog,
      showMpvOsdRuntime: (_client, text, fallbackLog) => {
        calls.push(`show:${text}`);
        fallbackLog('fallback');
        return false;
      },
      getMpvClient: () => null,
      logInfo: (line) => calls.push(`info:${line}`),
    }),
  });

  const shown = runtime.showMpvOsd('hello');
  await runtime.flushMpvLog();

  assert.equal(shown, false);
  assert.deepEqual(calls, [
    'show:hello',
    'info:fallback',
    'append:[2026-02-20T00:00:00.000Z] [OSD] hello',
  ]);
});
