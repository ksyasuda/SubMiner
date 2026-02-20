import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildAppendToMpvLogMainDepsHandler,
  createBuildShowMpvOsdMainDepsHandler,
} from './mpv-osd-log-main-deps';

test('append to mpv log main deps map filesystem functions and log path', () => {
  const calls: string[] = [];
  const deps = createBuildAppendToMpvLogMainDepsHandler({
    logPath: '/tmp/mpv.log',
    dirname: (targetPath) => {
      calls.push(`dirname:${targetPath}`);
      return '/tmp';
    },
    mkdirSync: (targetPath) => calls.push(`mkdir:${targetPath}`),
    appendFileSync: (_targetPath, data) => calls.push(`append:${data}`),
    now: () => new Date('2026-02-20T00:00:00.000Z'),
  })();

  assert.equal(deps.logPath, '/tmp/mpv.log');
  assert.equal(deps.dirname('/tmp/mpv.log'), '/tmp');
  deps.mkdirSync('/tmp', { recursive: true });
  deps.appendFileSync('/tmp/mpv.log', 'line', { encoding: 'utf8' });
  assert.equal(deps.now().toISOString(), '2026-02-20T00:00:00.000Z');
  assert.deepEqual(calls, ['dirname:/tmp/mpv.log', 'mkdir:/tmp', 'append:line']);
});

test('show mpv osd main deps map runtime delegates and logging callback', () => {
  const calls: string[] = [];
  const client = { id: 'mpv' };
  const deps = createBuildShowMpvOsdMainDepsHandler({
    appendToMpvLog: (message) => calls.push(`append:${message}`),
    showMpvOsdRuntime: (_mpvClient, text, fallbackLog) => {
      calls.push(`show:${text}`);
      fallbackLog('fallback');
    },
    getMpvClient: () => client,
    logInfo: (line) => calls.push(`info:${line}`),
  })();

  assert.deepEqual(deps.getMpvClient(), client);
  deps.appendToMpvLog('hello');
  deps.showMpvOsdRuntime(deps.getMpvClient(), 'subtitle', (line) => deps.logInfo(line));
  assert.deepEqual(calls, ['append:hello', 'show:subtitle', 'info:fallback']);
});
