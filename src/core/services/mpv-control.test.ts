import test from 'node:test';
import assert from 'node:assert/strict';
import {
  playNextSubtitleRuntime,
  replayCurrentSubtitleRuntime,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  showMpvOsdRuntime,
} from './mpv';

test('showMpvOsdRuntime sends show-text when connected', () => {
  const commands: (string | number)[][] = [];
  showMpvOsdRuntime(
    {
      connected: true,
      send: ({ command }) => {
        commands.push(command);
      },
    },
    'hello',
  );
  assert.deepEqual(commands, [['show-text', 'hello', '3000']]);
});

test('showMpvOsdRuntime logs fallback when disconnected', () => {
  const logs: string[] = [];
  showMpvOsdRuntime(
    {
      connected: false,
      send: () => {},
    },
    'hello',
    (line) => {
      logs.push(line);
    },
  );
  assert.deepEqual(logs, ['OSD (MPV not connected): hello']);
});

test('mpv runtime command wrappers call expected client methods', () => {
  const calls: string[] = [];
  const client = {
    connected: true,
    send: ({ command }: { command: (string | number)[] }) => {
      calls.push(`send:${command.join(',')}`);
    },
    replayCurrentSubtitle: () => {
      calls.push('replay');
    },
    playNextSubtitle: () => {
      calls.push('next');
    },
    setSubVisibility: (visible: boolean) => {
      calls.push(`subVisible:${visible}`);
    },
  };

  replayCurrentSubtitleRuntime(client);
  playNextSubtitleRuntime(client);
  sendMpvCommandRuntime(client, ['script-message', 'x']);
  setMpvSubVisibilityRuntime(client, false);

  assert.deepEqual(calls, ['replay', 'next', 'send:script-message,x', 'subVisible:false']);
});
