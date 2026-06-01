import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createLinuxX11CursorPointReader,
  parseXdotoolMouseLocation,
} from './linux-x11-cursor-point';

test('parseXdotoolMouseLocation parses root cursor coordinates', () => {
  assert.deepEqual(
    parseXdotoolMouseLocation(`X=1700
Y=1050
SCREEN=0
WINDOW=44040194
`),
    { x: 1700, y: 1050 },
  );
});

test('createLinuxX11CursorPointReader returns cached X11 cursor point over stale fallback', async () => {
  let now = 1000;
  const pendingCommand: { resolve?: (value: string) => void } = {};
  const calls: Array<{ command: string; args: string[] }> = [];
  const reader = createLinuxX11CursorPointReader({
    env: { DISPLAY: ':1' },
    platform: 'linux',
    now: () => now,
    runCommand: (command, args) => {
      calls.push({ command, args });
      return new Promise((resolve) => {
        pendingCommand.resolve = resolve;
      });
    },
  });

  assert.deepEqual(reader.getCursorScreenPoint({ x: 877, y: 718 }), { x: 877, y: 718 });
  assert.deepEqual(calls, [{ command: 'xdotool', args: ['getmouselocation', '--shell'] }]);

  assert.ok(pendingCommand.resolve);
  pendingCommand.resolve(`X=1700
Y=1050
SCREEN=0
WINDOW=44040194
`);
  await new Promise((resolve) => setImmediate(resolve));

  now += 60;
  assert.deepEqual(reader.getCursorScreenPoint({ x: 877, y: 718 }), { x: 1700, y: 1050 });
});

test('createLinuxX11CursorPointReader does not spawn off X11 Linux', () => {
  const calls: string[] = [];
  const reader = createLinuxX11CursorPointReader({
    env: {},
    platform: 'linux',
    runCommand: async (command) => {
      calls.push(command);
      return '';
    },
  });

  assert.deepEqual(reader.getCursorScreenPoint({ x: 5, y: 6 }), { x: 5, y: 6 });
  assert.deepEqual(calls, []);
});

test('createLinuxX11CursorPointReader does not spawn for supported native Wayland compositors', () => {
  const calls: string[] = [];
  const reader = createLinuxX11CursorPointReader({
    env: {
      DISPLAY: ':1',
      WAYLAND_DISPLAY: 'wayland-0',
      HYPRLAND_INSTANCE_SIGNATURE: 'hypr',
    },
    platform: 'linux',
    runCommand: async (command) => {
      calls.push(command);
      return '';
    },
  });

  assert.deepEqual(reader.getCursorScreenPoint({ x: 7, y: 8 }), { x: 7, y: 8 });
  assert.deepEqual(calls, []);
});
