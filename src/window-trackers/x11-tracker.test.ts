import test from 'node:test';
import assert from 'node:assert/strict';
import { parseX11WindowGeometry, parseX11WindowPid, X11WindowTracker } from './x11-tracker';
import { parseMacOSHelperOutput } from './macos-tracker';

test('parseX11WindowGeometry parses xwininfo output', () => {
  const geometry = parseX11WindowGeometry(`
Absolute upper-left X:  120
Absolute upper-left Y:  240
Width: 1280
Height: 720
`);
  assert.deepEqual(geometry, {
    x: 120,
    y: 240,
    width: 1280,
    height: 720,
  });
});

test('parseX11WindowGeometry preserves negative coordinates', () => {
  const geometry = parseX11WindowGeometry(`
Absolute upper-left X:  -1920
Absolute upper-left Y:  -24
Width: 1920
Height: 1080
`);
  assert.deepEqual(geometry, {
    x: -1920,
    y: -24,
    width: 1920,
    height: 1080,
  });
});

test('parseX11WindowPid parses xprop output', () => {
  assert.equal(parseX11WindowPid('_NET_WM_PID(CARDINAL) = 4242'), 4242);
  assert.equal(parseX11WindowPid('_NET_WM_PID(CARDINAL) = not-a-number'), null);
});

test('X11WindowTracker searches only visible mpv windows', async () => {
  const commands: Array<{ command: string; args: string[] }> = [];
  const tracker = new X11WindowTracker(undefined, async (command, args) => {
    commands.push({ command, args });
    if (command === 'xdotool') {
      return '123';
    }
    if (command === 'xwininfo') {
      return `Absolute upper-left X:  0
Absolute upper-left Y:  0
Width: 640
Height: 360`;
    }
    return '';
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(commands[0], {
    command: 'xdotool',
    args: ['search', '--onlyvisible', '--class', 'mpv'],
  });
});

test('X11WindowTracker skips overlapping polls while one command is in flight', async () => {
  let commandCalls = 0;
  let release: (() => void) | undefined;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });

  const tracker = new X11WindowTracker(undefined, async (command) => {
    commandCalls += 1;
    if (command === 'xdotool') {
      await gate;
      return '123';
    }
    if (command === 'xwininfo') {
      return `Absolute upper-left X:  0
Absolute upper-left Y:  0
Width: 640
Height: 360`;
    }
    return '';
  });

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  assert.equal(commandCalls, 1);

  assert.ok(release);
  release();
  await new Promise((resolve) => setTimeout(resolve, 0));
});

test('parseMacOSHelperOutput parses geometry and focused state', () => {
  assert.deepEqual(parseMacOSHelperOutput('120,240,1280,720,1'), {
    geometry: {
      x: 120,
      y: 240,
      width: 1280,
      height: 720,
    },
    focused: true,
  });
});

test('parseMacOSHelperOutput tolerates unfocused helper output', () => {
  assert.deepEqual(parseMacOSHelperOutput('120,240,1280,720,0'), {
    geometry: {
      x: 120,
      y: 240,
      width: 1280,
      height: 720,
    },
    focused: false,
  });
});
