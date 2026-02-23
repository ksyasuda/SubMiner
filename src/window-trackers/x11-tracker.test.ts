import test from 'node:test';
import assert from 'node:assert/strict';
import { parseX11WindowGeometry, parseX11WindowPid, X11WindowTracker } from './x11-tracker';

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

test('parseX11WindowPid parses xprop output', () => {
  assert.equal(parseX11WindowPid('_NET_WM_PID(CARDINAL) = 4242'), 4242);
  assert.equal(parseX11WindowPid('_NET_WM_PID(CARDINAL) = not-a-number'), null);
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
