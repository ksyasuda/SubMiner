import test from 'node:test';
import assert from 'node:assert/strict';
import {
  normalizeX11WindowId,
  parseX11RootActiveWindowId,
  parseX11WindowGeometry,
  parseX11WindowPid,
  X11WindowTracker,
} from './x11-tracker';
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

test('normalizeX11WindowId normalizes decimal and hex ids', () => {
  assert.equal(normalizeX11WindowId('123\n'), '123');
  assert.equal(normalizeX11WindowId('0x7b'), '123');
  assert.equal(normalizeX11WindowId(''), null);
  assert.equal(normalizeX11WindowId('nope'), null);
});

test('parseX11RootActiveWindowId parses root _NET_ACTIVE_WINDOW output', () => {
  assert.equal(parseX11RootActiveWindowId('_NET_ACTIVE_WINDOW(WINDOW): window id # 0x7b'), '123');
  assert.equal(parseX11RootActiveWindowId('_NET_ACTIVE_WINDOW(WINDOW): window id # 0x0'), '0');
  assert.equal(parseX11RootActiveWindowId('_NET_ACTIVE_WINDOW:  not found.'), null);
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

test('X11WindowTracker converts both physical rectangle corners to Electron DIP', async () => {
  const convertedPoints: Array<{ x: number; y: number }> = [];
  const tracker = new X11WindowTracker(
    undefined,
    async (command, args) => {
      if (command === 'xdotool' && args[0] === 'search') {
        return '123';
      }
      if (command === 'xdotool' && args[0] === 'getactivewindow') {
        return '123';
      }
      if (command === 'xwininfo') {
        return `Absolute upper-left X:  2000
Absolute upper-left Y:  125
Width: 1000
Height: 750`;
      }
      return '';
    },
    (point) => {
      convertedPoints.push(point);
      if (point.x === 2000) return { x: 1600, y: 100 };
      return { x: 2400, y: 700 };
    },
  );

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(convertedPoints, [
    { x: 2000, y: 125 },
    { x: 3000, y: 875 },
  ]);
  assert.deepEqual(tracker.getGeometry(), {
    x: 1600,
    y: 100,
    width: 800,
    height: 600,
  });
});

test('X11WindowTracker updates target focus from active X11 window', async () => {
  let activeWindowId = '999';
  const tracker = new X11WindowTracker(undefined, async (command, args) => {
    if (command === 'xdotool' && args[0] === 'search') {
      return '123';
    }
    if (command === 'xprop' && args.join(' ') === '-root _NET_ACTIVE_WINDOW') {
      return `_NET_ACTIVE_WINDOW(WINDOW): window id # ${activeWindowId}`;
    }
    if (command === 'xwininfo') {
      return `Absolute upper-left X:  0
Absolute upper-left Y:  0
Width: 640
Height: 360`;
    }
    return '';
  });

  const focusStates: boolean[] = [];
  tracker.onWindowFocusChange = (focused) => {
    focusStates.push(focused);
  };

  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tracker.isTargetWindowFocused(), false);
  assert.deepEqual(focusStates, []);

  activeWindowId = '123';
  (tracker as unknown as { pollGeometry: () => void }).pollGeometry();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(tracker.isTargetWindowFocused(), true);
  assert.deepEqual(focusStates, [true]);
});

test('X11WindowTracker falls back to xdotool active window when root active window is unavailable', async () => {
  const commands: Array<{ command: string; args: string[] }> = [];
  const tracker = new X11WindowTracker(undefined, async (command, args) => {
    commands.push({ command, args });
    if (command === 'xdotool' && args[0] === 'search') {
      return '123';
    }
    if (command === 'xprop' && args.join(' ') === '-root _NET_ACTIVE_WINDOW') {
      throw new Error('missing root active window');
    }
    if (command === 'xdotool' && args[0] === 'getactivewindow') {
      return '999';
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

  assert.equal(tracker.isTargetWindowFocused(), false);
  assert.ok(
    commands.some(
      (call) => call.command === 'xprop' && call.args.join(' ') === '-root _NET_ACTIVE_WINDOW',
    ),
  );
  assert.ok(
    commands.some((call) => call.command === 'xdotool' && call.args[0] === 'getactivewindow'),
  );
});

test('X11WindowTracker treats a different root active X11 window as mpv unfocused', async () => {
  const socketPath = '/tmp/subminer-mpv.sock';
  const tracker = new X11WindowTracker(socketPath, async (command, args) => {
    if (command === 'xdotool' && args[0] === 'search') {
      return '123';
    }
    if (command === 'xprop' && args.join(' ') === '-root _NET_ACTIVE_WINDOW') {
      return '_NET_ACTIVE_WINDOW(WINDOW): window id # 0x3e7';
    }
    if (command === 'xprop' && args.join(' ') === '-id 123 _NET_WM_PID') {
      return '_NET_WM_PID(CARDINAL) = 4242';
    }
    if (command === 'xprop' && args.join(' ') === '-id 999 _NET_WM_PID') {
      return '_NET_WM_PID(CARDINAL) = 9999';
    }
    if (command === 'ps') {
      return `mpv --input-ipc-server=${socketPath}`;
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

  assert.equal(tracker.isTargetWindowFocused(), false);
  assert.equal(tracker.getTargetWindowMediaSourceId(), 'window:123:0');
  assert.equal(tracker.getTargetWindowNativeId(), '123');
});

test('X11WindowTracker treats active X11 windows with matching PID as focused', async () => {
  const socketPath = '/tmp/subminer-mpv.sock';
  const tracker = new X11WindowTracker(socketPath, async (command, args) => {
    if (command === 'xdotool' && args[0] === 'search') {
      return '123';
    }
    if (command === 'xprop' && args.join(' ') === '-root _NET_ACTIVE_WINDOW') {
      return '_NET_ACTIVE_WINDOW(WINDOW): window id # 0x3e7';
    }
    if (command === 'xprop' && args.join(' ') === '-id 123 _NET_WM_PID') {
      return '_NET_WM_PID(CARDINAL) = 4242';
    }
    if (command === 'xprop' && args.join(' ') === '-id 999 _NET_WM_PID') {
      return '_NET_WM_PID(CARDINAL) = 4242';
    }
    if (command === 'ps') {
      return `mpv --input-ipc-server=${socketPath}`;
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

  assert.equal(tracker.isTargetWindowFocused(), true);
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

test('X11WindowTracker activates and raises the tracked mpv window without changing the target', async () => {
  const commands: Array<{ command: string; args: string[] }> = [];
  const tracker = new X11WindowTracker(undefined, async (command, args) => {
    commands.push({ command, args });
    if (command === 'xdotool' && args[0] === 'search') {
      return '123';
    }
    if (command === 'xdotool' && args[0] === 'windowactivate') {
      return '';
    }
    if (command === 'xdotool' && args[0] === 'windowraise') {
      return '';
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

  const raised = await tracker.raiseTargetWindow();

  assert.equal(raised, true);
  assert.ok(
    commands.some(
      (call) => call.command === 'xdotool' && call.args.join(' ') === 'windowactivate 123',
    ),
  );
  assert.ok(
    commands.some(
      (call) => call.command === 'xdotool' && call.args.join(' ') === 'windowraise 123',
    ),
  );
});

test('X11WindowTracker raises the same target id captured before activation', async () => {
  const commands: Array<{ command: string; args: string[] }> = [];
  const tracker = new X11WindowTracker(undefined, async (command, args) => {
    commands.push({ command, args });
    if (command === 'xdotool' && args[0] === 'windowactivate') {
      (tracker as unknown as { targetWindowId: string | null }).targetWindowId = '456';
      return '';
    }
    return '';
  });
  (tracker as unknown as { targetWindowId: string | null }).targetWindowId = '123';

  const raised = await tracker.raiseTargetWindow();

  assert.equal(raised, true);
  assert.deepEqual(
    commands.filter((call) => call.command === 'xdotool').map((call) => call.args.join(' ')),
    ['windowactivate 123', 'windowraise 123'],
  );
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
